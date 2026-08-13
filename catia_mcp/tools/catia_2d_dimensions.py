"""Reliable CATIA V5 2D drafting-dimension MCP tools.

This module deliberately separates four concerns which are often confused by
LLM generated CATIA macros:

1. resolve geometry inside one DrawingView;
2. ask CATIA to create an associative DrawingDimension;
3. calculate and verify the expected measured value;
4. place the dimension in the *view* coordinate system (or convert from sheet
   coordinates) and verify the resulting value boundary box.

Install this file beside the other CATIA MCP tool modules and let the existing
registry call ``register_tools(mcp, ctx)``.

Supported geometry selectors are the 1-based index returned by
``catia_list_2d_drawing_geometry`` or the exact AnyObject.Name.  Indices are
only stable until the view is regenerated; list again after a view update.

CATIA V5 Drafting Automation references used for the implementation:
https://catiadesign.org/_doc/V5Automation/generated/interfaces/DraftingInterfaces/interface_DrawingDimensions_34190.htm
https://catiadesign.org/_doc/V5Automation/generated/interfaces/DraftingInterfaces/interface_DrawingDimension_32350.htm
https://catiadesign.org/_doc/V5Automation/generated/interfaces/DraftingInterfaces/interface_DrawingDimValue_30339.htm
https://catiadesign.org/_doc/V5Automation/generated/interfaces/DraftingInterfaces/interface_DrawingView_25425.htm
"""

from __future__ import annotations

import math
from typing import Any, Optional


IMPLEMENTATION_VERSION = "catia-2d-dimensions-2026-08-12-v1"
_CATVB_SCRIPT_LANGUAGE = 1
_EPS = 1.0e-9

SUPPORTED_KINDS = {
    "distance": "Two points/curves: automatic shortest/true distance.",
    "horizontal_distance": "Horizontal distance between two anchors.",
    "vertical_distance": "Vertical distance between two anchors.",
    "aligned_distance": "True/aligned distance between two anchors.",
    "line_to_line": "Perpendicular distance between parallel Line2D objects.",
    "center_to_line": "Perpendicular distance from a Circle2D/Ellipse2D centre to a Line2D.",
    "center_to_center": "Distance between two Circle2D/Ellipse2D centres.",
    "radius": "Associative radial dimension on a Circle2D.",
    "diameter": "Associative diameter dimension on a Circle2D.",
    "angle": "Angle between two Line2D objects.",
    "line_length": "Length of one bounded Line2D/Curve2D.",
}

LINE_REP = {
    "auto": "catDimAuto",
    "horizontal": "catDimHoriz",
    "vertical": "catDimVert",
    "true": "catDimTrueDim",
}

_DESCRIBE_GEOMETRY_VBS = r'''
Public Function MCP_Describe2DGeometry(g)
    Dim t, n, d(11), p(3), q(1), c(1), a(1)
    Dim okRange, okPoint, okCenter, okLine, okRadius
    t = TypeName(g)
    n = ""
    On Error Resume Next
    n = CStr(g.Name)
    Err.Clear
    g.GetRangeBox p
    okRange = (Err.Number = 0)
    Err.Clear
    g.GetCoordinates q
    okPoint = (Err.Number = 0)
    Err.Clear
    g.GetCenter c
    okCenter = (Err.Number = 0)
    Err.Clear
    g.GetOrigin q
    g.GetDirection a
    okLine = (Err.Number = 0)
    Err.Clear
    d(10) = CDbl(g.Radius)
    okRadius = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    d(0) = CStr(t)
    d(1) = CStr(n)
    d(2) = CBool(okRange)
    If okRange Then
        d(3) = CDbl(p(0)): d(4) = CDbl(p(1))
        d(5) = CDbl(p(2)): d(6) = CDbl(p(3))
    Else
        d(3) = 0#: d(4) = 0#: d(5) = 0#: d(6) = 0#
    End If
    If okPoint Then
        d(7) = "point": d(8) = CDbl(q(0)): d(9) = CDbl(q(1))
    ElseIf okCenter Then
        d(7) = "center": d(8) = CDbl(c(0)): d(9) = CDbl(c(1))
    ElseIf okLine Then
        d(7) = "line": d(8) = CDbl(q(0)): d(9) = CDbl(q(1))
    Else
        d(7) = "unknown": d(8) = 0#: d(9) = 0#
    End If
    If okRadius Then d(10) = CDbl(g.Radius) Else d(10) = 0#
    If okLine Then
        d(11) = CStr(CDbl(a(0))) & "," & CStr(CDbl(a(1)))
    Else
        d(11) = ""
    End If
    MCP_Describe2DGeometry = d
End Function
'''

_CREATE_DIMENSION_VBS = r'''
Public Function MCP_Create2DDimension(viewObj, g1, g2, kindName, repName, _
                                      s1x, s1y, s2x, s2y, hasSecond)
    Dim elems, pts, dimObj, dimType, lineRep
    Select Case LCase(CStr(repName))
    Case "horizontal": lineRep = catDimHoriz
    Case "vertical":   lineRep = catDimVert
    Case "true":       lineRep = catDimTrueDim
    Case Else:          lineRep = catDimAuto
    End Select

    Select Case LCase(CStr(kindName))
    Case "radius"
        dimType = catDimRadiusTangent
        elems = Array(g1): pts = Array(CDbl(s1x), CDbl(s1y))
    Case "diameter"
        dimType = catDimDiameterTangent
        elems = Array(g1): pts = Array(CDbl(s1x), CDbl(s1y))
    Case "angle"
        dimType = catDimAngle
        elems = Array(g1, g2)
        pts = Array(CDbl(s1x), CDbl(s1y), CDbl(s2x), CDbl(s2y))
    Case "line_length"
        dimType = catDimLength
        elems = Array(g1): pts = Array(CDbl(s1x), CDbl(s1y))
    Case Else
        dimType = catDimDistance
        If CBool(hasSecond) Then
            elems = Array(g1, g2)
            pts = Array(CDbl(s1x), CDbl(s1y), CDbl(s2x), CDbl(s2y))
        Else
            elems = Array(g1): pts = Array(CDbl(s1x), CDbl(s1y))
        End If
    End Select
    Set dimObj = viewObj.Dimensions.Add(dimType, elems, pts, lineRep)
    Set MCP_Create2DDimension = dimObj
End Function
'''

_BOUNDARY_VBS = r'''
Public Function MCP_GetDimensionBoundary(dimObj)
    Dim b(7)
    dimObj.GetBoundaryBox b
    MCP_GetDimensionBoundary = b
End Function
'''

_TOLERANCE_VBS = r'''
Public Function MCP_SetNumericTolerance(dimObj, upTol, lowTol, displayMode)
    dimObj.SetTolerances 1, "", CStr(upTol), CStr(lowTol), _
                             CDbl(upTol), CDbl(lowTol), CLng(displayMode)
    MCP_SetNumericTolerance = True
End Function
'''


def _format_error(exc: BaseException) -> str:
    info = getattr(exc, "excepinfo", None)
    if info and len(info) > 2 and info[2]:
        return str(info[2])
    return str(exc)


def _success(data: Any, warnings: Optional[list[str]] = None) -> dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": True,
        "status": "success_with_warnings" if warning_list else "success",
        "data": data,
        "warnings": warning_list,
    }


def _error(message: str, data: Any = None) -> dict[str, Any]:
    result = {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": False,
        "status": "error",
        "error": str(message),
        "warnings": [],
    }
    if data is not None:
        result["data"] = data
    return result


def _finite(value: Any, name: str) -> float:
    if isinstance(value, bool):
        raise ValueError(f"{name} must be a finite number")
    number = float(value)
    if not math.isfinite(number):
        raise ValueError(f"{name} must be finite")
    return number


def _evaluate(application: Any, script: str, function: str, args: list[Any]) -> Any:
    try:
        return application.SystemService.Evaluate(
            script, _CATVB_SCRIPT_LANGUAGE, function, args
        )
    except Exception as exc:
        raise RuntimeError(
            f"CATIA SystemService.Evaluate/{function} failed: {_format_error(exc)}"
        ) from exc


def _active_context(conn: Any, view_name: str = "") -> tuple[Any, Any, Any, Any]:
    application = conn.connect(visible=True)
    document = conn.get_active_drawing_document()
    sheet = document.Sheets.ActiveSheet
    views = sheet.Views
    if str(view_name).strip():
        view = views.Item(str(view_name).strip())
    else:
        view = views.ActiveView
    # Background/Main views do not contain the intended generated geometry.
    try:
        view_type = int(view.ViewType)
    except Exception:
        view_type = -1
    if view_type == 0:
        raise ValueError(
            "The active view is Background View; activate a Front/Top/Right/detail view "
            "or pass view_name explicitly."
        )
    return application, document, sheet, view


def _geometry(view: Any, selector: int | str) -> tuple[Any, int]:
    elements = view.GeometricElements
    if isinstance(selector, bool):
        raise ValueError("geometry selector cannot be boolean")
    if isinstance(selector, int):
        if selector < 1 or selector > int(elements.Count):
            raise IndexError(
                f"geometry index {selector} is outside 1..{int(elements.Count)}"
            )
        return elements.Item(selector), selector
    name = str(selector).strip()
    if not name:
        raise ValueError("geometry selector cannot be empty")
    obj = elements.Item(name)
    for index in range(1, int(elements.Count) + 1):
        try:
            if elements.Item(index).Name == obj.Name:
                return obj, index
        except Exception:
            pass
    return obj, -1


def _describe(application: Any, obj: Any, index: int) -> dict[str, Any]:
    values = list(
        _evaluate(
            application,
            _DESCRIBE_GEOMETRY_VBS,
            "MCP_Describe2DGeometry",
            [obj],
        )
    )
    if len(values) != 12:
        raise RuntimeError(f"geometry description returned {len(values)} fields")
    direction = None
    if str(values[11]).strip():
        dx, dy = str(values[11]).replace(";", ",").split(",", 1)
        direction = [float(dx), float(dy)]
    result = {
        "index": index,
        "name": str(values[1]),
        "automation_type": str(values[0]),
        "has_range_box": bool(values[2]),
        "range_box_view_mm": (
            {
                "xmin": float(values[3]),
                "ymin": float(values[4]),
                "xmax": float(values[5]),
                "ymax": float(values[6]),
            }
            if bool(values[2])
            else None
        ),
        "location_kind": str(values[7]),
        "location_view_mm": [float(values[8]), float(values[9])],
        "radius_mm": float(values[10]) if float(values[10]) > 0 else None,
        "direction": direction,
    }
    return result


def _type_has(item: dict[str, Any], token: str) -> bool:
    return token.lower() in item["automation_type"].lower()


def _center(item: dict[str, Any]) -> tuple[float, float]:
    if item["location_kind"] != "center":
        raise ValueError(
            f"geometry {item['index']} ({item['automation_type']}) has no accessible centre"
        )
    return tuple(item["location_view_mm"])  # type: ignore[return-value]


def _point(item: dict[str, Any]) -> tuple[float, float]:
    if item["location_kind"] in {"point", "center", "line"}:
        return tuple(item["location_view_mm"])  # type: ignore[return-value]
    box = item.get("range_box_view_mm")
    if box:
        return ((box["xmin"] + box["xmax"]) / 2, (box["ymin"] + box["ymax"]) / 2)
    raise ValueError(f"geometry {item['index']} has no usable anchor")


def _line(item: dict[str, Any]) -> tuple[float, float, float, float]:
    if item["location_kind"] != "line" or not item["direction"]:
        raise ValueError(
            f"geometry {item['index']} ({item['automation_type']}) is not a Line2D"
        )
    x, y = item["location_view_mm"]
    dx, dy = item["direction"]
    length = math.hypot(dx, dy)
    if length <= _EPS:
        raise ValueError(f"geometry {item['index']} has a zero direction vector")
    return x, y, dx / length, dy / length


def _range_endpoints(item: dict[str, Any]) -> tuple[tuple[float, float], tuple[float, float]]:
    box = item.get("range_box_view_mm")
    if not box:
        raise ValueError(f"geometry {item['index']} has no range box")
    if item["direction"]:
        dx, dy = item["direction"]
        candidates = [
            (box["xmin"], box["ymin"]), (box["xmin"], box["ymax"]),
            (box["xmax"], box["ymin"]), (box["xmax"], box["ymax"]),
        ]
        projected = sorted(candidates, key=lambda p: p[0] * dx + p[1] * dy)
        return projected[0], projected[-1]
    return (box["xmin"], box["ymin"]), (box["xmax"], box["ymax"])


def _validate_kind(kind: str, a: dict[str, Any], b: Optional[dict[str, Any]]) -> None:
    if kind not in SUPPORTED_KINDS:
        raise ValueError(f"kind must be one of: {', '.join(SUPPORTED_KINDS)}")
    two = kind not in {"radius", "diameter", "line_length"}
    if two and b is None:
        raise ValueError(f"{kind} requires element2")
    if kind in {"radius", "diameter"}:
        if not _type_has(a, "Circle2D") or a["radius_mm"] is None:
            raise ValueError(f"{kind} requires a Circle2D as element1")
    if kind in {"line_to_line", "angle"}:
        _line(a)
        _line(b or {})
    if kind == "center_to_line":
        _center(a)
        _line(b or {})
    if kind == "center_to_center":
        _center(a)
        _center(b or {})


def _catia_geometry_objects(
    kind: str, obj1: Any, obj2: Optional[Any]
) -> tuple[Any, Optional[Any]]:
    # Centre dimensions must attach to the actual circle/ellipse CenterPoint,
    # not to a newly-created duplicate point.  This preserves associativity.
    if kind == "center_to_line":
        return obj1.CenterPoint, obj2
    if kind == "center_to_center":
        return obj1.CenterPoint, obj2.CenterPoint
    return obj1, obj2


def _expected_and_anchors(
    kind: str, a: dict[str, Any], b: Optional[dict[str, Any]]
) -> tuple[Optional[float], tuple[float, float], tuple[float, float], str]:
    if kind in {"radius", "diameter"}:
        c = _center(a)
        r = float(a["radius_mm"])
        return (r if kind == "radius" else 2 * r), (c[0] + r, c[1]), c, "length_mm"
    if kind == "line_length":
        p1, p2 = _range_endpoints(a)
        return math.dist(p1, p2), p1, p2, "length_mm"
    assert b is not None
    if kind == "center_to_line":
        p = _center(a)
        x, y, dx, dy = _line(b)
        signed = (p[0] - x) * (-dy) + (p[1] - y) * dx
        q = (p[0] - signed * (-dy), p[1] - signed * dx)
        return abs(signed), p, q, "length_mm"
    if kind == "center_to_center":
        p, q = _center(a), _center(b)
        return math.dist(p, q), p, q, "length_mm"
    if kind == "line_to_line":
        x1, y1, dx1, dy1 = _line(a)
        x2, y2, dx2, dy2 = _line(b)
        cross = dx1 * dy2 - dy1 * dx2
        if abs(cross) > 1.0e-5:
            raise ValueError("line_to_line requires parallel lines; use kind='angle'")
        signed = (x2 - x1) * (-dy1) + (y2 - y1) * dx1
        p = (x1, y1)
        q = (x1 + signed * (-dy1), y1 + signed * dx1)
        return abs(signed), p, q, "length_mm"
    if kind == "angle":
        x1, y1, dx1, dy1 = _line(a)
        x2, y2, dx2, dy2 = _line(b)
        angle = math.acos(max(-1.0, min(1.0, abs(dx1 * dx2 + dy1 * dy2))))
        return angle, (x1, y1), (x2, y2), "angle_radians"
    p, q = _point(a), _point(b)
    if kind == "horizontal_distance":
        value = abs(q[0] - p[0])
    elif kind == "vertical_distance":
        value = abs(q[1] - p[1])
    else:
        value = math.dist(p, q)
    return value, p, q, "length_mm"


def _view_transform(view: Any) -> dict[str, float]:
    scale = float(getattr(view, "Scale2", getattr(view, "Scale", 1.0)))
    if abs(scale) <= _EPS:
        raise ValueError("view scale is zero")
    return {
        "origin_sheet_x_mm": float(view.xAxisData),
        "origin_sheet_y_mm": float(view.yAxisData),
        "angle_radians": float(view.Angle),
        "scale": scale,
    }


def _sheet_to_view(x: float, y: float, tf: dict[str, float]) -> tuple[float, float]:
    dx = x - tf["origin_sheet_x_mm"]
    dy = y - tf["origin_sheet_y_mm"]
    c, s = math.cos(tf["angle_radians"]), math.sin(tf["angle_radians"])
    return ((c * dx + s * dy) / tf["scale"], (-s * dx + c * dy) / tf["scale"])


def _view_to_sheet(x: float, y: float, tf: dict[str, float]) -> tuple[float, float]:
    c, s = math.cos(tf["angle_radians"]), math.sin(tf["angle_radians"])
    sx = tf["origin_sheet_x_mm"] + tf["scale"] * (c * x - s * y)
    sy = tf["origin_sheet_y_mm"] + tf["scale"] * (s * x + c * y)
    return sx, sy


def _auto_position(
    kind: str,
    p: tuple[float, float],
    q: tuple[float, float],
    offset_sheet_mm: float,
    scale: float,
) -> tuple[float, float]:
    mx, my = (p[0] + q[0]) / 2, (p[1] + q[1]) / 2
    offset = offset_sheet_mm / scale
    if kind == "horizontal_distance":
        return mx, max(p[1], q[1]) + offset
    if kind == "vertical_distance":
        return max(p[0], q[0]) + offset, my
    dx, dy = q[0] - p[0], q[1] - p[1]
    length = math.hypot(dx, dy)
    if length <= _EPS:
        return mx + offset, my + offset
    nx, ny = -dy / length, dx / length
    return mx + nx * offset, my + ny * offset


def _dimension_value(dim: Any) -> Optional[float]:
    try:
        return float(dim.GetValue().Value)
    except Exception:
        try:
            return float(dim.Parameters.Item("Measured length").Value)
        except Exception:
            return None


def _boundary(application: Any, dim: Any) -> Optional[dict[str, Any]]:
    try:
        points = [float(x) for x in _evaluate(
            application, _BOUNDARY_VBS, "MCP_GetDimensionBoundary", [dim]
        )]
        if len(points) != 8 or not all(math.isfinite(x) for x in points):
            return None
        xs, ys = points[0::2], points[1::2]
        return {
            "corners_view_mm": [[points[i], points[i + 1]] for i in range(0, 8, 2)],
            "xmin": min(xs), "xmax": max(xs), "ymin": min(ys), "ymax": max(ys),
        }
    except Exception:
        return None


def _update(document: Any, sheet: Any, view: Any) -> list[str]:
    warnings: list[str] = []
    try:
        view.SaveEdition()
    except Exception as exc:
        warnings.append(f"view.SaveEdition failed: {_format_error(exc)}")
    for label, target, methods in (
        ("document", document, ("Update",)),
        ("sheet", sheet, ("ForceUpdate", "Update")),
    ):
        errors = []
        for method in methods:
            try:
                getattr(target, method)()
                break
            except Exception as exc:
                errors.append(f"{method}: {_format_error(exc)}")
        else:
            warnings.append(f"{label} update failed: {errors}")
    return warnings


def _set_tolerance(
    application: Any,
    dim: Any,
    upper: Optional[float],
    lower: Optional[float],
    display_mode: int,
) -> bool:
    if upper is None and lower is None:
        return False
    up = _finite(0.0 if upper is None else upper, "tolerance_upper")
    low = _finite(-up if lower is None else lower, "tolerance_lower")
    _evaluate(
        application,
        _TOLERANCE_VBS,
        "MCP_SetNumericTolerance",
        [dim, up, low, int(display_mode)],
    )
    return True


def _add_internal(
    *,
    application: Any,
    document: Any,
    sheet: Any,
    view: Any,
    kind: str,
    element1: int | str,
    element2: int | str | None,
    position_x: Optional[float],
    position_y: Optional[float],
    position_space: str,
    offset_mm: float,
    witness_points: Optional[list[float]],
    name: str,
    tolerance_upper: Optional[float],
    tolerance_lower: Optional[float],
    tolerance_display_mode: int,
) -> tuple[dict[str, Any], list[str]]:
    kind = str(kind).strip().lower()
    obj1, idx1 = _geometry(view, element1)
    a = _describe(application, obj1, idx1)
    obj2 = None
    b = None
    if element2 is not None:
        obj2, idx2 = _geometry(view, element2)
        b = _describe(application, obj2, idx2)
    _validate_kind(kind, a, b)
    expected, anchor1, anchor2, unit = _expected_and_anchors(kind, a, b)
    cat_obj1, cat_obj2 = _catia_geometry_objects(kind, obj1, obj2)

    rep = {
        "horizontal_distance": "horizontal",
        "vertical_distance": "vertical",
        "aligned_distance": "true",
    }.get(kind, "auto")
    if witness_points is None:
        witnesses = [anchor1[0], anchor1[1], anchor2[0], anchor2[1]]
    else:
        if len(witness_points) != 4:
            raise ValueError("witness_points must be [x1,y1,x2,y2] in view coordinates")
        witnesses = [_finite(v, f"witness_points[{i}]") for i, v in enumerate(witness_points)]

    before = int(view.Dimensions.Count)
    dim = _evaluate(
        application,
        _CREATE_DIMENSION_VBS,
        "MCP_Create2DDimension",
        [
            view, cat_obj1, cat_obj2, kind, rep,
            witnesses[0], witnesses[1], witnesses[2], witnesses[3],
            cat_obj2 is not None,
        ],
    )
    try:
        if str(name).strip():
            dim.Name = str(name).strip()
        created_index = int(view.Dimensions.Count)
        if created_index != before + 1:
            raise RuntimeError(
                f"dimension count changed from {before} to {created_index}; expected +1"
            )
    except Exception:
        # Do not leave an unreported annotation behind when creation verification
        # or naming fails.
        try:
            if int(view.Dimensions.Count) > before:
                view.Dimensions.Remove(int(view.Dimensions.Count))
        except Exception:
            pass
        raise

    tf = _view_transform(view)
    offset = _finite(offset_mm, "offset_mm")
    if position_x is None and position_y is None:
        vx, vy = _auto_position(kind, anchor1, anchor2, offset, tf["scale"])
        resolved_from = "automatic_feature_offset"
    elif position_x is None or position_y is None:
        view.Dimensions.Remove(created_index)
        raise ValueError("position_x and position_y must be supplied together")
    else:
        px = _finite(position_x, "position_x")
        py = _finite(position_y, "position_y")
        space = str(position_space).strip().lower()
        if space == "sheet":
            vx, vy = _sheet_to_view(px, py, tf)
            resolved_from = "sheet_coordinates_converted_to_view"
        elif space == "view":
            vx, vy = px, py
            resolved_from = "explicit_view_coordinates"
        else:
            view.Dimensions.Remove(created_index)
            raise ValueError("position_space must be 'view' or 'sheet'")

    warnings: list[str] = []
    try:
        try:
            dim.ValueAutoMode = False
        except Exception as exc:
            warnings.append(f"could not disable automatic value placement: {_format_error(exc)}")
        dim.MoveValue(float(vx), float(vy), 0, 0)
        tolerance_set = _set_tolerance(
            application, dim, tolerance_upper, tolerance_lower, tolerance_display_mode
        )
        warnings.extend(_update(document, sheet, view))
        # CATIA updates can recompute the annotation. Re-apply and verify the requested
        # position after the update; this is intentional, not a duplicate move.
        dim.MoveValue(float(vx), float(vy), 0, 0)
        try:
            view.SaveEdition()
        except Exception:
            pass
    except Exception as exc:
        cleanup_error = None
        try:
            view.Dimensions.Remove(created_index)
        except Exception as cleanup_exc:
            cleanup_error = _format_error(cleanup_exc)
        message = f"dimension post-processing failed: {_format_error(exc)}"
        if cleanup_error:
            message += f"; rollback also failed: {cleanup_error}"
        raise RuntimeError(message) from exc

    actual = _dimension_value(dim)
    value_matches = None
    if expected is not None and actual is not None:
        tolerance = max(1.0e-6, abs(expected) * 1.0e-6)
        value_matches = abs(actual - expected) <= tolerance
        if not value_matches:
            warnings.append(
                f"CATIA value {actual} does not match independently calculated "
                f"{expected} ({unit}); inspect the selected geometry/witness points."
            )
    box = _boundary(application, dim)
    if box is None:
        warnings.append("CATIA did not return a verifiable dimension value boundary box")
    sx, sy = _view_to_sheet(vx, vy, tf)
    data = {
        "view": str(getattr(view, "Name", "")),
        "dimension_index": created_index,
        "dimension_name": str(getattr(dim, "Name", "")),
        "kind": kind,
        "catia_dim_type": int(getattr(dim, "DimType", -1)),
        "catia_line_rep": int(getattr(dim.GetDimLine(), "DimLineRep", -1)),
        "element1": a,
        "element2": b,
        "associative_geometry": True,
        "expected_value": expected,
        "catia_measured_value": actual,
        "value_unit": unit,
        "value_matches_independent_calculation": value_matches,
        "witness_points_view_mm": witnesses,
        "placement": {
            "resolved_from": resolved_from,
            "view_mm": [vx, vy],
            "sheet_mm": [sx, sy],
            "view_transform": tf,
            "value_boundary_box_view_mm": box,
            "verified": box is not None,
        },
        "tolerance_set": tolerance_set,
        "dimension_count_before": before,
        "dimension_count_after": int(view.Dimensions.Count),
        "document_save_required": True,
    }
    return data, warnings


def _dimension_by_selector(view: Any, selector: int | str) -> tuple[Any, int]:
    dims = view.Dimensions
    dim = dims.Item(selector)
    if isinstance(selector, int):
        return dim, selector
    for i in range(1, int(dims.Count) + 1):
        try:
            if dims.Item(i).Name == dim.Name:
                return dim, i
        except Exception:
            pass
    return dim, -1


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    """Register all tools and return their registry names."""
    conn = ctx.conn
    names: list[str] = []

    @mcp.tool()
    def catia_list_2d_drawing_geometry(
        view_name: str = "", max_items: int = 500
    ) -> dict[str, Any]:
        """List selectable 2D geometry in a drawing view with local coordinates.

        Always call this before creating dimensions. Use the returned 1-based
        indices or exact names; call it again after regenerating the view.
        """
        try:
            application, document, sheet, view = _active_context(conn, view_name)
            count = int(view.GeometricElements.Count)
            limit = max(1, min(int(max_items), count))
            items = []
            failures = []
            for index in range(1, limit + 1):
                try:
                    items.append(_describe(application, view.GeometricElements.Item(index), index))
                except Exception as exc:
                    failures.append({"index": index, "error": _format_error(exc)})
            data = {
                "view": str(view.Name),
                "view_transform": _view_transform(view),
                "geometry_count": count,
                "returned_count": len(items),
                "truncated": limit < count,
                "geometry": items,
                "description_failures": failures,
                "supported_dimension_kinds": SUPPORTED_KINDS,
            }
            warnings = []
            if failures:
                warnings.append(f"{len(failures)} geometry objects could not be described")
            return _success(data, warnings)
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_list_2d_drawing_geometry")

    @mcp.tool()
    def catia_add_2d_drawing_dimension(
        kind: str,
        element1: int | str,
        element2: int | str | None = None,
        view_name: str = "",
        position_x: float | None = None,
        position_y: float | None = None,
        position_space: str = "view",
        offset_mm: float = 15.0,
        witness_points: list[float] | None = None,
        name: str = "",
        tolerance_upper: float | None = None,
        tolerance_lower: float | None = None,
        tolerance_display_mode: int = 0,
    ) -> dict[str, Any]:
        """Create, place and verify one associative 2D drawing dimension.

        kind supports distance/horizontal_distance/vertical_distance/
        aligned_distance/line_to_line/center_to_line/center_to_center/radius/
        diameter/angle/line_length. Coordinates are view-local by default.
        With position_space='sheet', coordinates are rigorously converted using
        view origin, rotation and scale. If position is omitted, a feature-based
        offset in paper millimetres is used.
        """
        try:
            application, document, sheet, view = _active_context(conn, view_name)
            data, warnings = _add_internal(
                application=application, document=document, sheet=sheet, view=view,
                kind=kind, element1=element1, element2=element2,
                position_x=position_x, position_y=position_y,
                position_space=position_space, offset_mm=offset_mm,
                witness_points=witness_points, name=name,
                tolerance_upper=tolerance_upper, tolerance_lower=tolerance_lower,
                tolerance_display_mode=tolerance_display_mode,
            )
            return _success(data, warnings)
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_add_2d_drawing_dimension")

    @mcp.tool()
    def catia_add_2d_drawing_dimensions_batch(
        dimensions: list[dict[str, Any]], view_name: str = "", stop_on_error: bool = True
    ) -> dict[str, Any]:
        """Create multiple verified dimensions from specification dictionaries.

        Each dictionary accepts the same fields as catia_add_2d_drawing_dimension
        except view_name. Existing successful dimensions are not rolled back when
        a later item fails; the result reports exact partial success.
        """
        try:
            if not dimensions:
                raise ValueError("dimensions cannot be empty")
            application, document, sheet, view = _active_context(conn, view_name)
            results, failures, warnings = [], [], []
            for index, spec in enumerate(dimensions):
                try:
                    data, item_warnings = _add_internal(
                        application=application, document=document, sheet=sheet, view=view,
                        kind=spec["kind"], element1=spec["element1"],
                        element2=spec.get("element2"),
                        position_x=spec.get("position_x"), position_y=spec.get("position_y"),
                        position_space=spec.get("position_space", "view"),
                        offset_mm=spec.get("offset_mm", 15.0),
                        witness_points=spec.get("witness_points"), name=spec.get("name", ""),
                        tolerance_upper=spec.get("tolerance_upper"),
                        tolerance_lower=spec.get("tolerance_lower"),
                        tolerance_display_mode=spec.get("tolerance_display_mode", 0),
                    )
                    results.append({"request_index": index, "data": data, "warnings": item_warnings})
                    warnings.extend(f"item {index}: {w}" for w in item_warnings)
                except Exception as exc:
                    failures.append({"request_index": index, "error": _format_error(exc), "spec": spec})
                    if stop_on_error:
                        break
            payload = {
                "view": str(view.Name), "requested_count": len(dimensions),
                "created_count": len(results), "failure_count": len(failures),
                "results": results, "failures": failures,
                "partial_success": bool(results and failures),
                "document_save_required": bool(results),
            }
            if failures:
                warnings.append(f"{len(failures)} batch item(s) failed")
            return _success(payload, warnings)
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_add_2d_drawing_dimensions_batch")

    @mcp.tool()
    def catia_list_2d_drawing_dimensions(view_name: str = "") -> dict[str, Any]:
        """List dimensions with measured values, types, placement and status."""
        try:
            application, document, sheet, view = _active_context(conn, view_name)
            items, warnings = [], []
            for i in range(1, int(view.Dimensions.Count) + 1):
                dim = view.Dimensions.Item(i)
                box = _boundary(application, dim)
                items.append({
                    "index": i, "name": str(getattr(dim, "Name", "")),
                    "dim_type": int(getattr(dim, "DimType", -1)),
                    "dim_status": int(getattr(dim, "DimStatus", -1)),
                    "measured_value": _dimension_value(dim),
                    "value_boundary_box_view_mm": box,
                    "extension_line_count": int(getattr(dim, "NbExtLine", 0)),
                    "symbol_count": int(getattr(dim, "NbSymb", 0)),
                })
                if box is None:
                    warnings.append(f"dimension {i}: boundary box unavailable")
            return _success({"view": str(view.Name), "count": len(items), "dimensions": items}, warnings)
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_list_2d_drawing_dimensions")

    @mcp.tool()
    def catia_edit_2d_drawing_dimension(
        dimension: int | str,
        view_name: str = "",
        position_x: float | None = None,
        position_y: float | None = None,
        position_space: str = "view",
        tolerance_upper: float | None = None,
        tolerance_lower: float | None = None,
        tolerance_display_mode: int = 0,
        prefix: str | None = None,
        suffix: str | None = None,
        symbols_side: int | None = None,
        restore_automatic_position: bool = False,
    ) -> dict[str, Any]:
        """Move/style an existing dimension and set numeric tolerances/text.

        This edits presentation only; it never replaces CATIA's measured value
        with a fake value. prefix/suffix use DrawingDimValue.SetPSText.
        """
        try:
            application, document, sheet, view = _active_context(conn, view_name)
            dim, index = _dimension_by_selector(view, dimension)
            tf = _view_transform(view)
            actions = []
            if restore_automatic_position:
                dim.RestoreValuePosition()
                actions.append("restore_automatic_position")
            if position_x is not None or position_y is not None:
                if position_x is None or position_y is None:
                    raise ValueError("position_x and position_y must be supplied together")
                x, y = _finite(position_x, "position_x"), _finite(position_y, "position_y")
                if str(position_space).lower() == "sheet":
                    x, y = _sheet_to_view(x, y, tf)
                elif str(position_space).lower() != "view":
                    raise ValueError("position_space must be 'view' or 'sheet'")
                try:
                    dim.ValueAutoMode = False
                except Exception:
                    pass
                dim.MoveValue(x, y, 0, 0)
                actions.append("move_value_and_dimension_line")
            if tolerance_upper is not None or tolerance_lower is not None:
                _set_tolerance(application, dim, tolerance_upper, tolerance_lower, tolerance_display_mode)
                actions.append("set_numeric_tolerance")
            if prefix is not None or suffix is not None:
                value = dim.GetValue()
                value.SetPSText(1, str(prefix or ""), str(suffix or ""))
                actions.append("set_prefix_suffix")
            if symbols_side is not None:
                side = int(symbols_side)
                if side not in range(0, 5):
                    raise ValueError("symbols_side must be 0..4")
                dim.SymbolsSide = side
                actions.append("set_symbols_side")
            warnings = _update(document, sheet, view)
            return _success({
                "view": str(view.Name), "dimension_index": index,
                "dimension_name": str(getattr(dim, "Name", "")), "actions": actions,
                "measured_value": _dimension_value(dim),
                "value_boundary_box_view_mm": _boundary(application, dim),
                "document_save_required": bool(actions),
            }, warnings)
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_edit_2d_drawing_dimension")

    @mcp.tool()
    def catia_remove_2d_drawing_dimension(
        dimension: int | str, view_name: str = ""
    ) -> dict[str, Any]:
        """Remove one dimension by 1-based index or exact name and verify count."""
        try:
            application, document, sheet, view = _active_context(conn, view_name)
            dim, index = _dimension_by_selector(view, dimension)
            old_name = str(getattr(dim, "Name", ""))
            before = int(view.Dimensions.Count)
            view.Dimensions.Remove(dimension)
            warnings = _update(document, sheet, view)
            after = int(view.Dimensions.Count)
            if after != before - 1:
                raise RuntimeError(f"dimension count changed {before}->{after}, expected -1")
            return _success({
                "view": str(view.Name), "removed_index": index, "removed_name": old_name,
                "dimension_count_before": before, "dimension_count_after": after,
                "document_save_required": True,
            }, warnings)
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_remove_2d_drawing_dimension")
    return names
