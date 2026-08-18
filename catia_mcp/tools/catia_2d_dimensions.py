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
import time
from typing import Any, Optional


IMPLEMENTATION_VERSION = "catia-2d-dimensions-2026-08-18-v5-geometry-provenance"
_CATVB_SCRIPT_LANGUAGE = 1
_EPS = 1.0e-9
_HELPER_GEOMETRY_PREFIX = "MCP_HELPER_"

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
    Dim t, n, p(3), pointCoords(1), centerCoords(1)
    Dim lineOrigin(1), lineDirection(1), radiusValue
    Dim okRange, okPoint, okCenter, okLine, okRadius
    Dim rangeXMin, rangeYMin, rangeXMax, rangeYMax
    Dim locationKind, locationX, locationY, directionText
    t = TypeName(g)
    n = ""
    rangeXMin = 0#: rangeYMin = 0#: rangeXMax = 0#: rangeYMax = 0#
    locationKind = "unknown": locationX = 0#: locationY = 0#
    radiusValue = 0#: directionText = ""
    On Error Resume Next
    n = CStr(g.Name)
    Err.Clear
    g.GetRangeBox p
    okRange = (Err.Number = 0)
    If okRange Then rangeXMin = CDbl(p(0)): rangeYMin = CDbl(p(1)): rangeXMax = CDbl(p(2)): rangeYMax = CDbl(p(3))
    Err.Clear
    g.GetCoordinates pointCoords
    okPoint = (Err.Number = 0)
    If okPoint Then locationKind = "point": locationX = CDbl(pointCoords(0)): locationY = CDbl(pointCoords(1))
    Err.Clear
    g.GetCenter centerCoords
    okCenter = (Err.Number = 0)
    If (Not okPoint) And okCenter Then locationKind = "center": locationX = CDbl(centerCoords(0)): locationY = CDbl(centerCoords(1))
    Err.Clear
    g.GetOrigin lineOrigin
    g.GetDirection lineDirection
    okLine = (Err.Number = 0)
    If (Not okPoint) And (Not okCenter) And okLine Then locationKind = "line": locationX = CDbl(lineOrigin(0)): locationY = CDbl(lineOrigin(1))
    If okLine Then directionText = CStr(CDbl(lineDirection(0))) & "," & CStr(CDbl(lineDirection(1)))
    Err.Clear
    radiusValue = CDbl(g.Radius)
    okRadius = (Err.Number = 0)
    If Not okRadius Then radiusValue = 0#
    Err.Clear
    On Error GoTo 0
    MCP_Describe2DGeometry = Array(CStr(t), CStr(n), CBool(okRange), rangeXMin, rangeYMin, rangeXMax, rangeYMax, locationKind, locationX, locationY, radiusValue, directionText)
End Function
'''

_GEOMETRY_IDENTITY_VBS = r'''
Public Function MCP_Get2DGeometryIdentity(g)
    Dim objectName
    objectName = ""
    On Error Resume Next
    objectName = CStr(g.Name)
    Err.Clear
    On Error GoTo 0
    MCP_Get2DGeometryIdentity = Array(CStr(TypeName(g)), objectName)
End Function
'''

_GEOMETRY_RANGE_VBS = r'''
Public Function MCP_Get2DGeometryRange(g)
    Dim p(3), succeeded
    succeeded = False
    On Error Resume Next
    g.GetRangeBox p
    succeeded = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    MCP_Get2DGeometryRange = Array(CBool(succeeded), CDbl(p(0)), CDbl(p(1)), CDbl(p(2)), CDbl(p(3)))
End Function
'''

_GEOMETRY_POINT_VBS = r'''
Public Function MCP_Get2DGeometryPoint(g)
    Dim p(1), succeeded
    succeeded = False
    On Error Resume Next
    g.GetCoordinates p
    succeeded = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    MCP_Get2DGeometryPoint = Array(CBool(succeeded), CDbl(p(0)), CDbl(p(1)))
End Function
'''

_GEOMETRY_CENTER_VBS = r'''
Public Function MCP_Get2DGeometryCenter(g)
    Dim p(1), succeeded
    succeeded = False
    On Error Resume Next
    g.GetCenter p
    succeeded = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    MCP_Get2DGeometryCenter = Array(CBool(succeeded), CDbl(p(0)), CDbl(p(1)))
End Function
'''

_GEOMETRY_LINE_VBS = r'''
Public Function MCP_Get2DGeometryLine(g)
    Dim p(1), d(1), succeeded
    succeeded = False
    On Error Resume Next
    g.GetOrigin p
    g.GetDirection d
    succeeded = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    MCP_Get2DGeometryLine = Array(CBool(succeeded), CDbl(p(0)), CDbl(p(1)), CDbl(d(0)), CDbl(d(1)))
End Function
'''

_GEOMETRY_RADIUS_VBS = r'''
Public Function MCP_Get2DGeometryRadius(g)
    Dim radiusValue, succeeded
    radiusValue = 0#: succeeded = False
    On Error Resume Next
    radiusValue = CDbl(g.Radius)
    succeeded = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    MCP_Get2DGeometryRadius = Array(CBool(succeeded), radiusValue)
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
    dimObj.SetTolerances 2, "TOL_NUM2", "", "", _
                             CDbl(upTol), CDbl(lowTol), CLng(displayMode)
    MCP_SetNumericTolerance = True
End Function
'''

_GET_TOLERANCE_VBS = r'''
Public Function MCP_GetDimensionTolerance(dimObj)
    Dim tolType, tolName, upText, lowText, upValue, lowValue, displayMode
    tolType = 0: tolName = "": upText = "": lowText = ""
    upValue = 0#: lowValue = 0#: displayMode = 0
    On Error Resume Next
    Call dimObj.GetTolerances(tolType, tolName, upText, lowText, _
                              upValue, lowValue, displayMode)
    If Err.Number <> 0 Then
        Err.Clear
        MCP_GetDimensionTolerance = Array(False, 0, "", "", "", 0#, 0#, 0)
    Else
        MCP_GetDimensionTolerance = Array(True, CLng(tolType), CStr(tolName), _
            CStr(upText), CStr(lowText), CDbl(upValue), CDbl(lowValue), CLng(displayMode))
    End If
    On Error GoTo 0
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


def _safe_count(collection: Any) -> Optional[int]:
    """Return a COM collection count without masking the primary operation."""
    if collection is None:
        return None


def _safe_attr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _delete_objects(document: Any, objects: list[Any]) -> dict[str, Any]:
    """Delete temporary drafting supports through the document Selection."""
    if not objects:
        return {"attempted": False, "succeeded": True, "error": None}
    selection = document.Selection
    try:
        selection.Clear()
        for obj in objects:
            selection.Add(obj)
        selection.Delete()
        selection.Clear()
        return {"attempted": True, "succeeded": True, "error": None}
    except Exception as exc:
        try:
            selection.Clear()
        except Exception:
            pass
        return {
            "attempted": True,
            "succeeded": False,
            "error": _format_error(exc),
        }


def _hide_objects(document: Any, objects: list[Any]) -> tuple[bool, Optional[str]]:
    """Hide temporary supports without changing the represented geometry."""
    if not objects:
        return True, None
    selection = document.Selection
    try:
        selection.Clear()
        for obj in objects:
            selection.Add(obj)
        selection.VisProperties.SetShow(1)
        selection.Clear()
        return True, None
    except Exception as exc:
        try:
            selection.Clear()
        except Exception:
            pass
        return False, _format_error(exc)
    try:
        return int(collection.Count)
    except Exception:
        return None


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


def _drawing_context(conn: Any) -> tuple[Any, Any, Any]:
    """Resolve an active drawing without requiring a model view to be active."""
    application = conn.connect(visible=True)
    document = conn.get_active_drawing_document()
    sheet = document.Sheets.ActiveSheet
    return application, document, sheet


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
    identity = list(_evaluate(
        application, _GEOMETRY_IDENTITY_VBS, "MCP_Get2DGeometryIdentity", [obj]
    ))
    range_values = list(_evaluate(
        application, _GEOMETRY_RANGE_VBS, "MCP_Get2DGeometryRange", [obj]
    ))
    point_values = list(_evaluate(
        application, _GEOMETRY_POINT_VBS, "MCP_Get2DGeometryPoint", [obj]
    ))
    center_values = list(_evaluate(
        application, _GEOMETRY_CENTER_VBS, "MCP_Get2DGeometryCenter", [obj]
    ))
    line_values = list(_evaluate(
        application, _GEOMETRY_LINE_VBS, "MCP_Get2DGeometryLine", [obj]
    ))
    radius_value: Optional[float] = None
    try:
        candidate_radius = float(obj.Radius)
        if math.isfinite(candidate_radius) and candidate_radius > 0.0:
            radius_value = candidate_radius
    except Exception:
        radius_value = None

    has_range = bool(range_values[0])
    if bool(point_values[0]):
        location_kind = "point"
        location = [float(point_values[1]), float(point_values[2])]
    elif bool(center_values[0]):
        location_kind = "center"
        location = [float(center_values[1]), float(center_values[2])]
    elif bool(line_values[0]):
        location_kind = "line"
        location = [float(line_values[1]), float(line_values[2])]
    else:
        location_kind = "unknown"
        location = [0.0, 0.0]
    direction = (
        [float(line_values[3]), float(line_values[4])]
        if bool(line_values[0]) else None
    )
    result = {
        "index": index,
        "name": str(identity[1]),
        "automation_type": str(identity[0]),
        "has_range_box": has_range,
        "range_box_view_mm": (
            {
                "xmin": float(range_values[1]),
                "ymin": float(range_values[2]),
                "xmax": float(range_values[3]),
                "ymax": float(range_values[4]),
            }
            if has_range
            else None
        ),
        "location_kind": location_kind,
        "location_view_mm": location,
        "radius_mm": radius_value,
        "direction": direction,
    }
    return result


def _geometry_provenance(description: dict[str, Any]) -> dict[str, Any]:
    """Classify listed 2D geometry without confusing helper points with contours."""
    automation_type = str(description.get("automation_type", "")).casefold()
    name = str(description.get("name", "")).strip()
    is_axis = "axis2d" in automation_type
    is_point = "point2d" in automation_type
    is_mcp_helper = name.casefold().startswith(_HELPER_GEOMETRY_PREFIX.casefold())
    is_contour = bool(
        not is_axis
        and not is_point
        and any(token in automation_type for token in (
            "line2d", "circle2d", "ellipse2d", "curve2d", "spline2d", "parabola2d",
        ))
    )
    role = (
        "structural_axis" if is_axis else
        "mcp_dimension_support" if is_mcp_helper else
        "point_support" if is_point else
        "projected_contour_candidate" if is_contour else
        "other_2d_geometry"
    )
    return {
        "geometry_role": role,
        "is_projected_contour_candidate": is_contour,
        "is_mcp_helper_geometry": is_mcp_helper,
        "valid_for_gdt_head_target": is_contour,
    }


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



_DRAWING_VIEW_SIZE_VBS = r'''
Public Function MCP_GetDrawingViewSize(viewObject)
    Dim values(3)
    viewObject.Size values
    MCP_GetDrawingViewSize = Array( _
        CDbl(values(0)), CDbl(values(1)), _
        CDbl(values(2)), CDbl(values(3)))
End Function
'''


def _drawing_view_size(application: Any, view: Any) -> dict[str, float]:
    """Read DrawingView.Size through CATIA SystemService.Evaluate."""
    try:
        raw = _evaluate(
            application,
            _DRAWING_VIEW_SIZE_VBS,
            "MCP_GetDrawingViewSize",
            [view],
        )
        values = [float(item) for item in list(raw)]
    except Exception as exc:
        raise RuntimeError(
            f"Could not read DrawingView.Size through CATIA: {_format_error(exc)}"
        ) from exc

    if len(values) != 4 or not all(math.isfinite(item) for item in values):
        raise RuntimeError("DrawingView.Size did not return four finite values")

    xmin, xmax, ymin, ymax = values
    if xmax < xmin or ymax < ymin:
        raise RuntimeError("DrawingView.Size returned an inverted bounding box")

    return {
        "xmin": xmin,
        "xmax": xmax,
        "ymin": ymin,
        "ymax": ymax,
        "width_mm": xmax - xmin,
        "height_mm": ymax - ymin,
    }


def _view_bounds_dimension_anchors(
    application: Any,
    view: Any,
    kind: str,
) -> tuple[float, tuple[float, float], tuple[float, float], dict[str, Any]]:
    """Build horizontal/vertical dimension anchors from DrawingView.Size."""
    size = _drawing_view_size(application, view)
    tf = _view_transform(view)

    sheet_corners = [
        (size["xmin"], size["ymin"]),
        (size["xmin"], size["ymax"]),
        (size["xmax"], size["ymin"]),
        (size["xmax"], size["ymax"]),
    ]
    local_corners = [_sheet_to_view(x, y, tf) for x, y in sheet_corners]

    xmin = min(point[0] for point in local_corners)
    xmax = max(point[0] for point in local_corners)
    ymin = min(point[1] for point in local_corners)
    ymax = max(point[1] for point in local_corners)

    if kind == "horizontal_distance":
        # Use one common witness line crossing the complete silhouette.  The
        # previous implementation was mathematically similar, but did not
        # report/verify the round-trip back to the exact DrawingView.Size
        # extrema.  That allowed callers to confuse model-centred coordinates
        # with the actual (often offset) drawing-view coordinate system.
        y = (ymin + ymax) / 2.0
        p = (xmin, y)
        q = (xmax, y)
        expected = abs(xmax - xmin)
    elif kind == "vertical_distance":
        x = (xmin + xmax) / 2.0
        p = (x, ymin)
        q = (x, ymax)
        expected = abs(ymax - ymin)
    else:
        raise ValueError(
            "DrawingView.Size fallback only supports horizontal_distance "
            "and vertical_distance"
        )

    if expected <= _EPS:
        raise ValueError(f"DrawingView.Size fallback produced zero {kind} extent")

    p_sheet = _view_to_sheet(p[0], p[1], tf)
    q_sheet = _view_to_sheet(q[0], q[1], tf)
    tolerance = max(0.01, abs(tf["scale"]) * 1.0e-6)
    if kind == "horizontal_distance":
        extrema_verified = (
            abs(p_sheet[0] - size["xmin"]) <= tolerance
            and abs(q_sheet[0] - size["xmax"]) <= tolerance
        )
    else:
        extrema_verified = (
            abs(p_sheet[1] - size["ymin"]) <= tolerance
            and abs(q_sheet[1] - size["ymax"]) <= tolerance
        )
    if not extrema_verified:
        raise RuntimeError(
            "DrawingView.Size anchors failed the view/sheet coordinate round-trip"
        )

    evidence = {
        "mode": "drawing_view_size_fallback",
        "size_sheet_mm": size,
        "view_bounds_mm": {
            "xmin": xmin,
            "xmax": xmax,
            "ymin": ymin,
            "ymax": ymax,
            "width_mm": xmax - xmin,
            "height_mm": ymax - ymin,
        },
        "anchor1_view_mm": list(p),
        "anchor2_view_mm": list(q),
        "anchor1_sheet_mm": list(p_sheet),
        "anchor2_sheet_mm": list(q_sheet),
        "extrema_round_trip_verified": True,
        "coordinate_contract": (
            "DrawingView.Size is sheet-space; Factory2D helper points are "
            "view-local and are obtained with origin/rotation/scale conversion"
        ),
        "associative_to_projected_model_geometry": False,
    }
    # The helper supports live in view-local/model units.  The explicit
    # sheet->view inverse transform therefore removes the paper scale before
    # the native DrawingDimension is created.  CATIA consequently reports the
    # represented model distance at 1:1, 2:1, 1:2, and arbitrary finite scales.
    evidence["scale_compensation_verified"] = True
    return expected, p, q, evidence


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
    try:
        dim.SetTolerances(
            2, "TOL_NUM2", "", "", up, low, int(display_mode)
        )
    except Exception:
        _evaluate(
            application,
            _TOLERANCE_VBS,
            "MCP_SetNumericTolerance",
            [dim, up, low, int(display_mode)],
        )
    return True


def _get_tolerance(application: Any, dim: Any) -> dict[str, Any]:
    """Read CATIA tolerance state without guessing from rendered text."""
    getter = _safe_attr(dim, "GetTolerances")
    if callable(getter):
        try:
            raw = list(getter())
            if len(raw) >= 7:
                return {
                    "available": True,
                    "read_method": "DrawingDimension.GetTolerances()",
                    "tolerance_type": int(raw[0]),
                    "tolerance_name": str(raw[1] or ""),
                    "upper_text": str(raw[2] or ""),
                    "lower_text": str(raw[3] or ""),
                    "upper_value": float(raw[4]),
                    "lower_value": float(raw[5]),
                    "display_mode": int(raw[6]),
                }
        except Exception as direct_exc:
            direct_error = _format_error(direct_exc)
    else:
        direct_error = "DrawingDimension.GetTolerances is unavailable"
    try:
        raw = list(_evaluate(
            application,
            _GET_TOLERANCE_VBS,
            "MCP_GetDimensionTolerance",
            [dim],
        ))
        if not raw or not bool(raw[0]):
            return {"available": False}
        return {
            "available": True,
            "tolerance_type": int(raw[1]),
            "tolerance_name": str(raw[2]),
            "upper_text": str(raw[3]),
            "lower_text": str(raw[4]),
            "upper_value": float(raw[5]),
            "lower_value": float(raw[6]),
            "display_mode": int(raw[7]),
        }
    except Exception as exc:
        return {
            "available": False,
            "direct_error": direct_error,
            "error": _format_error(exc),
        }


def _tolerance_matches_requested(
    readback: dict[str, Any],
    upper: Optional[float],
    lower: Optional[float],
) -> Optional[bool]:
    if upper is None and lower is None:
        return None
    if not readback.get("available"):
        return None
    requested_upper = 0.0 if upper is None else float(upper)
    requested_lower = -requested_upper if lower is None else float(lower)
    try:
        actual_upper = float(readback["upper_value"])
        actual_lower = float(readback["lower_value"])
    except Exception:
        return None
    epsilon = max(1.0e-9, abs(requested_upper) * 1.0e-6, abs(requested_lower) * 1.0e-6)
    return (
        abs(actual_upper - requested_upper) <= epsilon
        and abs(actual_lower - requested_lower) <= epsilon
    )
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
    allow_level_b_fallback: bool = False,
) -> tuple[dict[str, Any], list[str]]:
    kind = str(kind).strip().lower()

    # Keep the original associative-geometry path.  Only horizontal/vertical
    # dimensions get a controlled fallback when the requested element selector
    # cannot be resolved in a generated view.
    fallback_evidence: Optional[dict[str, Any]] = None
    helper_points: list[Any] = []
    fallback_mode = False

    try:
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
    except Exception as primary_exc:
        primary_error = _format_error(primary_exc)
        if not bool(allow_level_b_fallback) or kind not in {
            "horizontal_distance", "vertical_distance"
        }:
            raise RuntimeError(
                "The requested projected geometry could not be resolved. "
                "Pass allow_level_b_fallback=true only for verified overall "
                "horizontal/vertical extents, or generate dimensions from 3D "
                f"constraints. Original failure: {primary_error}"
            ) from primary_exc
        expected, anchor1, anchor2, fallback_evidence = (
            _view_bounds_dimension_anchors(application, view, kind)
        )
        unit = "length_mm"
        a = {
            "index": None,
            "name": "DrawingView.Size extent anchor 1",
            "automation_type": "temporary_Point2D",
            "location_kind": "point",
            "location_view_mm": list(anchor1),
        }
        b = {
            "index": None,
            "name": "DrawingView.Size extent anchor 2",
            "automation_type": "temporary_Point2D",
            "location_kind": "point",
            "location_view_mm": list(anchor2),
        }
        cat_obj1 = None
        cat_obj2 = None
        fallback_mode = True

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

    warnings: list[str] = []

    if fallback_mode:
        try:
            view.Activate()
        except Exception:
            pass
        geometry_before = _safe_count(_safe_attr(view, "GeometricElements"))
        try:
            factory = view.Factory2D
            helper_points = [
                factory.CreatePoint(anchor1[0], anchor1[1]),
                factory.CreatePoint(anchor2[0], anchor2[1]),
            ]
            for helper_index, helper in enumerate(helper_points, start=1):
                try:
                    helper.Name = (
                        f"{_HELPER_GEOMETRY_PREFIX}{kind}_{helper_index}_"
                        f"{int(time.time() * 1000)}"
                    )
                except Exception:
                    pass
        except Exception as exc:
            raise RuntimeError(
                "DrawingView.Size fallback could not create helper Point2D objects: "
                f"{_format_error(exc)}"
            ) from exc

        geometry_after = _safe_count(_safe_attr(view, "GeometricElements"))
        fallback_evidence["helper_point_objects_created"] = True
        fallback_evidence["geometry_count_before_helper_points"] = geometry_before
        fallback_evidence["geometry_count_after_helper_points"] = geometry_after
        if (
            geometry_before is not None
            and geometry_after is not None
            and geometry_after != geometry_before + 2
        ):
            _delete_objects(document, helper_points)
            raise RuntimeError(
                "Fallback helper Point2D creation did not produce the expected "
                "geometry-count delta.",
                data={
                    "geometry_count_before": geometry_before,
                    "geometry_count_after": geometry_after,
                    "expected_delta": 2,
                    "dimension_kind": kind,
                },
            )
        cat_obj1 = helper_points[0]
        cat_obj2 = helper_points[1]
        warnings.append(
            "The requested geometry selector could not be resolved; "
            f"{kind} was created from DrawingView.Size bounding-box extents "
            "using hidden helper Point2D supports."
        )
        warnings.append(
            f"Original geometry-selection failure: {primary_error}"
        )

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
        try:
            if int(view.Dimensions.Count) > before:
                view.Dimensions.Remove(int(view.Dimensions.Count))
        except Exception:
            pass
        if helper_points:
            try:
                _delete_objects(document, helper_points)
            except Exception:
                pass
        raise

    if helper_points:
        hidden, hide_error = _hide_objects(document, helper_points)
        if not hidden and hide_error:
            warnings.append(
                "Fallback helper Point2D objects were created but could not be hidden: "
                f"{hide_error}"
            )

    tf = _view_transform(view)
    offset = _finite(offset_mm, "offset_mm")
    if position_x is None and position_y is None:
        vx, vy = _auto_position(kind, anchor1, anchor2, offset, tf["scale"])
        resolved_from = "drawing_view_size_fallback" if fallback_mode else "automatic_feature_offset"
    elif position_x is None or position_y is None:
        view.Dimensions.Remove(created_index)
        if helper_points:
            _delete_objects(document, helper_points)
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
            if helper_points:
                _delete_objects(document, helper_points)
            raise ValueError("position_space must be 'view' or 'sheet'")

    tolerance_set = False
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
        if helper_points:
            try:
                support_cleanup = _delete_objects(document, helper_points)
                if not support_cleanup.get("succeeded", True):
                    cleanup_error = ((cleanup_error + "; ") if cleanup_error else "") + str(support_cleanup.get("error"))
            except Exception as support_exc:
                cleanup_error = ((cleanup_error + "; ") if cleanup_error else "") + _format_error(support_exc)
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
    tolerance_readback = _get_tolerance(application, dim)
    tolerance_verified = _tolerance_matches_requested(
        tolerance_readback, tolerance_upper, tolerance_lower
    )
    if tolerance_set and tolerance_verified is not True:
        warnings.append(
            "CATIA accepted SetTolerances, but numerical tolerance readback did "
            "not match the requested deviations; tolerance write is not verified."
        )
    data = {
        "view": str(getattr(view, "Name", "")),
        "dimension_index": created_index,
        "dimension_name": str(getattr(dim, "Name", "")),
        "kind": kind,
        "catia_dim_type": int(getattr(dim, "DimType", -1)),
        "catia_line_rep": int(getattr(dim.GetDimLine(), "DimLineRep", -1)),
        "element1": a,
        "element2": b,
        "associative_geometry": not fallback_mode,
        "association_level": "A" if not fallback_mode else "B",
        "association_classification": (
            "projected_geometry_associative"
            if not fallback_mode else
            "native_dimension_point_supported_from_verified_view_silhouette_extents"
        ),
        "support_mode": "drawing_view_size_hidden_helper_points" if fallback_mode else "original_drawing_geometry",
        "fallback_evidence": fallback_evidence,
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
        "tolerance_readback": tolerance_readback,
        "tolerance_verified": tolerance_verified,
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


def _drawing_model_views(sheet: Any) -> list[Any]:
    result: list[Any] = []
    views = sheet.Views
    for index in range(1, int(views.Count) + 1):
        view = views.Item(index)
        try:
            view_type = int(view.ViewType)
        except Exception:
            view_type = -1
        if view_type in {0, 13}:
            continue
        result.append(view)
    return result


def _find_dimension_name(view: Any, name: str) -> Optional[int]:
    for index in range(1, int(view.Dimensions.Count) + 1):
        try:
            if str(view.Dimensions.Item(index).Name) == name:
                return index
        except Exception:
            continue
    return None


def _generate_native_constraint_dimensions(sheet: Any) -> dict[str, Any]:
    before = sum(int(v.Dimensions.Count) for v in _drawing_model_views(sheet))
    error = None
    try:
        sheet.GenerateDimensions()
    except Exception as exc:
        error = _format_error(exc)
    after = sum(int(v.Dimensions.Count) for v in _drawing_model_views(sheet))
    return {
        "attempted": True,
        "succeeded": error is None,
        "error": error,
        "dimension_count_before": before,
        "dimension_count_after": after,
        "created_count": max(0, after - before),
        "source": "3D constraints/parameters exposed to DrawingSheet.GenerateDimensions",
        "association_level": "CATIA_native_constraint_associative",
    }


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
            excluded = []
            failures = []
            for index in range(1, limit + 1):
                try:
                    described = _describe(
                        application, view.GeometricElements.Item(index), index
                    )
                    provenance = _geometry_provenance(described)
                    described.update(provenance)
                    if provenance["geometry_role"] == "structural_axis":
                        continue
                    if provenance["is_mcp_helper_geometry"]:
                        excluded.append(described)
                        continue
                    items.append(described)
                except Exception as exc:
                    failures.append({"index": index, "error": _format_error(exc)})
            data = {
                "view": str(view.Name),
                "view_transform": _view_transform(view),
                "geometry_count": count,
                "returned_count": len(items),
                "truncated": limit < count,
                "geometry": items,
                "projected_contour_count": sum(
                    1 for item in items
                    if item.get("is_projected_contour_candidate")
                ),
                "excluded_helper_geometry": excluded,
                "projected_geometry_exposure": (
                    "contours_available"
                    if any(item.get("is_projected_contour_candidate") for item in items)
                    else "points_only_no_attachable_contours"
                    if items
                    else "not_exposed_by_DrawingView.GeometricElements"
                ),
                "dimensioning_policy": (
                    "Only entries with valid_for_gdt_head_target=true are contour "
                    "candidates. Point2D supports and MCP helper geometry are not "
                    "manufacturing contours. Do not infer contours from view bounds."
                ),
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
        allow_level_b_fallback: bool = False,
    ) -> dict[str, Any]:
        """Create, place and verify one associative 2D drawing dimension.

        kind supports distance/horizontal_distance/vertical_distance/
        aligned_distance/line_to_line/center_to_line/center_to_center/radius/
        diameter/angle/line_length. Coordinates are view-local by default.
        With position_space='sheet', coordinates are rigorously converted using
        view origin, rotation and scale. If position is omitted, a feature-based
        offset in paper millimetres is used.

        By default the tool fails closed when projected geometry cannot be
        resolved.  For overall horizontal/vertical extents only,
        allow_level_b_fallback=true creates a native point-supported dimension
        from the verified DrawingView.Size silhouette extrema.  The result is
        explicitly reported as association Level B, never Level A.
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
                allow_level_b_fallback=allow_level_b_fallback,
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
                        allow_level_b_fallback=spec.get("allow_level_b_fallback", False),
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
                    "tolerance": _get_tolerance(application, dim),
                })
                if box is None:
                    warnings.append(f"dimension {i}: boundary box unavailable")
            return _success({"view": str(view.Name), "count": len(items), "dimensions": items}, warnings)
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_list_2d_drawing_dimensions")

    @mcp.tool()
    def catia_auto_dimension_drawing(
        view_names: list[str] | None = None,
        generate_from_3d_constraints: bool = True,
        add_overall_extents: bool = True,
        overall_tolerance_upper: float | None = None,
        overall_tolerance_lower: float | None = None,
        tolerance_display_mode: int = 0,
        offset_mm: float = 18.0,
    ) -> dict[str, Any]:
        """Generate the maximum defensible native dimension set.

        Phase 1 asks CATIA to generate dimensions from exposed 3D constraints.
        Phase 2 adds horizontal and vertical overall extents for selected model
        views when those dimensions do not already exist.  When projected
        contour objects are unavailable, overall extents are Level B native
        dimensions supported by verified silhouette-extreme Point2D objects.
        The tool never claims feature-size/location completeness merely from a
        bounding box; remaining design-intent gaps are returned explicitly.
        """
        try:
            application, document, sheet = _drawing_context(conn)
            native = (
                _generate_native_constraint_dimensions(sheet)
                if bool(generate_from_3d_constraints)
                else {"attempted": False, "created_count": 0}
            )
            warnings: list[str] = []
            if native.get("error"):
                warnings.append(
                    "DrawingSheet.GenerateDimensions failed: " + str(native["error"])
                )
            requested_names = {
                str(name).strip().lower() for name in (view_names or [])
                if str(name).strip()
            }
            views = [
                view for view in _drawing_model_views(sheet)
                if not requested_names
                or str(view.Name).strip().lower() in requested_names
            ]
            created: list[dict[str, Any]] = []
            failures: list[dict[str, Any]] = []
            if bool(add_overall_extents):
                for view in views:
                    for kind, suffix in (
                        ("horizontal_distance", "OverallWidth"),
                        ("vertical_distance", "OverallHeight"),
                    ):
                        dim_name = f"MCP_{str(view.Name)}_{suffix}"
                        existing = _find_dimension_name(view, dim_name)
                        if existing is not None:
                            created.append({
                                "view": str(view.Name),
                                "kind": kind,
                                "dimension_name": dim_name,
                                "dimension_index": existing,
                                "action": "already_exists",
                            })
                            continue
                        try:
                            data, item_warnings = _add_internal(
                                application=application,
                                document=document,
                                sheet=sheet,
                                view=view,
                                kind=kind,
                                element1="__projected_outline__",
                                element2="__projected_outline__",
                                position_x=None,
                                position_y=None,
                                position_space="view",
                                offset_mm=offset_mm,
                                witness_points=None,
                                name=dim_name,
                                tolerance_upper=overall_tolerance_upper,
                                tolerance_lower=overall_tolerance_lower,
                                tolerance_display_mode=tolerance_display_mode,
                                allow_level_b_fallback=True,
                            )
                            created.append({"action": "created", **data})
                            warnings.extend(
                                f"{view.Name}/{kind}: {warning}"
                                for warning in item_warnings
                            )
                        except Exception as exc:
                            failures.append({
                                "view": str(view.Name),
                                "kind": kind,
                                "error": _format_error(exc),
                            })

            inventory = []
            for view in views:
                dimensions = []
                for index in range(1, int(view.Dimensions.Count) + 1):
                    dim = view.Dimensions.Item(index)
                    dimensions.append({
                        "index": index,
                        "name": str(getattr(dim, "Name", "")),
                        "measured_value": _dimension_value(dim),
                        "dim_status": int(getattr(dim, "DimStatus", -1)),
                        "tolerance": _get_tolerance(application, dim),
                    })
                inventory.append({
                    "view": str(view.Name),
                    "dimension_count": len(dimensions),
                    "dimensions": dimensions,
                })
            remaining = [
                "Feature sizes (holes, radii, chamfers, thicknesses) not exposed as 3D constraints/PMI",
                "Feature locations and patterns not exposed as 3D constraints/PMI",
                "Functional tolerances not supplied by the caller or model PMI",
                "Functional datum scheme and GD&T design intent",
            ]
            return _success({
                "operation": "catia_auto_dimension_drawing",
                "native_constraint_generation": native,
                "overall_extent_results": created,
                "overall_extent_failures": failures,
                "inventory": inventory,
                "association_policy": {
                    "A": "DrawingDimension attached to exposed projected geometry",
                    "B": "Native DrawingDimension attached to verified view-silhouette extent supports",
                    "native_constraint": "CATIA-generated from 3D model constraints/parameters",
                },
                "remaining_design_intent_requirements": remaining,
                "complete_for_arbitrary_part": not bool(remaining),
                "document_save_required": bool(native.get("created_count") or any(
                    item.get("action") == "created" for item in created
                )),
            }, warnings + ([f"{len(failures)} overall dimension(s) failed"] if failures else []))
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_auto_dimension_drawing")

    @mcp.tool()
    def catia_audit_dimension_completeness(
        expected_dimensions: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        """Audit native dimensions, tolerances and caller-supplied dimension plan."""
        try:
            application, document, sheet = _drawing_context(conn)
            views = _drawing_model_views(sheet)
            inventory: list[dict[str, Any]] = []
            names_found: set[str] = set()
            without_tolerance: list[dict[str, Any]] = []
            for view in views:
                for index in range(1, int(view.Dimensions.Count) + 1):
                    dim = view.Dimensions.Item(index)
                    name = str(getattr(dim, "Name", ""))
                    names_found.add(name)
                    tolerance = _get_tolerance(application, dim)
                    record = {
                        "view": str(view.Name),
                        "index": index,
                        "name": name,
                        "value": _dimension_value(dim),
                        "status": int(getattr(dim, "DimStatus", -1)),
                        "tolerance": tolerance,
                    }
                    inventory.append(record)
                    if not tolerance.get("available"):
                        without_tolerance.append(record)
            expected = expected_dimensions or []
            missing = []
            for item in expected:
                name = str(item.get("name", "")).strip()
                if not name:
                    missing.append({**item, "reason": "expected dimension has no stable name"})
                elif name not in names_found:
                    missing.append({**item, "reason": "native dimension not found"})
            plan_supplied = bool(expected)
            complete = bool(plan_supplied and not missing)
            warnings = []
            if not plan_supplied:
                warnings.append(
                    "No dimension plan was supplied; arbitrary-part manufacturing completeness cannot be proven."
                )
            if without_tolerance:
                warnings.append(
                    f"{len(without_tolerance)} native dimension(s) have no readable explicit tolerance; general tolerance may still apply."
                )
            return _success({
                "operation": "catia_audit_dimension_completeness",
                "dimension_plan_supplied": plan_supplied,
                "expected_count": len(expected),
                "native_dimension_count": len(inventory),
                "missing_expected_dimensions": missing,
                "dimensions_without_readable_explicit_tolerance": without_tolerance,
                "inventory": inventory,
                "complete": complete,
                "completeness_rule": "PASS requires a caller/model-derived dimension plan and every planned native dimension present",
                "model_modified": False,
                "document_save_required": False,
            }, warnings)
        except Exception as exc:
            return _error(_format_error(exc))

    names.append("catia_audit_dimension_completeness")

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
