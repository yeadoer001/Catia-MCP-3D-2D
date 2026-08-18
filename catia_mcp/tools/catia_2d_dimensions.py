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


IMPLEMENTATION_VERSION = "catia-2d-dimensions-2026-08-18-v4-full-semantic-layout"
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
    "line_length": "Length of one bounded Line2D using its visible segment endpoints.",
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


class AnnotationOperationError(RuntimeError):
    """Operation error carrying optional diagnostic data."""

    def __init__(self, message: str, *, data: Any = None) -> None:
        super().__init__(message)
        self.data = data


def _safe_attr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _safe_count(collection: Any) -> Optional[int]:
    try:
        return int(collection.Count)
    except Exception:
        return None


def _hide_objects(document: Any, objects: list[Any]) -> tuple[bool, Optional[str]]:
    selection = None
    try:
        selection = document.Selection
        selection.Clear()
        for obj in objects:
            selection.Add(obj)
        selection.VisProperties.SetShow(1)
        return True, None
    except Exception as exc:
        return False, _format_error(exc)
    finally:
        if selection is not None:
            try:
                selection.Clear()
            except Exception:
                pass


def _delete_objects(document: Any, objects: list[Any]) -> dict[str, Any]:
    result = {"attempted": bool(objects), "succeeded": True, "error": None}
    if not objects:
        return result
    selection = None
    try:
        selection = document.Selection
        selection.Clear()
        for obj in objects:
            selection.Add(obj)
        selection.Delete()
        return result
    except Exception as exc:
        result["succeeded"] = False
        result["error"] = _format_error(exc)
        return result
    finally:
        if selection is not None:
            try:
                selection.Clear()
            except Exception:
                pass


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


def _bbox_center(item: dict[str, Any]) -> tuple[float, float]:
    box = item.get("range_box_view_mm")
    if not box:
        raise ValueError(f"geometry {item.get('index')} has no range box")
    return ((float(box["xmin"]) + float(box["xmax"])) / 2.0,
            (float(box["ymin"]) + float(box["ymax"])) / 2.0)


def _point(item: dict[str, Any]) -> tuple[float, float]:
    """Return a visually meaningful anchor, never a Line2D origin by default.

    Point/centre geometry keeps its native coordinate.  Line2D uses the midpoint of
    its visible segment; other bounded curves use their range-box centre as a
    conservative fallback.  This avoids the old failure where GetOrigin() could be
    near or outside the visible segment.
    """
    kind = str(item.get("location_kind", ""))
    if kind in {"point", "center"}:
        return tuple(item["location_view_mm"])  # type: ignore[return-value]
    if kind == "line" and item.get("direction") and item.get("range_box_view_mm"):
        p0, p1 = _line_visible_segment(item)
        return ((p0[0] + p1[0]) / 2.0, (p0[1] + p1[1]) / 2.0)
    if item.get("range_box_view_mm"):
        return _bbox_center(item)
    if kind == "line":
        return tuple(item["location_view_mm"])  # type: ignore[return-value]
    raise ValueError(f"geometry {item.get('index')} has no usable anchor")


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


def _line_visible_segment(
    item: dict[str, Any],
) -> tuple[tuple[float, float], tuple[float, float]]:
    """Recover the visible Line2D segment from its line equation + range box.

    CATIA Line2D.GetOrigin() returns a point on the underlying line, but that point
    is not guaranteed to be the visual centre of the projected segment.  For
    dimension placement we therefore intersect the infinite line with the
    geometry RangeBox and use the two visible boundary intersections.
    """
    box = item.get("range_box_view_mm")
    if not box:
        raise ValueError(f"geometry {item['index']} has no range box")

    x0, y0, ux, uy = _line(item)
    xmin, xmax = float(box["xmin"]), float(box["xmax"])
    ymin, ymax = float(box["ymin"]), float(box["ymax"])
    tolerance = 1.0e-7
    candidates: list[tuple[float, tuple[float, float]]] = []

    def add_candidate(t: float) -> None:
        x = x0 + t * ux
        y = y0 + t * uy
        if (
            xmin - tolerance <= x <= xmax + tolerance
            and ymin - tolerance <= y <= ymax + tolerance
        ):
            for old_t, _ in candidates:
                if abs(old_t - t) <= tolerance:
                    return
            candidates.append((t, (x, y)))

    if abs(ux) > _EPS:
        add_candidate((xmin - x0) / ux)
        add_candidate((xmax - x0) / ux)
    if abs(uy) > _EPS:
        add_candidate((ymin - y0) / uy)
        add_candidate((ymax - y0) / uy)

    if len(candidates) >= 2:
        candidates.sort(key=lambda entry: entry[0])
        return candidates[0][1], candidates[-1][1]

    # Degenerate/numerically tiny range boxes: keep a conservative fallback.
    corners = [
        (xmin, ymin), (xmin, ymax), (xmax, ymin), (xmax, ymax),
    ]
    projected = sorted(
        corners,
        key=lambda p: (p[0] - x0) * ux + (p[1] - y0) * uy,
    )
    return projected[0], projected[-1]


def _range_endpoints(
    item: dict[str, Any],
) -> tuple[tuple[float, float], tuple[float, float]]:
    if item.get("location_kind") == "line" and item.get("direction"):
        return _line_visible_segment(item)
    box = item.get("range_box_view_mm")
    if not box:
        raise ValueError(f"geometry {item['index']} has no range box")
    return (box["xmin"], box["ymin"]), (box["xmax"], box["ymax"])


def _parallel_line_dimension_geometry(
    a: dict[str, Any],
    b: dict[str, Any],
) -> tuple[float, tuple[float, float], tuple[float, float], dict[str, Any]]:
    """Choose dimension anchors in the common visible span of two parallel lines.

    The old implementation used Line2D.GetOrigin() from line A and projected it
    onto line B.  That is mathematically valid for the distance, but the origin can
    lie near an end or even outside the visually useful common span.  Here we use
    each line's visible segment and choose the centre of their overlapping span
    along the line direction.
    """
    x1, y1, ux1, uy1 = _line(a)
    x2, y2, ux2, uy2 = _line(b)
    cross = ux1 * uy2 - uy1 * ux2
    if abs(cross) > 1.0e-5:
        raise ValueError("line_to_line requires parallel lines; use kind='angle'")

    # Keep both directions consistent so longitudinal parameters compare cleanly.
    if ux1 * ux2 + uy1 * uy2 < 0.0:
        ux2, uy2 = -ux2, -uy2

    n1x, n1y = -uy1, ux1
    signed_distance = (x2 - x1) * n1x + (y2 - y1) * n1y
    distance = abs(signed_distance)

    a0, a1 = _line_visible_segment(a)
    b0, b1 = _line_visible_segment(b)

    def longitudinal(point: tuple[float, float]) -> float:
        return point[0] * ux1 + point[1] * uy1

    a_lo, a_hi = sorted((longitudinal(a0), longitudinal(a1)))
    b_lo, b_hi = sorted((longitudinal(b0), longitudinal(b1)))
    overlap_lo = max(a_lo, b_lo)
    overlap_hi = min(a_hi, b_hi)

    if overlap_lo <= overlap_hi + 1.0e-7:
        longitudinal_target = (overlap_lo + overlap_hi) / 2.0
        span_mode = "visible_overlap_midpoint"
    else:
        # No common projected span: choose the midpoint between the nearest ends.
        if a_hi < b_lo:
            longitudinal_target = (a_hi + b_lo) / 2.0
        else:
            longitudinal_target = (b_hi + a_lo) / 2.0
        span_mode = "nearest_visible_span_midpoint_no_overlap"

    # Point p lies on A at the selected longitudinal station.
    t1 = longitudinal_target - (x1 * ux1 + y1 * uy1)
    p = (x1 + t1 * ux1, y1 + t1 * uy1)

    # q is the normal projection of p onto B.
    q = (p[0] + signed_distance * n1x, p[1] + signed_distance * n1y)

    evidence = {
        "strategy": "parallel_line_common_visible_span",
        "line1_visible_segment_view_mm": [list(a0), list(a1)],
        "line2_visible_segment_view_mm": [list(b0), list(b1)],
        "line1_longitudinal_interval": [a_lo, a_hi],
        "line2_longitudinal_interval": [b_lo, b_hi],
        "overlap_interval": [overlap_lo, overlap_hi],
        "span_mode": span_mode,
        "chosen_longitudinal_coordinate": longitudinal_target,
        "anchor1_view_mm": list(p),
        "anchor2_view_mm": list(q),
        "distance_mm": distance,
        "line_direction": [ux1, uy1],
    }
    return distance, p, q, evidence



def _line_intersection(
    a: dict[str, Any], b: dict[str, Any]
) -> tuple[float, float]:
    x1, y1, ux1, uy1 = _line(a)
    x2, y2, ux2, uy2 = _line(b)
    det = ux1 * uy2 - uy1 * ux2
    if abs(det) <= 1.0e-7:
        raise ValueError("angle requires two non-parallel Line2D objects")
    dx, dy = x2 - x1, y2 - y1
    t = (dx * uy2 - dy * ux2) / det
    return x1 + t * ux1, y1 + t * uy1


def _ray_toward_visible_segment(
    item: dict[str, Any], vertex: tuple[float, float]
) -> tuple[float, float]:
    """Choose the line direction that points from the vertex toward visible geometry."""
    _, _, ux, uy = _line(item)
    try:
        a0, a1 = _line_visible_segment(item)
        mx, my = (a0[0] + a1[0]) / 2.0, (a0[1] + a1[1]) / 2.0
        if (mx - vertex[0]) * ux + (my - vertex[1]) * uy < 0.0:
            ux, uy = -ux, -uy
    except Exception:
        pass
    return ux, uy


def _angle_dimension_geometry(
    a: dict[str, Any], b: dict[str, Any]
) -> tuple[float, tuple[float, float], tuple[float, float], dict[str, Any]]:
    """Build angle anchors from the actual line intersection and visible rays."""
    vertex = _line_intersection(a, b)
    u1 = _ray_toward_visible_segment(a, vertex)
    u2 = _ray_toward_visible_segment(b, vertex)
    dot = max(-1.0, min(1.0, u1[0] * u2[0] + u1[1] * u2[1]))
    angle = math.acos(dot)
    # Prefer the smaller included angle; CATIA's catDimAngle can still choose its
    # own representation, but the witness points now identify the intended sector.
    if angle > math.pi:
        angle = 2.0 * math.pi - angle
    try:
        s10, s11 = _line_visible_segment(a)
        s20, s21 = _line_visible_segment(b)
        len1 = max(math.dist(vertex, s10), math.dist(vertex, s11))
        len2 = max(math.dist(vertex, s20), math.dist(vertex, s21))
        radius = max(3.0, min(len1, len2) * 0.35)
    except Exception:
        radius = 10.0
    p = (vertex[0] + u1[0] * radius, vertex[1] + u1[1] * radius)
    q = (vertex[0] + u2[0] * radius, vertex[1] + u2[1] * radius)
    bisector = (u1[0] + u2[0], u1[1] + u2[1])
    bisector_norm = math.hypot(*bisector)
    if bisector_norm <= _EPS:
        bisector = (-u1[1], u1[0])
        bisector_norm = 1.0
    bisector = (bisector[0] / bisector_norm, bisector[1] / bisector_norm)
    evidence = {
        "strategy": "angle_vertex_visible_rays",
        "vertex_view_mm": list(vertex),
        "ray1": list(u1),
        "ray2": list(u2),
        "bisector": list(bisector),
        "witness_radius_view_mm": radius,
        "anchor1_view_mm": list(p),
        "anchor2_view_mm": list(q),
        "angle_radians": angle,
    }
    return angle, p, q, evidence


def _axis_distance_geometry(
    kind: str, a: dict[str, Any], b: dict[str, Any]
) -> tuple[float, tuple[float, float], tuple[float, float], dict[str, Any]]:
    """Select visually meaningful anchors for horizontal/vertical dimensions."""
    pa = _point(a)
    pb = _point(b)
    evidence: dict[str, Any] = {"strategy": "visual_anchor_axis_distance"}
    # If both supports are lines, use their visible spans rather than line origins.
    if a.get("location_kind") == "line" and b.get("location_kind") == "line":
        a0, a1 = _line_visible_segment(a)
        b0, b1 = _line_visible_segment(b)
        if kind == "horizontal_distance":
            ay = (a0[1] + a1[1]) / 2.0
            by = (b0[1] + b1[1]) / 2.0
            # Choose endpoints/points nearest in X so witness lines do not span the view arbitrarily.
            pair = min(((x, y) for x, y in (a0, a1) for _x, _y in [b0, b1]), key=lambda p: 0.0)
            candidates = [((x1, y1), (x2, y2)) for x1, y1 in (a0, a1) for x2, y2 in (b0, b1)]
            p, q = min(candidates, key=lambda pair: abs(pair[1][0] - pair[0][0]))
            p, q = (p[0], ay), (q[0], by)
        else:
            ax = (a0[0] + a1[0]) / 2.0
            bx = (b0[0] + b1[0]) / 2.0
            candidates = [((x1, y1), (x2, y2)) for x1, y1 in (a0, a1) for x2, y2 in (b0, b1)]
            p, q = min(candidates, key=lambda pair: abs(pair[1][1] - pair[0][1]))
            p, q = (ax, p[1]), (bx, q[1])
        evidence.update({
            "strategy": "visible_line_segments_axis_distance",
            "line1_visible_segment_view_mm": [list(a0), list(a1)],
            "line2_visible_segment_view_mm": [list(b0), list(b1)],
        })
    else:
        p, q = pa, pb
    value = abs(q[0] - p[0]) if kind == "horizontal_distance" else abs(q[1] - p[1])
    evidence["anchor1_view_mm"] = list(p)
    evidence["anchor2_view_mm"] = list(q)
    evidence["distance_mm"] = value
    return value, p, q, evidence


def _circle_dimension_geometry(
    kind: str, item: dict[str, Any]
) -> tuple[float, tuple[float, float], tuple[float, float], dict[str, Any]]:
    c = _center(item)
    r = float(item["radius_mm"])
    # Anchors identify actual geometry. Placement is decided later from free-space scoring.
    p = (c[0] + r, c[1])
    q = c if kind == "radius" else (c[0] - r, c[1])
    return (r if kind == "radius" else 2.0 * r), p, q, {
        "strategy": "circle_center_and_diameter_axis",
        "center_view_mm": list(c),
        "radius_mm": r,
        "anchor1_view_mm": list(p),
        "anchor2_view_mm": list(q),
    }


def _center_to_line_geometry(
    circle: dict[str, Any], line: dict[str, Any]
) -> tuple[float, tuple[float, float], tuple[float, float], dict[str, Any]]:
    p = _center(circle)
    x, y, dx, dy = _line(line)
    signed = (p[0] - x) * (-dy) + (p[1] - y) * dx
    q = (p[0] - signed * (-dy), p[1] - signed * dx)
    # Clamp the visual witness to the nearest point on the visible segment if the
    # perpendicular foot lies beyond it. Value remains the support-line distance.
    try:
        s0, s1 = _line_visible_segment(line)
        seg_dx, seg_dy = s1[0] - s0[0], s1[1] - s0[1]
        seg_len2 = seg_dx * seg_dx + seg_dy * seg_dy
        if seg_len2 > _EPS:
            t = ((q[0] - s0[0]) * seg_dx + (q[1] - s0[1]) * seg_dy) / seg_len2
            if t < 0.0 or t > 1.0:
                q_visual = s0 if t < 0.0 else s1
            else:
                q_visual = q
        else:
            q_visual = q
    except Exception:
        q_visual = q
    return abs(signed), p, q_visual, {
        "strategy": "center_to_visible_line_projection",
        "center_view_mm": list(p),
        "infinite_line_projection_view_mm": list(q),
        "visual_line_anchor_view_mm": list(q_visual),
        "distance_mm": abs(signed),
    }

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
    if kind == "line_length":
        _line(a)
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
) -> tuple[Optional[float], tuple[float, float], tuple[float, float], str, dict[str, Any]]:
    evidence: dict[str, Any] = {"strategy": "default_geometry_anchor"}
    if kind in {"radius", "diameter"}:
        value, p, q, evidence = _circle_dimension_geometry(kind, a)
        return value, p, q, "length_mm", evidence
    if kind == "line_length":
        p1, p2 = _range_endpoints(a)
        evidence = {
            "strategy": "visible_bounded_geometry_endpoints",
            "anchor1_view_mm": list(p1),
            "anchor2_view_mm": list(p2),
        }
        return math.dist(p1, p2), p1, p2, "length_mm", evidence
    assert b is not None
    if kind == "center_to_line":
        value, p, q, evidence = _center_to_line_geometry(a, b)
        return value, p, q, "length_mm", evidence
    if kind == "center_to_center":
        p, q = _center(a), _center(b)
        return math.dist(p, q), p, q, "length_mm", {
            "strategy": "center_to_center_midline",
            "anchor1_view_mm": list(p),
            "anchor2_view_mm": list(q),
        }
    if kind == "line_to_line":
        value, p, q, line_evidence = _parallel_line_dimension_geometry(a, b)
        return value, p, q, "length_mm", line_evidence
    if kind == "angle":
        value, p, q, angle_evidence = _angle_dimension_geometry(a, b)
        return value, p, q, "angle_radians", angle_evidence
    if kind in {"horizontal_distance", "vertical_distance"}:
        value, p, q, axis_evidence = _axis_distance_geometry(kind, a, b)
        return value, p, q, "length_mm", axis_evidence
    p, q = _point(a), _point(b)
    evidence = {
        "strategy": "visual_feature_anchors",
        "anchor1_view_mm": list(p),
        "anchor2_view_mm": list(q),
    }
    return math.dist(p, q), p, q, "length_mm", evidence


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
        "associative_to_projected_model_geometry": False,
    }
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


def _bbox_intersects(a: dict[str, float], b: dict[str, float], margin: float = 0.0) -> bool:
    return not (
        a["xmax"] + margin < b["xmin"]
        or b["xmax"] + margin < a["xmin"]
        or a["ymax"] + margin < b["ymin"]
        or b["ymax"] + margin < a["ymin"]
    )


def _point_clearance_score(
    point: tuple[float, float], obstacles: list[dict[str, float]], clearance: float
) -> float:
    score = 0.0
    x, y = point
    for box in obstacles:
        if (box["xmin"] - clearance <= x <= box["xmax"] + clearance and
                box["ymin"] - clearance <= y <= box["ymax"] + clearance):
            score += 1000.0
        cx = min(max(x, box["xmin"]), box["xmax"])
        cy = min(max(y, box["ymin"]), box["ymax"])
        d = math.hypot(x - cx, y - cy)
        score += 1.0 / max(d, 0.25)
    return score


def _geometry_obstacles(view: Any, application: Any, exclude_indices: set[int] | None = None) -> list[dict[str, float]]:
    excluded = exclude_indices or set()
    boxes: list[dict[str, float]] = []
    try:
        elements = view.GeometricElements
        count = int(elements.Count)
    except Exception:
        return boxes
    for i in range(1, count + 1):
        if i in excluded:
            continue
        try:
            desc = _describe(application, elements.Item(i), i)
            box = desc.get("range_box_view_mm")
            if box:
                boxes.append({k: float(box[k]) for k in ("xmin", "xmax", "ymin", "ymax")})
        except Exception:
            continue
    return boxes


def _existing_dimension_obstacles(application: Any, view: Any, exclude_index: Optional[int] = None) -> list[dict[str, float]]:
    boxes: list[dict[str, float]] = []
    try:
        count = int(view.Dimensions.Count)
    except Exception:
        return boxes
    for i in range(1, count + 1):
        if exclude_index is not None and i == exclude_index:
            continue
        try:
            box = _boundary(application, view.Dimensions.Item(i))
            if box:
                boxes.append({k: float(box[k]) for k in ("xmin", "xmax", "ymin", "ymax")})
        except Exception:
            continue
    return boxes


def _candidate_best(
    candidates: list[tuple[float, float, str]],
    obstacles: list[dict[str, float]],
    clearance: float,
) -> tuple[float, float, dict[str, Any]]:
    if not candidates:
        raise ValueError("no placement candidates were generated")
    scored = []
    for x, y, label in candidates:
        score = _point_clearance_score((x, y), obstacles, clearance)
        scored.append((score, x, y, label))
    scored.sort(key=lambda row: row[0])
    score, x, y, label = scored[0]
    return x, y, {
        "candidate_strategy": "minimum_obstacle_score",
        "selected_candidate": label,
        "selected_score": score,
        "candidates": [
            {"label": label_i, "view_mm": [x_i, y_i], "score": score_i}
            for score_i, x_i, y_i, label_i in scored
        ],
    }


def _semantic_candidates(
    kind: str,
    p: tuple[float, float],
    q: tuple[float, float],
    offset_view: float,
    evidence: dict[str, Any],
) -> list[tuple[float, float, str]]:
    mx, my = (p[0] + q[0]) / 2.0, (p[1] + q[1]) / 2.0
    gap = max(abs(offset_view), 4.0)
    if kind == "line_to_line":
        ux, uy = evidence.get("line_direction", [1.0, 0.0]) if evidence.get("line_direction") else (1.0, 0.0)
        return [(mx, my, "between_common_span"), (mx + ux * gap, my + uy * gap, "between_shift_plus"), (mx - ux * gap, my - uy * gap, "between_shift_minus")]
    if kind == "angle":
        vertex = tuple(evidence.get("vertex_view_mm", [mx, my]))
        bx, by = evidence.get("bisector", [1.0, 0.0])
        radius = float(evidence.get("witness_radius_view_mm", max(math.dist(p, vertex), 8.0)))
        base = radius + gap
        return [
            (vertex[0] + bx * base, vertex[1] + by * base, "inside_angle_bisector"),
            (vertex[0] - bx * base, vertex[1] - by * base, "opposite_angle_sector"),
            (vertex[0] + by * base, vertex[1] - bx * base, "rotated_sector_plus"),
            (vertex[0] - by * base, vertex[1] + bx * base, "rotated_sector_minus"),
        ]
    if kind in {"radius", "diameter"}:
        c = tuple(evidence.get("center_view_mm", [mx, my]))
        r = float(evidence.get("radius_mm", max(math.dist(p, q), 1.0)))
        d = r + gap
        inv = 1.0 / math.sqrt(2.0)
        dirs = [(1,0,"east"),(-1,0,"west"),(0,1,"north"),(0,-1,"south"),(inv,inv,"north_east"),(-inv,inv,"north_west"),(inv,-inv,"south_east"),(-inv,-inv,"south_west")]
        return [(c[0] + dx*d, c[1] + dy*d, label) for dx,dy,label in dirs]
    if kind == "horizontal_distance":
        y_hi = max(p[1], q[1]) + gap
        y_lo = min(p[1], q[1]) - gap
        return [(mx, y_hi, "above_features"), (mx, y_lo, "below_features"), (mx, my, "between_features")]
    if kind == "vertical_distance":
        x_hi = max(p[0], q[0]) + gap
        x_lo = min(p[0], q[0]) - gap
        return [(x_hi, my, "right_of_features"), (x_lo, my, "left_of_features"), (mx, my, "between_features")]
    if kind == "center_to_line":
        dx, dy = q[0] - p[0], q[1] - p[1]
        norm = math.hypot(dx, dy)
        if norm <= _EPS:
            return [(mx + gap, my + gap, "fallback_diagonal")]
        nx, ny = -dy/norm, dx/norm
        return [(mx + nx*gap, my + ny*gap, "side_plus"), (mx - nx*gap, my - ny*gap, "side_minus"), (mx,my,"between_center_and_line")]
    if kind in {"center_to_center", "distance", "aligned_distance", "line_length"}:
        dx, dy = q[0] - p[0], q[1] - p[1]
        norm = math.hypot(dx, dy)
        if norm <= _EPS:
            return [(mx + gap, my + gap, "fallback_diagonal")]
        nx, ny = -dy/norm, dx/norm
        return [(mx + nx*gap, my + ny*gap, "normal_plus"), (mx - nx*gap, my - ny*gap, "normal_minus"), (mx,my,"between_features")]
    return [(mx, my, "midpoint")]

def _auto_position(
    kind: str,
    p: tuple[float, float],
    q: tuple[float, float],
    offset_sheet_mm: float,
    scale: float,
    placement_mode: str = "smart",
    line_direction: Optional[tuple[float, float]] = None,
    anchor_evidence: Optional[dict[str, Any]] = None,
    obstacles: Optional[list[dict[str, float]]] = None,
) -> tuple[float, float, dict[str, Any]]:
    """Choose a semantic drafting position and score alternatives against obstacles."""
    mode = str(placement_mode or "smart").strip().lower().replace("-", "_")
    if mode not in {"smart", "between", "outside"}:
        raise ValueError("placement_mode must be 'smart', 'between', or 'outside'")
    mx, my = (p[0] + q[0]) / 2.0, (p[1] + q[1]) / 2.0
    offset = float(offset_sheet_mm) / float(scale)
    evidence = dict(anchor_evidence or {})
    if line_direction is not None:
        evidence["line_direction"] = list(line_direction)
    if mode == "between":
        return mx, my, {
            "strategy": "forced_between_feature_anchors",
            "midpoint_view_mm": [mx, my],
            "applied_offset_sheet_mm": 0.0,
        }
    obs = list(obstacles or [])
    candidates = _semantic_candidates(kind, p, q, offset, evidence)
    # outside mode removes explicit in-between choices when available.
    if mode == "outside":
        filtered = [c for c in candidates if "between" not in c[2]]
        if filtered:
            candidates = filtered
    clearance = max(1.5, 2.0 / max(scale, _EPS))
    vx, vy, scoring = _candidate_best(candidates, obs, clearance)
    scoring.update({
        "strategy": f"semantic_{kind}_placement",
        "midpoint_view_mm": [mx, my],
        "placement_mode": mode,
        "obstacle_count": len(obs),
        "clearance_view_mm": clearance,
        "applied_offset_sheet_mm": offset_sheet_mm,
    })
    return vx, vy, scoring


def _postplacement_collision_check(
    application: Any,
    view: Any,
    dim: Any,
    dimension_index: int,
    geometry_obstacles: list[dict[str, float]],
    margin: float = 0.5,
) -> dict[str, Any]:
    box = _boundary(application, dim)
    if box is None:
        return {"verified": False, "collision_free": None, "boundary": None, "collisions": []}
    subject = {k: float(box[k]) for k in ("xmin", "xmax", "ymin", "ymax")}
    existing = _existing_dimension_obstacles(application, view, exclude_index=dimension_index)
    collisions = []
    for label, obstacle_list in (("geometry", geometry_obstacles), ("dimension", existing)):
        for i, obstacle in enumerate(obstacle_list):
            if _bbox_intersects(subject, obstacle, margin=margin):
                collisions.append({"type": label, "index": i, "bbox": obstacle})
    return {
        "verified": True,
        "collision_free": not collisions,
        "boundary": box,
        "collision_margin_view_mm": margin,
        "collisions": collisions,
    }


def _try_collision_reposition(
    application: Any,
    view: Any,
    dim: Any,
    dimension_index: int,
    candidates: list[dict[str, Any]],
    geometry_obstacles: list[dict[str, float]],
) -> dict[str, Any]:
    attempts: list[dict[str, Any]] = []
    # Candidates are already score-ordered. Try the next positions if the initial
    # CATIA boundary collides with other geometry or dimensions.
    for candidate in candidates[1:6]:
        point = candidate.get("view_mm")
        if not isinstance(point, list) or len(point) != 2:
            continue
        try:
            dim.MoveValue(float(point[0]), float(point[1]), 0, 0)
            check = _postplacement_collision_check(
                application, view, dim, dimension_index, geometry_obstacles
            )
            attempts.append({"candidate": candidate, "check": check, "error": None})
            if check.get("collision_free") is True:
                return {
                    "repositioned": True,
                    "selected_view_mm": [float(point[0]), float(point[1])],
                    "attempts": attempts,
                    "final_check": check,
                }
        except Exception as exc:
            attempts.append({"candidate": candidate, "check": None, "error": _format_error(exc)})
    return {"repositioned": False, "selected_view_mm": None, "attempts": attempts, "final_check": None}


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
    placement_mode: str,
    allow_view_bounds_fallback: bool,
    witness_points: Optional[list[float]],
    name: str,
    tolerance_upper: Optional[float],
    tolerance_lower: Optional[float],
    tolerance_display_mode: int,
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
        expected, anchor1, anchor2, unit, anchor_evidence = _expected_and_anchors(kind, a, b)
        cat_obj1, cat_obj2 = _catia_geometry_objects(kind, obj1, obj2)
    except Exception as primary_exc:
        if (
            kind not in {"horizontal_distance", "vertical_distance"}
            or not bool(allow_view_bounds_fallback)
        ):
            raise

        expected, anchor1, anchor2, fallback_evidence = _view_bounds_dimension_anchors(
            application, view, kind
        )
        fallback_mode = True
        obj1 = None
        obj2 = None
        idx1 = -1
        a = {
            "index": None,
            "name": "DrawingView.Size@lower_x/upper_x" if kind == "horizontal_distance" else "DrawingView.Size@lower_y/upper_y",
            "automation_type": "DrawingView.Size bounding-box fallback",
            "has_range_box": True,
            "range_box_view_mm": fallback_evidence["view_bounds_mm"],
            "location_kind": "point",
            "location_view_mm": list(anchor1),
            "radius_mm": None,
            "direction": None,
        }
        b = {
            "index": None,
            "name": "DrawingView.Size@upper_x" if kind == "horizontal_distance" else "DrawingView.Size@upper_y",
            "automation_type": "DrawingView.Size bounding-box fallback",
            "has_range_box": True,
            "range_box_view_mm": fallback_evidence["view_bounds_mm"],
            "location_kind": "point",
            "location_view_mm": list(anchor2),
            "radius_mm": None,
            "direction": None,
        }
        unit = "length_mm"
        anchor_evidence = {"strategy": "drawing_view_size_fallback"}
        cat_obj1 = None
        cat_obj2 = None

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
    # Snapshot obstacles before creating the new dimension. Geometry range boxes and
    # existing dimension value boxes are used to score semantic placement candidates.
    geometry_obstacles = _geometry_obstacles(view, application)
    existing_dimension_obstacles = _existing_dimension_obstacles(application, view)
    placement_obstacles = geometry_obstacles + existing_dimension_obstacles

    if fallback_mode:
        geometry_before = _safe_count(_safe_attr(view, "GeometricElements"))
        try:
            factory = view.Factory2D
            helper_points = [
                factory.CreatePoint(anchor1[0], anchor1[1]),
                factory.CreatePoint(anchor2[0], anchor2[1]),
            ]
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
            raise AnnotationOperationError(
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
            f"Original geometry-selection failure: {_format_error(primary_exc)}"
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
    placement_evidence: dict[str, Any]
    if position_x is None and position_y is None:
        line_direction = None
        if kind == "line_to_line" and a is not None and a.get("direction"):
            line_direction = tuple(float(v) for v in a["direction"])
        vx, vy, placement_evidence = _auto_position(
            kind, anchor1, anchor2, offset, tf["scale"],
            placement_mode=placement_mode,
            line_direction=line_direction,
            anchor_evidence=anchor_evidence,
            obstacles=placement_obstacles,
        )
        resolved_from = (
            "drawing_view_size_fallback"
            if fallback_mode
            else placement_evidence["strategy"]
        )
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
            placement_evidence = {"strategy": resolved_from}
        elif space == "view":
            vx, vy = px, py
            resolved_from = "explicit_view_coordinates"
            placement_evidence = {"strategy": resolved_from}
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

        # Verify the actual CATIA dimension boundary. If it still overlaps other
        # geometry/dimensions, try the next semantic candidate positions.
        collision_check = _postplacement_collision_check(
            application, view, dim, created_index, geometry_obstacles
        )
        collision_reposition = {"repositioned": False, "attempts": [], "final_check": None}
        if (
            position_x is None
            and position_y is None
            and collision_check.get("collision_free") is False
            and isinstance(placement_evidence.get("candidates"), list)
        ):
            collision_reposition = _try_collision_reposition(
                application, view, dim, created_index,
                placement_evidence["candidates"], geometry_obstacles
            )
            if collision_reposition.get("repositioned"):
                vx, vy = collision_reposition["selected_view_mm"]
                collision_check = collision_reposition["final_check"]
                try:
                    view.SaveEdition()
                except Exception:
                    pass
            elif collision_check.get("collisions"):
                warnings.append(
                    "Semantic placement completed, but no tested candidate was fully "
                    "collision-free against visible geometry/existing dimensions."
                )
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
    if 'collision_check' not in locals():
        collision_check = _postplacement_collision_check(
            application, view, dim, created_index, geometry_obstacles
        )
    if 'collision_reposition' not in locals():
        collision_reposition = {"repositioned": False, "attempts": [], "final_check": None}
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
        "associative_geometry": not fallback_mode,
        "support_mode": "drawing_view_size_hidden_helper_points" if fallback_mode else "original_drawing_geometry",
        "fallback_evidence": fallback_evidence,
        "expected_value": expected,
        "anchor_strategy": anchor_evidence,
        "catia_measured_value": actual,
        "value_unit": unit,
        "value_matches_independent_calculation": value_matches,
        "witness_points_view_mm": witnesses,
        "placement": {
            "resolved_from": resolved_from,
            "placement_mode": placement_mode,
            "semantic_placement": placement_evidence,
            "view_mm": [vx, vy],
            "sheet_mm": [sx, sy],
            "view_transform": tf,
            "value_boundary_box_view_mm": box,
            "verified": box is not None,
            "collision_check": collision_check,
            "collision_reposition": collision_reposition,
            "collision_free": collision_check.get("collision_free"),
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
        offset_mm: float = 0.0,
        placement_mode: str = "smart",
        allow_view_bounds_fallback: bool = False,
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
        view origin, rotation and scale. If position is omitted, semantic smart
        placement is used. For line_to_line, the value is placed between the two
        visible parallel lines, centred on their common visible span. offset_mm
        then slides it only along the lines instead of pushing it out of the gap.

        The old whole-view DrawingView.Size fallback is disabled by default.
        Set allow_view_bounds_fallback=True only when the caller intentionally
        wants the full visible view width/height rather than selected geometry.
        """
        try:
            application, document, sheet, view = _active_context(conn, view_name)
            data, warnings = _add_internal(
                application=application, document=document, sheet=sheet, view=view,
                kind=kind, element1=element1, element2=element2,
                position_x=position_x, position_y=position_y,
                position_space=position_space, offset_mm=offset_mm,
                placement_mode=placement_mode,
                allow_view_bounds_fallback=allow_view_bounds_fallback,
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
                        offset_mm=spec.get("offset_mm", 0.0),
                        placement_mode=spec.get("placement_mode", "smart"),
                        allow_view_bounds_fallback=spec.get("allow_view_bounds_fallback", False),
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
