"""Native CATIA V5 drafting dimension, section-view and GD&T MCP tools.

These tools target classic CATDrawing Automation.  ``catia_create_gdt_frame``
creates a native two-dimensional ``DrawingGDT``; it is intentionally *not*
reported as a semantic three-dimensional FT&A tolerance.
"""

from __future__ import annotations

import math
import re
import time
from typing import Any, Optional, Sequence, Tuple

from catia_mcp.connection import CATIAError


IMPLEMENTATION_VERSION = "drawing-annotation-tools-fixed-2026-08-05-v7"
_CATVB_SCRIPT_LANGUAGE = 1

# Zero-based CATIA enum values.
CAT_DIM_DISTANCE = 0
CAT_DIM_AUTO = 3
CAT_VIS_PROPERTY_NO_SHOW = 1

# DrawingViewGenerativeBehavior.DefineSectionView does *not* use integer
# enums for section/profile type.  The CAA Automation signature declares both
# values as CATBSTR and accepts exactly: SectionView/SectionCut and
# Offset/Aligned.  Passing zero-based integers can be accepted by late-bound
# COM without raising while leaving the new DrawingView non-generative.
# v4 additionally copies DrawingView.GenerativeLinks from the parent to the
# child after DefineSectionView and before the first Update.  v5 makes the
# section-profile coordinate system explicit, converts optional sheet-space
# inputs to the parent-view axis system, rejects profiles that do not cross
# the parent view, and can marshal CATSafeArrayVariant explicitly.  v6 replaces
# unreliable Python-list DrawingView.Size handling with CATIA-side
# SystemService.Evaluate.  v7 resolves the live root Application through the
# same conn.connect(visible=True) entry point proven by drafting.py v7 and makes
# a disabled profile-intersection precheck a real bypass instead of an implicit
# dependency on DrawingView.Size.
SECTION_TYPE_BSTRS = {"SectionView", "SectionCut"}
PROFILE_TYPE_BSTRS = {"Offset", "Aligned"}

GDT_SYMBOL_CODES = {
    "straightness": 1,
    "flatness": 2,
    "circularity": 3,
    "roundness": 3,
    "cylindricity": 4,
    "line_profile": 5,
    "profile_of_line": 5,
    "surface_profile": 6,
    "profile_of_surface": 6,
    "angularity": 7,
    "perpendicularity": 8,
    "parallelism": 9,
    "position": 10,
    "true_position": 10,
    "concentricity": 11,
    "coaxiality": 11,
    "symmetry": 12,
    "circular_runout": 13,
    "total_runout": 14,
}


class CapabilityUnavailableError(RuntimeError):
    """Raised when the installed CATIA drafting interface is unavailable."""


class AnnotationOperationError(RuntimeError):
    """Operation failure with rollback/status evidence for the MCP response."""

    def __init__(
        self,
        message: str,
        *,
        data: Any = None,
        warnings: Optional[list[str]] = None,
        status: str = "error",
    ) -> None:
        super().__init__(message)
        self.data = data
        self.warnings = list(warnings or [])
        self.status = status


def _success(data: Any, warnings: Optional[list[str]] = None) -> dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": True,
        "status": "success_with_warnings" if warning_list else "success",
        "data": data,
        "warnings": warning_list,
    }


def _error(
    operation: str,
    exc: BaseException,
    *,
    data: Any = None,
    warnings: Optional[list[str]] = None,
    status: Optional[str] = None,
) -> dict[str, Any]:
    operation_data = data
    warning_list = list(warnings or [])
    resolved_status = status

    if isinstance(exc, AnnotationOperationError):
        if operation_data is None:
            operation_data = exc.data
        warning_list.extend(exc.warnings)
        if resolved_status is None:
            resolved_status = exc.status

    if resolved_status is None:
        resolved_status = (
            "capability_unavailable"
            if isinstance(exc, CapabilityUnavailableError)
            else "error"
        )

    result: dict[str, Any] = {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": False,
        "status": resolved_status,
        "operation": operation,
        "error": str(exc),
        "error_type": type(exc).__name__,
        "warnings": warning_list,
    }
    if operation_data is not None:
        result["data"] = operation_data
    if isinstance(exc, CapabilityUnavailableError):
        result["capability"] = {
            "name": "CATIA V5 2D Drafting Automation",
            "required_interface": "DrawingView.GDTs / DrawingGDTs.Add",
            "semantic_3d_fta": False,
        }
    return result


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


def _catia_application(
    ctx: Any,
    document: Any = None,
) -> tuple[Any, dict[str, Any]]:
    """Resolve the live root CATIA Application with auditable evidence.

    The connection object's ``connect(visible=True)`` method is the primary
    path because drafting.py v7 has already proven that object can execute
    ``SystemService.Evaluate`` in the same MCP runtime.  Attribute probing is
    retained only as a compatibility fallback.
    """
    conn = _safe_attr(ctx, "conn")
    attempts: list[dict[str, Any]] = []

    if conn is not None:
        try:
            connector = getattr(conn, "connect")
            application = connector(visible=True)
            _ = application.SystemService
            attempts.append(
                {
                    "method": "conn.connect(visible=True)",
                    "succeeded": True,
                    "system_service_available": True,
                    "error": None,
                }
            )
            return application, {
                "resolved": True,
                "method": "conn.connect(visible=True)",
                "system_service_available": True,
                "attempts": attempts,
            }
        except Exception as exc:
            attempts.append(
                {
                    "method": "conn.connect(visible=True)",
                    "succeeded": False,
                    "system_service_available": False,
                    "error": str(exc),
                }
            )

    legacy_candidates = [
        ("conn.app", _safe_attr(conn, "app")),
        ("conn.application", _safe_attr(conn, "application")),
        ("conn._app", _safe_attr(conn, "_app")),
        (
            "document.Application",
            _safe_attr(document, "Application") if document is not None else None,
        ),
    ]
    for method, candidate in legacy_candidates:
        if candidate is None:
            continue
        try:
            application = candidate() if callable(candidate) else candidate
            _ = application.SystemService
            attempts.append(
                {
                    "method": method,
                    "succeeded": True,
                    "system_service_available": True,
                    "error": None,
                }
            )
            return application, {
                "resolved": True,
                "method": method,
                "system_service_available": True,
                "attempts": attempts,
                "compatibility_fallback_used": True,
            }
        except Exception as exc:
            attempts.append(
                {
                    "method": method,
                    "succeeded": False,
                    "system_service_available": False,
                    "error": str(exc),
                }
            )

    raise AnnotationOperationError(
        "Cannot resolve the live root CATIA Application for DrawingView.Size readback.",
        data={
            "failure_stage": "B_application_resolution",
            "application_resolution": {
                "resolved": False,
                "method": None,
                "system_service_available": False,
                "attempts": attempts,
            },
        },
    )


def _evaluate(
    application: Any,
    script: str,
    function_name: str,
    parameters: list[Any],
) -> Any:
    try:
        system_service = application.SystemService
    except Exception as exc:
        raise RuntimeError(f"Cannot access CATIA SystemService: {exc}") from exc

    try:
        return system_service.Evaluate(
            script,
            _CATVB_SCRIPT_LANGUAGE,
            function_name,
            parameters,
        )
    except Exception as exc:
        raise RuntimeError(
            f"SystemService.Evaluate failed for {function_name}: {exc}"
        ) from exc


def _numeric_sequence(value: Any, expected_length: int) -> list[float]:
    try:
        sequence = list(value)
    except Exception as exc:
        raise RuntimeError(f"CATIA did not return an array: {exc}") from exc
    if len(sequence) != expected_length:
        raise RuntimeError(
            f"CATIA returned {len(sequence)} values; {expected_length} were required."
        )
    result: list[float] = []
    for index, item in enumerate(sequence):
        try:
            number = float(item)
        except (TypeError, ValueError) as exc:
            raise RuntimeError(
                f"CATIA array element {index} is not numeric."
            ) from exc
        if not math.isfinite(number):
            raise RuntimeError(
                f"CATIA array element {index} is not finite."
            )
        result.append(number)
    return result


def _nonempty_text(value: Any, parameter_name: str) -> str:
    text = str(value).strip()
    if not text:
        raise ValueError(f"{parameter_name} cannot be empty.")
    return text


def _finite_float(value: Any, parameter_name: str) -> float:
    if isinstance(value, bool):
        raise ValueError(f"{parameter_name} must be a finite number.")
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{parameter_name} must be a finite number.") from exc
    if not math.isfinite(number):
        raise ValueError(f"{parameter_name} must be finite.")
    return number


def _point2(value: Sequence[Any], parameter_name: str) -> tuple[float, float]:
    if isinstance(value, (str, bytes)):
        raise ValueError(f"{parameter_name} must contain exactly two numbers.")
    try:
        coordinates = list(value)
    except TypeError as exc:
        raise ValueError(
            f"{parameter_name} must contain exactly two numbers."
        ) from exc
    if len(coordinates) != 2:
        raise ValueError(f"{parameter_name} must contain exactly two numbers.")
    return (
        _finite_float(coordinates[0], f"{parameter_name}[0]"),
        _finite_float(coordinates[1], f"{parameter_name}[1]"),
    )


def _active_drawing_document(ctx: Any) -> Any:
    conn = ctx.conn
    conn.ensure_connected()

    document = None
    getter = _safe_attr(conn, "get_active_drawing_document", None)
    if callable(getter):
        document = getter()
    else:
        legacy_getter = _safe_attr(ctx, "get_active_drawing", None)
        if callable(legacy_getter):
            document = legacy_getter()

    if document is None:
        raise CATIAError(
            "The connection does not provide get_active_drawing_document()."
        )
    try:
        if int(document.Sheets.Count) < 1:
            raise CATIAError("The active CATDrawing contains no sheets.")
    except CATIAError:
        raise
    except Exception as exc:
        raise CATIAError("The active document is not a classic CATDrawing.") from exc
    return document


def _active_sheet(document: Any) -> Any:
    try:
        return document.Sheets.ActiveSheet
    except Exception as exc:
        raise CATIAError("Cannot access DrawingDocument.Sheets.ActiveSheet.") from exc


def _drawing_view(sheet: Any, view_name: str) -> Any:
    requested_name = _nonempty_text(view_name, "view_name")
    try:
        return sheet.Views.Item(requested_name)
    except Exception as exc:
        raise LookupError(f"Drawing view '{requested_name}' was not found.") from exc


def _hide_objects(document: Any, objects: Sequence[Any]) -> tuple[bool, Optional[str]]:
    selection = None
    try:
        selection = document.Selection
        selection.Clear()
        for item in objects:
            selection.Add(item)
        selection.VisProperties.SetShow(CAT_VIS_PROPERTY_NO_SHOW)
        return True, None
    except Exception as exc:
        return False, str(exc)
    finally:
        if selection is not None:
            try:
                selection.Clear()
            except Exception:
                pass


def _delete_objects(document: Any, objects: Sequence[Any]) -> dict[str, Any]:
    details = {
        "attempted": bool(objects),
        "requested_count": len(objects),
        "succeeded": True if not objects else None,
        "error": None,
    }
    if not objects:
        return details

    selection = None
    try:
        selection = document.Selection
        selection.Clear()
        for item in objects:
            selection.Add(item)
        selection.Delete()
        details["succeeded"] = True
    except Exception as exc:
        details["succeeded"] = False
        details["error"] = str(exc)
    finally:
        if selection is not None:
            try:
                selection.Clear()
            except Exception as exc:
                details["selection_clear_error"] = str(exc)
    return details


def _save_view_edition(view: Any, warnings: list[str]) -> None:
    try:
        view.SaveEdition()
    except Exception as exc:
        warnings.append(f"DrawingView.SaveEdition could not be completed: {exc}")


def _update_drawing(sheet: Any, document: Any, warnings: list[str]) -> None:
    updated = False
    for target_name, target in (("sheet", sheet), ("document", document)):
        method = _safe_attr(target, "Update", None)
        if not callable(method):
            continue
        try:
            method()
            updated = True
            break
        except Exception as exc:
            warnings.append(f"{target_name}.Update failed: {exc}")
    if not updated:
        warnings.append("Neither DrawingSheet.Update nor DrawingDocument.Update succeeded.")


def _view_position(view: Any) -> tuple[float, float]:
    for x_name, y_name in (("x", "y"), ("xAxisData", "yAxisData")):
        try:
            return float(getattr(view, x_name)), float(getattr(view, y_name))
        except Exception:
            continue
    raise CATIAError("Cannot read the drawing view position.")


def _set_view_position(
    view: Any,
    x: float,
    y: float,
    tolerance_mm: float = 0.01,
) -> dict[str, Any]:
    attempts: list[dict[str, Any]] = []
    for x_name, y_name in (("x", "y"), ("xAxisData", "yAxisData")):
        try:
            setattr(view, x_name, float(x))
            setattr(view, y_name, float(y))
            actual_x = float(getattr(view, x_name))
            actual_y = float(getattr(view, y_name))
            verified = bool(
                abs(actual_x - x) <= tolerance_mm
                and abs(actual_y - y) <= tolerance_mm
            )
            attempts.append(
                {
                    "method": f"DrawingView.{x_name}/{y_name}",
                    "succeeded": True,
                    "actual_x_mm": actual_x,
                    "actual_y_mm": actual_y,
                    "verified": verified,
                    "error": None,
                }
            )
            if verified:
                return {
                    "requested_x_mm": float(x),
                    "requested_y_mm": float(y),
                    "actual_x_mm": actual_x,
                    "actual_y_mm": actual_y,
                    "tolerance_mm": tolerance_mm,
                    "selected_method": attempts[-1]["method"],
                    "verified": True,
                    "attempts": attempts,
                }
        except Exception as exc:
            attempts.append(
                {
                    "method": f"DrawingView.{x_name}/{y_name}",
                    "succeeded": False,
                    "verified": False,
                    "error": str(exc),
                }
            )
    raise AnnotationOperationError(
        "Cannot position the new drawing view.",
        data={
            "requested_x_mm": float(x),
            "requested_y_mm": float(y),
            "tolerance_mm": tolerance_mm,
            "attempts": attempts,
        },
    )


def _normalise_profile(
    cut_line_points: Sequence[Sequence[Any]],
) -> tuple[tuple[float, ...], list[list[float]]]:
    if isinstance(cut_line_points, (str, bytes)):
        raise ValueError("cut_line_points must be a list of [x, y] points.")
    try:
        raw_points = list(cut_line_points)
    except TypeError as exc:
        raise ValueError("cut_line_points must be a list of [x, y] points.") from exc
    if len(raw_points) < 2:
        raise ValueError("cut_line_points requires at least two points.")

    points = [_point2(point, f"cut_line_points[{index}]") for index, point in enumerate(raw_points)]
    for index in range(1, len(points)):
        if points[index] == points[index - 1]:
            raise ValueError(
                f"cut_line_points[{index - 1}] and cut_line_points[{index}] "
                "cannot be identical."
            )

    flattened = tuple(value for point in points for value in point)
    return flattened, [[point[0], point[1]] for point in points]

def _normalise_coordinate_mode(value: str) -> str:
    aliases = {
        "local": "parent_local",
        "parent": "parent_local",
        "parent_view": "parent_local",
        "view_local": "parent_local",
        "sheet_space": "sheet",
        "paper": "sheet",
    }
    key = str(value or "parent_local").strip().lower().replace("-", "_").replace(" ", "_")
    key = aliases.get(key, key)
    if key not in {"parent_local", "sheet"}:
        raise ValueError("coordinate_mode must be 'parent_local' or 'sheet'.")
    return key


def _normalise_profile_marshalling(value: str) -> str:
    aliases = {
        "variant": "explicit_variant",
        "safe_array_variant": "explicit_variant",
        "catsafearrayvariant": "explicit_variant",
        "tuple": "python_tuple",
        "legacy": "python_tuple",
    }
    key = str(value or "explicit_variant").strip().lower().replace("-", "_").replace(" ", "_")
    key = aliases.get(key, key)
    if key not in {"explicit_variant", "python_tuple"}:
        raise ValueError(
            "profile_marshalling must be 'explicit_variant' or 'python_tuple'."
        )
    return key


def _view_axis_contract(view: Any) -> dict[str, Any]:
    """Read the parent-view axis-system placement in sheet coordinates.

    CAA V5 documents xAxisData/yAxisData as the drawing-view coordinate-system
    origin expressed in the sheet coordinate system. DrawingView.Angle is the
    counterclockwise rotation from the sheet axis system to the view axis
    system, in radians.
    """
    attempts: list[dict[str, Any]] = []

    def read_number(attribute: str) -> float:
        try:
            raw = getattr(view, attribute)
            raw = raw() if callable(raw) else raw
            value = float(raw)
            if not math.isfinite(value):
                raise ValueError(f"{attribute} is not finite.")
            attempts.append(
                {"attribute": attribute, "succeeded": True, "value": value, "error": None}
            )
            return value
        except Exception as exc:
            attempts.append(
                {"attribute": attribute, "succeeded": False, "value": None, "error": str(exc)}
            )
            raise AnnotationOperationError(
                f"Cannot read parent DrawingView.{attribute}.",
                data={"axis_contract_attempts": attempts},
            ) from exc

    origin_x = read_number("xAxisData")
    origin_y = read_number("yAxisData")
    angle = read_number("Angle")

    scale: Optional[float] = None
    for attribute in ("Scale2", "Scale"):
        try:
            raw = getattr(view, attribute)
            raw = raw() if callable(raw) else raw
            candidate = float(raw)
            if math.isfinite(candidate) and candidate > 0.0:
                scale = candidate
                attempts.append(
                    {"attribute": attribute, "succeeded": True, "value": scale, "error": None}
                )
                break
        except Exception as exc:
            attempts.append(
                {"attribute": attribute, "succeeded": False, "value": None, "error": str(exc)}
            )

    return {
        "available": True,
        "origin_sheet_mm": [origin_x, origin_y],
        "angle_rad": angle,
        "angle_deg": math.degrees(angle),
        "scale": scale,
        "angle_convention": "counterclockwise_from_sheet_axis_to_view_axis",
        "attempts": attempts,
    }


def _sheet_to_parent_local_point(
    point: Sequence[Any],
    axis_contract: dict[str, Any],
) -> tuple[float, float]:
    x_sheet, y_sheet = _point2(point, "sheet_point")
    origin_x, origin_y = axis_contract["origin_sheet_mm"]
    angle = float(axis_contract["angle_rad"])
    dx = x_sheet - float(origin_x)
    dy = y_sheet - float(origin_y)
    cosine = math.cos(angle)
    sine = math.sin(angle)
    # Inverse of local->sheet rotation R(angle).
    x_local = cosine * dx + sine * dy
    y_local = -sine * dx + cosine * dy
    return x_local, y_local


def _parent_local_to_sheet_point(
    point: Sequence[Any],
    axis_contract: dict[str, Any],
) -> tuple[float, float]:
    x_local, y_local = _point2(point, "parent_local_point")
    origin_x, origin_y = axis_contract["origin_sheet_mm"]
    angle = float(axis_contract["angle_rad"])
    cosine = math.cos(angle)
    sine = math.sin(angle)
    x_sheet = float(origin_x) + cosine * x_local - sine * y_local
    y_sheet = float(origin_y) + sine * x_local + cosine * y_local
    return x_sheet, y_sheet


def _bbox_from_points(points: Sequence[Sequence[Any]]) -> dict[str, float]:
    numeric = [_point2(point, "bbox_point") for point in points]
    xs = [point[0] for point in numeric]
    ys = [point[1] for point in numeric]
    return {
        "xmin": min(xs),
        "xmax": max(xs),
        "ymin": min(ys),
        "ymax": max(ys),
        "width_mm": max(xs) - min(xs),
        "height_mm": max(ys) - min(ys),
    }


def _parent_local_bbox(
    application: Any,
    view: Any,
    axis_contract: dict[str, Any],
) -> dict[str, Any]:
    observed = _view_size(application, view)
    if observed["width_mm"] <= 1e-6 or observed["height_mm"] <= 1e-6:
        raise AnnotationOperationError(
            "The parent generative view has a reliably read but empty DrawingView.Size.",
            data={
                "failure_stage": "C_parent_view_bbox_read",
                "parent_view_size": observed,
            },
        )
    sheet_corners = [
        [observed["xmin"], observed["ymin"]],
        [observed["xmin"], observed["ymax"]],
        [observed["xmax"], observed["ymin"]],
        [observed["xmax"], observed["ymax"]],
    ]
    local_corners = [
        list(_sheet_to_parent_local_point(point, axis_contract))
        for point in sheet_corners
    ]
    local_bbox = _bbox_from_points(local_corners)
    return {
        "observed_drawing_view_size": observed,
        "observed_size_assumption": "sheet_coordinate_bounding_box",
        "sheet_corners_mm": sheet_corners,
        "parent_local_corners_mm": local_corners,
        "parent_local_bbox_mm": local_bbox,
    }


def _point_inside_bbox(
    point: tuple[float, float],
    bbox: dict[str, float],
    margin_mm: float,
) -> bool:
    x, y = point
    return bool(
        bbox["xmin"] - margin_mm <= x <= bbox["xmax"] + margin_mm
        and bbox["ymin"] - margin_mm <= y <= bbox["ymax"] + margin_mm
    )


def _segment_intersects_bbox(
    start: tuple[float, float],
    end: tuple[float, float],
    bbox: dict[str, float],
    margin_mm: float = 1e-6,
) -> bool:
    """Liang-Barsky segment/axis-aligned-rectangle intersection test."""
    if _point_inside_bbox(start, bbox, margin_mm) or _point_inside_bbox(end, bbox, margin_mm):
        return True

    xmin = bbox["xmin"] - margin_mm
    xmax = bbox["xmax"] + margin_mm
    ymin = bbox["ymin"] - margin_mm
    ymax = bbox["ymax"] + margin_mm
    x0, y0 = start
    x1, y1 = end
    dx = x1 - x0
    dy = y1 - y0
    p = (-dx, dx, -dy, dy)
    q = (x0 - xmin, xmax - x0, y0 - ymin, ymax - y0)
    lower = 0.0
    upper = 1.0
    for p_value, q_value in zip(p, q):
        if abs(p_value) <= 1e-15:
            if q_value < 0.0:
                return False
            continue
        ratio = q_value / p_value
        if p_value < 0.0:
            lower = max(lower, ratio)
        else:
            upper = min(upper, ratio)
        if lower > upper:
            return False
    return True


def _profile_intersection_evidence(
    points: Sequence[Sequence[Any]],
    bbox: dict[str, float],
    margin_mm: float = 1e-6,
) -> dict[str, Any]:
    numeric = [_point2(point, "profile_point") for point in points]
    segments: list[dict[str, Any]] = []
    intersects = False
    for index in range(len(numeric) - 1):
        start = numeric[index]
        end = numeric[index + 1]
        segment_intersects = _segment_intersects_bbox(start, end, bbox, margin_mm)
        segments.append(
            {
                "index": index,
                "start_mm": list(start),
                "end_mm": list(end),
                "intersects": segment_intersects,
            }
        )
        intersects = intersects or segment_intersects
    return {
        "checked": True,
        "margin_mm": margin_mm,
        "profile_bbox_mm": _bbox_from_points(numeric),
        "parent_local_bbox_mm": dict(bbox),
        "segments": segments,
        "intersects_parent_bbox": intersects,
    }


def _prepare_section_profile(
    application: Any,
    parent: Any,
    input_points: Sequence[Sequence[Any]],
    coordinate_mode: str,
    require_intersection: bool,
) -> dict[str, Any]:
    """Normalize a section profile and optionally prove parent-view overlap.

    When ``require_intersection`` is false, DrawingView.Size is deliberately
    not read.  This is a true bypass intended for diagnosis or environments
    where the reliable in-process Size reader is temporarily unavailable.
    """
    mode = _normalise_coordinate_mode(coordinate_mode)
    axis_contract = _view_axis_contract(parent)

    normalized_input = [list(_point2(point, "cut_line_point")) for point in input_points]
    if mode == "sheet":
        local_points = [
            list(_sheet_to_parent_local_point(point, axis_contract))
            for point in normalized_input
        ]
    else:
        local_points = [list(point) for point in normalized_input]

    if require_intersection:
        if application is None:
            raise AnnotationOperationError(
                "Profile intersection verification requires a live CATIA Application.",
                data={
                    "failure_stage": "C_parent_view_bbox_read",
                    "profile_intersection_precheck": {
                        "checked": False,
                        "skipped": False,
                        "reason": "application_unavailable",
                    },
                },
            )
        local_bbox_evidence = _parent_local_bbox(application, parent, axis_contract)
        intersection = _profile_intersection_evidence(
            local_points,
            local_bbox_evidence["parent_local_bbox_mm"],
        )
    else:
        local_bbox_evidence = {
            "read_skipped": True,
            "reason": "disabled_by_caller",
            "observed_drawing_view_size": None,
            "parent_local_bbox_mm": None,
        }
        intersection = {
            "checked": False,
            "skipped": True,
            "reason": "disabled_by_caller",
            "profile_bbox_mm": _bbox_from_points(local_points),
            "parent_local_bbox_mm": None,
            "segments": [],
            "intersects_parent_bbox": None,
        }

    evidence = {
        "coordinate_mode_requested": str(coordinate_mode),
        "coordinate_mode_resolved": mode,
        "input_points_mm": normalized_input,
        "parent_local_points_mm": local_points,
        "parent_local_points_as_sheet_mm": [
            list(_parent_local_to_sheet_point(point, axis_contract))
            for point in local_points
        ],
        "axis_contract": axis_contract,
        "parent_bbox": local_bbox_evidence,
        "intersection": intersection,
        "require_intersection": bool(require_intersection),
        "coordinate_contract": (
            "DefineSectionView receives coordinates expressed in the parent "
            "drawing-view axis system. coordinate_mode='sheet' is converted "
            "with xAxisData/yAxisData and DrawingView.Angle before COM invocation."
        ),
    }
    if require_intersection and not intersection["intersects_parent_bbox"]:
        raise AnnotationOperationError(
            "The section profile does not intersect the parent view local bounding box.",
            data={
                "failure_stage": "C_profile_intersection_precheck",
                "profile_precheck": evidence,
            },
        )

    evidence["flattened_parent_local_profile"] = [
        value for point in local_points for value in point
    ]
    return evidence


def _marshal_section_profile(
    profile: Sequence[float],
    strategy: str,
) -> tuple[Any, dict[str, Any]]:
    values = [float(value) for value in profile]
    if not values or len(values) % 2 != 0:
        raise ValueError("Section profile must contain an even number of coordinates.")
    if not all(math.isfinite(value) for value in values):
        raise ValueError("Section profile coordinates must be finite.")

    resolved = _normalise_profile_marshalling(strategy)
    base_evidence: dict[str, Any] = {
        "requested_strategy": str(strategy),
        "resolved_strategy": resolved,
        "profile_length": len(values),
        "point_count": len(values) // 2,
        "values": list(values),
        "element_python_types": [type(value).__name__ for value in values],
        "expected_catia_type": "CATSafeArrayVariant",
    }
    if resolved == "python_tuple":
        payload = tuple(values)
        base_evidence.update(
            {
                "payload_python_type": type(payload).__name__,
                "variant_vt": None,
                "variant_vt_expression": None,
                "actual_com_vt_verified": False,
                "note": (
                    "Legacy pywin32 automatic tuple marshalling; the actual COM "
                    "SAFEARRAY element type cannot be proven from Python alone."
                ),
            }
        )
        return payload, base_evidence

    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore
    except Exception as exc:
        raise AnnotationOperationError(
            "Explicit CATSafeArrayVariant marshalling requires pywin32.",
            data={"profile_marshalling": base_evidence, "import_error": str(exc)},
        ) from exc

    variant_vt = int(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT)
    payload = VARIANT(variant_vt, values)
    observed_vt = getattr(payload, "varianttype", None)
    base_evidence.update(
        {
            "payload_python_type": type(payload).__name__,
            "variant_vt": int(observed_vt) if observed_vt is not None else variant_vt,
            "variant_vt_expression": "VT_ARRAY | VT_VARIANT",
            "configured_vt_array": bool(variant_vt & int(pythoncom.VT_ARRAY)),
            "configured_element_vt": int(pythoncom.VT_VARIANT),
            "payload_varianttype_verified": observed_vt is not None,
            "actual_com_vt_verified": False,
        }
    )
    return payload, base_evidence



def _unique_view_name(views: Any, requested_name: str) -> str:
    base = re.sub(r"[^A-Za-z0-9_.-]+", "_", requested_name).strip("_")
    if not base:
        base = "SectionView"
    existing: set[str] = set()
    count = _safe_count(views) or 0
    for index in range(1, count + 1):
        try:
            existing.add(str(views.Item(index).Name).casefold())
        except Exception:
            continue
    if base.casefold() not in existing:
        return base
    suffix = 2
    while f"{base}_{suffix}".casefold() in existing:
        suffix += 1
    return f"{base}_{suffix}"


def _remove_collection_item(
    collection: Any,
    *,
    index: Optional[int] = None,
    name: str = "",
) -> dict[str, Any]:
    before_count = _safe_count(collection)
    attempts: list[dict[str, Any]] = []
    identifiers: list[Any] = []
    if index is not None:
        identifiers.append(index)
    if name:
        identifiers.append(name)

    for identifier in identifiers:
        try:
            collection.Remove(identifier)
            after_count = _safe_count(collection)
            verified = bool(
                before_count is None
                or after_count is None
                or after_count == before_count - 1
            )
            attempts.append(
                {
                    "identifier": identifier,
                    "succeeded": True,
                    "count_after": after_count,
                    "verified": verified,
                    "error": None,
                }
            )
            if verified:
                return {
                    "attempted": True,
                    "succeeded": True,
                    "verified": True,
                    "count_before": before_count,
                    "count_after": after_count,
                    "selected_identifier": identifier,
                    "attempts": attempts,
                }
        except Exception as exc:
            attempts.append(
                {
                    "identifier": identifier,
                    "succeeded": False,
                    "verified": False,
                    "error": str(exc),
                }
            )

    return {
        "attempted": bool(identifiers),
        "succeeded": False,
        "verified": False,
        "count_before": before_count,
        "count_after": _safe_count(collection),
        "selected_identifier": None,
        "attempts": attempts,
    }


def _remove_view(views: Any, name: str, index: Optional[int]) -> dict[str, Any]:
    return _remove_collection_item(views, index=index, name=name)


def _is_generative(view: Any) -> Optional[bool]:
    value = _safe_attr(view, "IsGenerative", None)
    if value is None:
        return None
    try:
        return bool(value() if callable(value) else value)
    except Exception:
        return None


def _dimension_value_readback(dimension: Any) -> dict[str, Any]:
    attempts: list[dict[str, Any]] = []
    try:
        getter = _safe_attr(dimension, "GetValue", None)
        value_object = getter() if callable(getter) else getter
        if value_object is None:
            raise RuntimeError("DrawingDimension.GetValue returned no object.")
        value = float(value_object.Value)
        attempts.append(
            {
                "method": "DrawingDimension.GetValue().Value",
                "succeeded": True,
                "value": value,
                "error": None,
            }
        )
        return {
            "available": True,
            "value": value,
            "read_method": attempts[-1]["method"],
            "attempts": attempts,
        }
    except Exception as exc:
        attempts.append(
            {
                "method": "DrawingDimension.GetValue().Value",
                "succeeded": False,
                "value": None,
                "error": str(exc),
            }
        )
    return {
        "available": False,
        "value": None,
        "read_method": None,
        "attempts": attempts,
    }


def _view_size(application: Any, view: Any) -> dict[str, Any]:
    """Read DrawingView.Size without mistaking an untouched Python list for output.

    CATIA exposes Size as an output-array method.  The primary implementation
    executes that method inside CATIA through SystemService.Evaluate, matching
    the drafting.py v7 path already proven against real Front/Top/Right views.
    Direct COM fallbacks are accepted only when they return a non-zero box; an
    all-zero direct result is retained as diagnostic evidence, never treated as
    a verified empty view.
    """
    attempts: list[dict[str, Any]] = []
    suspicious_zero_results: list[dict[str, Any]] = []

    def build(
        values: Any,
        method: str,
        *,
        read_verified: bool,
    ) -> dict[str, Any]:
        numeric = [float(item) for item in list(values)]
        if len(numeric) != 4 or not all(math.isfinite(item) for item in numeric):
            raise RuntimeError("DrawingView.Size did not return four finite values.")
        xmin, xmax, ymin, ymax = numeric
        if xmax < xmin or ymax < ymin:
            raise RuntimeError("DrawingView.Size returned an inverted bounding box.")
        return {
            "xmin": xmin,
            "xmax": xmax,
            "ymin": ymin,
            "ymax": ymax,
            "width_mm": xmax - xmin,
            "height_mm": ymax - ymin,
            "read_method": method,
            "read_verified": bool(read_verified),
            "read_attempts": attempts,
        }

    script = (
        "Public Function MCP_GetDrawingViewSize(viewObject)\n"
        "    Dim values(3)\n"
        "    viewObject.Size values\n"
        "    MCP_GetDrawingViewSize = Array("
        "CDbl(values(0)), CDbl(values(1)), "
        "CDbl(values(2)), CDbl(values(3)))\n"
        "End Function"
    )
    if application is not None:
        try:
            values = _numeric_sequence(
                _evaluate(
                    application,
                    script,
                    "MCP_GetDrawingViewSize",
                    [view],
                ),
                4,
            )
            result = build(
                values,
                "SystemService.Evaluate.DrawingView.Size",
                read_verified=True,
            )
            attempts.append(
                {
                    "method": result["read_method"],
                    "succeeded": True,
                    "read_verified": True,
                    "values": values,
                    "error": None,
                }
            )
            result["read_attempts"] = attempts
            return result
        except Exception as exc:
            attempts.append(
                {
                    "method": "SystemService.Evaluate.DrawingView.Size",
                    "succeeded": False,
                    "read_verified": False,
                    "values": None,
                    "error": str(exc),
                }
            )
    else:
        attempts.append(
            {
                "method": "SystemService.Evaluate.DrawingView.Size",
                "succeeded": False,
                "read_verified": False,
                "values": None,
                "error": "Skipped because no live CATIA Application was resolved.",
            }
        )

    # Compatibility fallback: explicit ByRef SAFEARRAY.  A non-zero returned
    # box is usable; an unchanged all-zero payload is ambiguous and rejected.
    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore

        variant = VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
            [0.0, 0.0, 0.0, 0.0],
        )
        view.Size(variant)
        candidate = build(
            variant.value,
            "DrawingView.Size_typed_BYREF_VARIANT",
            read_verified=False,
        )
        values = [
            candidate["xmin"], candidate["xmax"],
            candidate["ymin"], candidate["ymax"],
        ]
        if candidate["width_mm"] > 1e-6 and candidate["height_mm"] > 1e-6:
            candidate["read_verified"] = True
            attempts.append(
                {
                    "method": candidate["read_method"],
                    "succeeded": True,
                    "read_verified": True,
                    "values": values,
                    "error": None,
                }
            )
            candidate["read_attempts"] = attempts
            return candidate
        suspicious_zero_results.append(
            {"method": candidate["read_method"], "values": values}
        )
        attempts.append(
            {
                "method": candidate["read_method"],
                "succeeded": False,
                "read_verified": False,
                "values": values,
                "error": "All-zero output is ambiguous outside SystemService.Evaluate.",
            }
        )
    except Exception as exc:
        attempts.append(
            {
                "method": "DrawingView.Size_typed_BYREF_VARIANT",
                "succeeded": False,
                "read_verified": False,
                "values": None,
                "error": str(exc),
            }
        )

    # Diagnostic fallbacks only.  A Python list can remain unchanged even when
    # COM returns successfully, which caused v5 to misclassify valid views.
    for method, action in (
        ("DrawingView.Size_return_value", lambda: view.Size()),
        ("DrawingView.Size_mutable_array", None),
    ):
        try:
            if action is not None:
                raw = action()
            else:
                values = [0.0, 0.0, 0.0, 0.0]
                returned = view.Size(values)
                raw = returned if returned is not None else values
            candidate = build(raw, method, read_verified=False)
            candidate_values = [
                candidate["xmin"], candidate["xmax"],
                candidate["ymin"], candidate["ymax"],
            ]
            if candidate["width_mm"] > 1e-6 and candidate["height_mm"] > 1e-6:
                candidate["read_verified"] = True
                attempts.append(
                    {
                        "method": method,
                        "succeeded": True,
                        "read_verified": True,
                        "values": candidate_values,
                        "error": None,
                    }
                )
                candidate["read_attempts"] = attempts
                return candidate
            suspicious_zero_results.append(
                {"method": method, "values": candidate_values}
            )
            attempts.append(
                {
                    "method": method,
                    "succeeded": False,
                    "read_verified": False,
                    "values": candidate_values,
                    "error": "All-zero direct-COM output was rejected as unverified.",
                }
            )
        except Exception as exc:
            attempts.append(
                {
                    "method": method,
                    "succeeded": False,
                    "read_verified": False,
                    "values": None,
                    "error": str(exc),
                }
            )

    raise AnnotationOperationError(
        "Could not reliably read DrawingView.Size; no verified bounding box is available.",
        data={
            "failure_stage": "C_parent_view_bbox_read",
            "size_read_verified": False,
            "read_attempts": attempts,
            "suspicious_zero_results": suspicious_zero_results,
        },
    )


def _section_reference_summary(section: Any, parent: Any) -> dict[str, Any]:
    expected_name = str(_safe_attr(parent, "Name", ""))
    parent_view_name: Optional[str] = None
    reference_view_name: Optional[str] = None
    errors: list[str] = []

    try:
        parent_view_name = str(section.GenerativeBehavior.ParentView.Name)
    except Exception as exc:
        errors.append(f"GenerativeBehavior.ParentView: {exc}")
    try:
        reference_view_name = str(section.ReferenceView.Name)
    except Exception as exc:
        errors.append(f"DrawingView.ReferenceView: {exc}")

    verified = bool(
        expected_name
        and expected_name in {parent_view_name, reference_view_name}
    )
    return {
        "expected_parent_view_name": expected_name,
        "generative_parent_view_name": parent_view_name,
        "reference_view_name": reference_view_name,
        "verified": verified,
        "read_errors": errors,
    }


def _section_snapshot(
    application: Any,
    section: Any,
    parent: Any,
    target_x: float,
    target_y: float,
    position_tolerance_mm: float = 0.05,
) -> dict[str, Any]:
    generative = _is_generative(section)
    geometric_elements_count = _safe_count(_safe_attr(section, "GeometricElements"))
    size: Optional[dict[str, Any]] = None
    size_error: Optional[str] = None
    try:
        size = _view_size(application, section)
    except Exception as exc:
        size_error = str(exc)

    # An empty section-view shell can still expose one GeometricElement while
    # DrawingView.Size remains (0, 0, 0, 0).  Collection count is therefore a
    # diagnostic signal only; a usable generated view must have a non-zero
    # two-dimensional bounding box.
    size_read_verified = bool(size is not None and size.get("read_verified"))
    bounding_box_nonempty = bool(
        size_read_verified
        and abs(size["width_mm"]) > 1e-6
        and abs(size["height_mm"]) > 1e-6
    )
    geometry_nonempty = bounding_box_nonempty
    try:
        actual_x, actual_y = _view_position(section)
        position_verified = bool(
            abs(actual_x - target_x) <= position_tolerance_mm
            and abs(actual_y - target_y) <= position_tolerance_mm
        )
        position_error = None
    except Exception as exc:
        actual_x = None
        actual_y = None
        position_verified = False
        position_error = str(exc)

    reference = _section_reference_summary(section, parent)
    return {
        "is_generative": generative,
        "geometric_elements_count": geometric_elements_count,
        "bounding_box": size,
        "bounding_box_error": size_error,
        "bounding_box_read_verified": size_read_verified,
        "bounding_box_nonempty": bounding_box_nonempty,
        "geometry_nonempty": geometry_nonempty,
        "geometric_elements_nonzero_signal": bool(
            geometric_elements_count is not None and geometric_elements_count > 0
        ),
        "requested_position_mm": [float(target_x), float(target_y)],
        "actual_position_mm": [actual_x, actual_y],
        "position_tolerance_mm": position_tolerance_mm,
        "position_verified": position_verified,
        "position_error": position_error,
        "parent_reference": reference,
        "generation_verified": bool(
            generative is True
            and geometry_nonempty
            and reference["verified"]
        ),
        "verified": bool(
            generative is True
            and geometry_nonempty
            and position_verified
            and reference["verified"]
        ),
    }


def _wait_for_section_generation(
    application: Any,
    sheet: Any,
    document: Any,
    section: Any,
    parent: Any,
    target_x: float,
    target_y: float,
    timeout_seconds: float = 20.0,
    poll_interval_seconds: float = 0.25,
) -> dict[str, Any]:
    deadline = time.monotonic() + timeout_seconds
    polls: list[dict[str, Any]] = []
    final_snapshot: Optional[dict[str, Any]] = None

    while True:
        refresh_attempts: list[dict[str, Any]] = []
        for name, action in (
            ("section.GenerativeBehavior.Update", lambda: section.GenerativeBehavior.Update()),
            ("section.GenerativeBehavior.ForceUpdate", lambda: section.GenerativeBehavior.ForceUpdate()),
            ("sheet.ForceUpdate", lambda: sheet.ForceUpdate()),
            ("sheet.Update", lambda: sheet.Update()),
            ("document.Update", lambda: document.Update()),
        ):
            try:
                action()
                refresh_attempts.append(
                    {"method": name, "succeeded": True, "error": None}
                )
            except Exception as exc:
                refresh_attempts.append(
                    {"method": name, "succeeded": False, "error": str(exc)}
                )

        time.sleep(poll_interval_seconds)
        final_snapshot = _section_snapshot(
            application,
            section,
            parent,
            target_x,
            target_y,
        )
        polls.append(
            {
                "elapsed_seconds": timeout_seconds
                - max(0.0, deadline - time.monotonic()),
                "refresh_attempts": refresh_attempts,
                "verified": final_snapshot["verified"],
                "generation_verified": final_snapshot["generation_verified"],
                "is_generative": final_snapshot["is_generative"],
                "geometry_nonempty": final_snapshot["geometry_nonempty"],
                "bounding_box_nonempty": final_snapshot["bounding_box_nonempty"],
                "position_verified": final_snapshot["position_verified"],
                "parent_reference_verified": final_snapshot[
                    "parent_reference"
                ]["verified"],
            }
        )
        if final_snapshot["generation_verified"] or time.monotonic() >= deadline:
            break

    return {
        "verified": bool(final_snapshot and final_snapshot["verified"]),
        "generation_verified": bool(
            final_snapshot and final_snapshot["generation_verified"]
        ),
        "timeout_seconds": timeout_seconds,
        "poll_interval_seconds": poll_interval_seconds,
        "poll_count": len(polls),
        "final_snapshot": final_snapshot,
        "polls": polls,
    }


def _normalise_choice(value: str, parameter_name: str, choices: dict[str, str]) -> str:
    key = _nonempty_text(value, parameter_name).lower().replace("-", "_").replace(" ", "_")
    if key not in choices:
        raise ValueError(
            f"{parameter_name} must be one of: {', '.join(sorted(choices))}."
        )
    return choices[key]


def _catia_constant(ctx: Any, names: Sequence[str], fallback: int) -> int:
    """Resolve a generated CATIA enum when the host connection exposes one."""
    constants = _safe_attr(ctx.conn, "constants", None)
    for name in names:
        try:
            return int(getattr(constants, name))
        except Exception:
            continue
    return fallback



def _reference_object_summary(value: Any) -> dict[str, Any]:
    """Return stable, serialization-safe evidence for a CATIA link/document."""
    if value is None:
        return {
            "available": False,
            "python_type": None,
            "name": None,
            "full_name": None,
            "parent_name": None,
            "has_oleobj": False,
        }

    name: Optional[str] = None
    full_name: Optional[str] = None
    parent_name: Optional[str] = None
    try:
        candidate = str(value.Name).strip()
        name = candidate or None
    except Exception:
        pass
    try:
        candidate = str(value.FullName).strip()
        full_name = candidate or None
    except Exception:
        pass
    try:
        candidate = str(value.Parent.Name).strip()
        parent_name = candidate or None
    except Exception:
        pass

    return {
        "available": True,
        "python_type": type(value).__name__,
        "python_module": type(value).__module__,
        "name": name,
        "full_name": full_name,
        "parent_name": parent_name,
        "has_oleobj": bool(hasattr(value, "_oleobj_")),
    }


def _call_noarg_member(value: Any, member_name: str) -> Any:
    member = getattr(value, member_name)
    return member() if callable(member) else member


def _generative_links_snapshot(
    links: Any,
    *,
    max_links: int = 16,
) -> dict[str, Any]:
    """Enumerate DrawingViewGenerativeLinks without relying on Count.

    CATIA exposes FirstLink/NextLink rather than a collection Count.  Some
    generated wrappers raise at the end of NextLink iteration; after at least
    one link has been read, that terminal error is retained as diagnostic
    evidence rather than treated as proof that the collection is empty.
    """
    result: dict[str, Any] = {
        "collection_available": links is not None,
        "enumeration_attempted": False,
        "enumeration_succeeded": False,
        "has_link": False,
        "link_count_observed": 0,
        "links": [],
        "first_link_error": None,
        "next_link_terminal_error": None,
        "truncated": False,
    }
    if links is None:
        return result

    result["enumeration_attempted"] = True
    try:
        current = _call_noarg_member(links, "FirstLink")
    except Exception as exc:
        result["first_link_error"] = str(exc)
        return result

    result["enumeration_succeeded"] = True
    seen_keys: set[tuple[Any, ...]] = set()
    while current is not None and len(result["links"]) < max_links:
        summary = _reference_object_summary(current)
        key = (
            summary.get("full_name"),
            summary.get("name"),
            summary.get("parent_name"),
            summary.get("python_type"),
        )
        if key in seen_keys and any(key):
            break
        seen_keys.add(key)
        result["links"].append(summary)
        try:
            current = _call_noarg_member(links, "NextLink")
        except Exception as exc:
            result["next_link_terminal_error"] = str(exc)
            break

    if current is not None and len(result["links"]) >= max_links:
        result["truncated"] = True
    result["link_count_observed"] = len(result["links"])
    result["has_link"] = bool(result["links"])
    return result


def _generative_document_snapshot(view: Any) -> dict[str, Any]:
    try:
        document = view.GenerativeBehavior.Document
    except Exception as exc:
        return {
            "available": False,
            "object": _reference_object_summary(None),
            "error": str(exc),
        }
    return {
        "available": document is not None,
        "object": _reference_object_summary(document),
        "error": None,
    }


def _reference_summaries_match(
    first: dict[str, Any],
    second: dict[str, Any],
) -> bool:
    if not first.get("available") or not second.get("available"):
        return False
    for key in ("full_name", "name"):
        first_value = str(first.get(key) or "").strip().casefold()
        second_value = str(second.get(key) or "").strip().casefold()
        if first_value and second_value and first_value == second_value:
            return True
    return False


def _generative_link_state(parent: Any, section: Any) -> dict[str, Any]:
    parent_links = None
    section_links = None
    parent_links_error: Optional[str] = None
    section_links_error: Optional[str] = None
    try:
        parent_links = parent.GenerativeLinks
    except Exception as exc:
        parent_links_error = str(exc)
    try:
        section_links = section.GenerativeLinks
    except Exception as exc:
        section_links_error = str(exc)

    parent_snapshot = _generative_links_snapshot(parent_links)
    section_snapshot = _generative_links_snapshot(section_links)
    parent_document = _generative_document_snapshot(parent)
    section_document = _generative_document_snapshot(section)
    document_match = _reference_summaries_match(
        parent_document["object"],
        section_document["object"],
    )
    verified = bool(section_snapshot["has_link"] or document_match)
    return {
        "parent_collection_error": parent_links_error,
        "section_collection_error": section_links_error,
        "parent_links": parent_snapshot,
        "section_links": section_snapshot,
        "parent_document": parent_document,
        "section_document": section_document,
        "document_match": document_match,
        "verified": verified,
    }


def _parent_generative_link_precheck(parent: Any) -> dict[str, Any]:
    """Verify that the parent view points at a usable 3D source."""
    try:
        parent_links = parent.GenerativeLinks
    except Exception as exc:
        raise AnnotationOperationError(
            "The parent drawing view does not expose GenerativeLinks.",
            data={
                "failure_stage": "B_parent_generative_link_precheck",
                "parent_view_name": str(_safe_attr(parent, "Name", "")),
                "collection_error": str(exc),
            },
        ) from exc

    links_snapshot = _generative_links_snapshot(parent_links)
    document_snapshot = _generative_document_snapshot(parent)
    verified = bool(
        links_snapshot["has_link"]
        or document_snapshot["available"]
    )
    evidence = {
        "parent_view_name": str(_safe_attr(parent, "Name", "")),
        "links": links_snapshot,
        "generative_document": document_snapshot,
        "verified": verified,
    }
    if not verified:
        raise AnnotationOperationError(
            "The parent view is generative but no usable 3D generative link or "
            "GenerativeBehavior.Document could be verified.",
            data={
                "failure_stage": "B_parent_generative_link_precheck",
                "parent_link_precheck": evidence,
            },
        )
    return evidence


def _copy_parent_generative_links(parent: Any, section: Any) -> dict[str, Any]:
    """Copy the parent's 3D generative links to the new section view.

    CATIA section-view macros generated by the application call
    DefineSectionView and then copy DrawingView.GenerativeLinks from the parent
    view to the child before updating the child GenerativeBehavior.
    """
    try:
        parent_links = parent.GenerativeLinks
    except Exception as exc:
        raise AnnotationOperationError(
            "Cannot access the parent DrawingView.GenerativeLinks collection.",
            data={
                "failure_stage": "E_GenerativeLinks_copy",
                "parent_collection_error": str(exc),
            },
        ) from exc
    try:
        section_links = section.GenerativeLinks
    except Exception as exc:
        raise AnnotationOperationError(
            "Cannot access the new section DrawingView.GenerativeLinks collection.",
            data={
                "failure_stage": "E_GenerativeLinks_copy",
                "section_collection_error": str(exc),
            },
        ) from exc

    before = _generative_link_state(parent, section)
    attempts: list[dict[str, Any]] = []
    selected_strategy: Optional[str] = None

    try:
        parent_links.CopyLinksTo(section_links)
        selected_strategy = "generated_or_native_proxy_CopyLinksTo"
        attempts.append(
            {
                "strategy": selected_strategy,
                "succeeded": True,
                "error": None,
            }
        )
    except Exception as exc:
        attempts.append(
            {
                "strategy": "generated_or_native_proxy_CopyLinksTo",
                "succeeded": False,
                "error": str(exc),
            }
        )

    if selected_strategy is None:
        try:
            from win32com.client.dynamic import DumbDispatch  # type: ignore

            ole_object = getattr(parent_links, "_oleobj_", None)
            if ole_object is None:
                raise RuntimeError("Parent GenerativeLinks has no _oleobj_.")
            dynamic_parent_links = DumbDispatch(
                ole_object,
                "DrawingViewGenerativeLinks",
            )
            dynamic_parent_links.CopyLinksTo(section_links)
            selected_strategy = "late_bound_DumbDispatch_CopyLinksTo"
            attempts.append(
                {
                    "strategy": selected_strategy,
                    "succeeded": True,
                    "error": None,
                }
            )
        except Exception as exc:
            attempts.append(
                {
                    "strategy": "late_bound_DumbDispatch_CopyLinksTo",
                    "succeeded": False,
                    "error": str(exc),
                }
            )

    if selected_strategy is None:
        raise AnnotationOperationError(
            "CATIA could not copy the parent view's generative links to the "
            "new section view.",
            data={
                "failure_stage": "E_GenerativeLinks_copy",
                "before": before,
                "attempts": attempts,
            },
        )

    after_immediate = _generative_link_state(parent, section)
    return {
        "copy_attempted": True,
        "copy_succeeded": True,
        "selected_strategy": selected_strategy,
        "attempts": attempts,
        "before": before,
        "after_immediate": after_immediate,
        "copy_call_order": "DefineSectionView -> CopyLinksTo -> Update",
    }


def _define_section_view_bstr(
    behavior: Any,
    profile: Sequence[float],
    section_type_bstr: str,
    profile_type_bstr: str,
    side_to_draw: int,
    parent_behavior: Any,
    profile_marshalling: str,
) -> dict[str, Any]:
    """Call DefineSectionView with CATBSTR values and explicit profile evidence.

    The CAA signature requires CATSafeArrayVariant for the profile.  v5 defaults
    to an explicit ``VT_ARRAY | VT_VARIANT`` payload; ``python_tuple`` remains
    available as a diagnostic compatibility strategy.  A generated pywin32
    wrapper is attempted first, followed by late-bound DumbDispatch only when
    the first invocation raises.
    """
    section_value = str(section_type_bstr)
    profile_value = str(profile_type_bstr)
    if section_value not in SECTION_TYPE_BSTRS:
        raise ValueError("section_type must resolve to SectionView or SectionCut.")
    if profile_value not in PROFILE_TYPE_BSTRS:
        raise ValueError("profile_type must resolve to Offset or Aligned.")

    profile_payload, marshalling = _marshal_section_profile(
        profile,
        profile_marshalling,
    )
    attempts: list[dict[str, Any]] = []

    try:
        behavior.DefineSectionView(
            profile_payload,
            section_value,
            profile_value,
            int(side_to_draw),
            parent_behavior,
        )
        attempts.append(
            {
                "strategy": "generated_or_native_proxy_CATBSTR",
                "succeeded": True,
                "error": None,
            }
        )
        return {
            "succeeded": True,
            "selected_strategy": attempts[-1]["strategy"],
            "profile_marshalling": marshalling,
            "section_type_argument": section_value,
            "section_type_argument_type": "CATBSTR",
            "profile_type_argument": profile_value,
            "profile_type_argument_type": "CATBSTR",
            "side_to_draw_argument": int(side_to_draw),
            "attempts": attempts,
        }
    except Exception as exc:
        attempts.append(
            {
                "strategy": "generated_or_native_proxy_CATBSTR",
                "succeeded": False,
                "error": str(exc),
            }
        )

    try:
        from win32com.client.dynamic import DumbDispatch  # type: ignore

        ole_object = getattr(behavior, "_oleobj_", None)
        if ole_object is None:
            raise RuntimeError("DrawingViewGenerativeBehavior has no _oleobj_.")
        dynamic_behavior = DumbDispatch(
            ole_object,
            "DrawingViewGenerativeBehavior",
        )
        dynamic_behavior.DefineSectionView(
            profile_payload,
            section_value,
            profile_value,
            int(side_to_draw),
            parent_behavior,
        )
        attempts.append(
            {
                "strategy": "late_bound_DumbDispatch_CATBSTR",
                "succeeded": True,
                "error": None,
            }
        )
        return {
            "succeeded": True,
            "selected_strategy": attempts[-1]["strategy"],
            "profile_marshalling": marshalling,
            "section_type_argument": section_value,
            "section_type_argument_type": "CATBSTR",
            "profile_type_argument": profile_value,
            "profile_type_argument_type": "CATBSTR",
            "side_to_draw_argument": int(side_to_draw),
            "attempts": attempts,
        }
    except Exception as exc:
        attempts.append(
            {
                "strategy": "late_bound_DumbDispatch_CATBSTR",
                "succeeded": False,
                "error": str(exc),
            }
        )

    raise AnnotationOperationError(
        "DefineSectionView failed with the requested CATBSTR and profile marshalling strategy.",
        data={
            "failure_stage": "E_DefineSectionView_call",
            "profile_marshalling": marshalling,
            "section_type_argument": section_value,
            "profile_type_argument": profile_value,
            "side_to_draw_argument": int(side_to_draw),
            "attempts": attempts,
        },
    )

def _gdt_symbol(symbol_type: str) -> tuple[str, int]:
    key = _nonempty_text(symbol_type, "symbol_type").lower()
    key = re.sub(r"[^a-z0-9]+", "_", key).strip("_")
    if key not in GDT_SYMBOL_CODES:
        canonical = sorted(
            name
            for name in GDT_SYMBOL_CODES
            if name
            not in {"roundness", "profile_of_line", "profile_of_surface", "true_position", "coaxiality"}
        )
        raise ValueError(
            f"Unsupported symbol_type '{symbol_type}'. Supported: {', '.join(canonical)}."
        )
    return key, GDT_SYMBOL_CODES[key]


def _datum_references(values: Optional[Sequence[Any]]) -> list[str]:
    if values is None:
        return []
    if isinstance(values, (str, bytes)):
        raise ValueError("datum_refs must be a list, for example ['A', 'B', 'C'].")
    result = [_nonempty_text(value, f"datum_refs[{index}]").upper() for index, value in enumerate(values)]
    if len(result) > 3:
        raise ValueError("datum_refs supports at most three reference compartments.")
    for index, value in enumerate(result):
        if any(character.isspace() or character == "|" for character in value):
            raise ValueError(
                f"datum_refs[{index}] cannot contain whitespace or '|'."
            )
    return result


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    """Register native classic-CATDrawing annotation tools."""
    names: list[str] = []

    @mcp.tool()
    def catia_add_linear_dimension(
        view_name: str,
        point1_coords: Tuple[float, float],
        point2_coords: Tuple[float, float],
    ) -> dict[str, Any]:
        """Create a native point-supported distance dimension.

        Coordinates are expressed in the target view coordinate system in
        millimetres. The dimension is associated with two hidden helper
        Point2D objects created by this tool; it is not associated with
        projected model edges or circles.
        """
        support_points: list[Any] = []
        dimension = None
        dimensions = None
        created_dimension_index: Optional[int] = None
        dimensions_before: Optional[int] = None
        geometry_before: Optional[int] = None
        warnings: list[str] = []

        try:
            first = _point2(point1_coords, "point1_coords")
            second = _point2(point2_coords, "point2_coords")
            if first == second:
                raise ValueError("point1_coords and point2_coords must be different.")

            document = _active_drawing_document(ctx)
            sheet = _active_sheet(document)
            view = _drawing_view(sheet, view_name)
            view.Activate()

            dimensions = view.Dimensions
            dimensions_before = int(dimensions.Count)
            geometric_elements = _safe_attr(view, "GeometricElements")
            geometry_before = _safe_count(geometric_elements)

            factory = view.Factory2D
            support_points = [
                factory.CreatePoint(first[0], first[1]),
                factory.CreatePoint(second[0], second[1]),
            ]
            geometry_after_supports = _safe_count(geometric_elements)
            if (
                geometry_before is not None
                and geometry_after_supports is not None
                and geometry_after_supports != geometry_before + 2
            ):
                raise AnnotationOperationError(
                    "Helper Point2D creation did not produce the expected geometry-count delta.",
                    data={
                        "geometry_count_before": geometry_before,
                        "geometry_count_after_supports": geometry_after_supports,
                        "expected_delta": 2,
                    },
                )

            dim_type_code = _catia_constant(
                ctx, ("catDimDistance",), CAT_DIM_DISTANCE
            )
            line_representation_code = _catia_constant(
                ctx, ("catDimAuto",), CAT_DIM_AUTO
            )
            dimension = dimensions.Add(
                dim_type_code,
                tuple(support_points),
                (0.0, 0.0, 0.0, 0.0),
                line_representation_code,
            )
            created_dimension_index = int(dimensions.Count)
            if created_dimension_index != dimensions_before + 1:
                raise AnnotationOperationError(
                    "DrawingDimensions.Add returned without increasing the collection count.",
                    data={
                        "dimensions_count_before": dimensions_before,
                        "dimensions_count_after": created_dimension_index,
                    },
                )

            hidden, hide_error = _hide_objects(document, support_points)
            if hide_error:
                warnings.append(
                    "Dimension support points were created but could not be hidden: "
                    f"{hide_error}"
                )

            _save_view_edition(view, warnings)
            _update_drawing(sheet, document, warnings)

            actual_dim_type: Optional[int]
            try:
                actual_dim_type = int(dimension.DimType)
                if actual_dim_type != dim_type_code:
                    raise AnnotationOperationError(
                        "Created DrawingDimension type does not match catDimDistance.",
                        data={
                            "requested_dim_type": dim_type_code,
                            "actual_dim_type": actual_dim_type,
                        },
                    )
            except AnnotationOperationError:
                raise
            except Exception as exc:
                actual_dim_type = None
                warnings.append(f"DrawingDimension.DimType could not be read back: {exc}")

            expected_value = math.hypot(
                second[0] - first[0],
                second[1] - first[1],
            )
            value_readback = _dimension_value_readback(dimension)
            value_tolerance = max(0.01, expected_value * 1e-6)
            value_verified: Optional[bool] = None
            if value_readback["available"]:
                value_verified = bool(
                    abs(value_readback["value"] - expected_value) <= value_tolerance
                )
                if not value_verified:
                    raise AnnotationOperationError(
                        "Created dimension value does not match the helper-point distance.",
                        data={
                            "expected_value_mm": expected_value,
                            "actual_value_mm": value_readback["value"],
                            "value_tolerance_mm": value_tolerance,
                            "value_readback": value_readback,
                        },
                    )
            else:
                warnings.append(
                    "DrawingDimension value could not be read back; collection and type "
                    "verification succeeded, but the numeric value remains NOT VERIFIED."
                )

            return _success(
                {
                    "dimension_name": str(_safe_attr(dimension, "Name", "")),
                    "dimension_index": created_dimension_index,
                    "dimension_type": "distance",
                    "requested_dim_type_code": dim_type_code,
                    "actual_dim_type_code": actual_dim_type,
                    "line_representation": "automatic",
                    "line_representation_code": line_representation_code,
                    "view_name": str(_safe_attr(view, "Name", view_name)),
                    "support_points_mm": [list(first), list(second)],
                    "support_geometry_created": True,
                    "support_geometry_hidden": hidden,
                    "support_associativity": (
                        "native_DrawingDimension_associated_to_hidden_helper_Point2D; "
                        "not_associated_to_projected_model_geometry"
                    ),
                    "dimensions_count_before": dimensions_before,
                    "dimensions_count_after": created_dimension_index,
                    "geometry_count_before": geometry_before,
                    "geometry_count_after_supports": geometry_after_supports,
                    "expected_value_mm": expected_value,
                    "value_tolerance_mm": value_tolerance,
                    "value_readback": value_readback,
                    "value_verified": value_verified,
                    "native_drawing_dimension_created": True,
                    "projected_geometry_associativity": False,
                    "model_modified": True,
                    "document_save_required": True,
                },
                warnings,
            )
        except Exception as exc:
            dimension_cleanup = {
                "attempted": False,
                "succeeded": True,
                "verified": True,
            }
            if dimensions is not None:
                current_dimension_count = _safe_count(dimensions)
                if (
                    created_dimension_index is None
                    and dimensions_before is not None
                    and current_dimension_count is not None
                    and current_dimension_count > dimensions_before
                ):
                    created_dimension_index = current_dimension_count
                if created_dimension_index is not None:
                    dimension_cleanup = _remove_collection_item(
                        dimensions,
                        index=created_dimension_index,
                        name=str(_safe_attr(dimension, "Name", "")),
                    )

            support_cleanup = {
                "attempted": False,
                "requested_count": 0,
                "succeeded": True,
                "error": None,
            }
            if support_points:
                try:
                    support_cleanup = _delete_objects(document, support_points)
                except UnboundLocalError:
                    pass

            dimensions_after_rollback = _safe_count(dimensions) if dimensions is not None else None
            geometry_after_rollback = None
            try:
                geometry_after_rollback = _safe_count(view.GeometricElements)
            except Exception:
                pass
            rollback_verified = bool(
                dimension_cleanup.get("verified", True)
                and support_cleanup.get("succeeded", True)
                and (
                    dimensions_before is None
                    or dimensions_after_rollback is None
                    or dimensions_after_rollback == dimensions_before
                )
                and (
                    geometry_before is None
                    or geometry_after_rollback is None
                    or geometry_after_rollback == geometry_before
                )
            )
            return _error(
                "catia_add_linear_dimension",
                exc,
                data={
                    "dimension_cleanup": dimension_cleanup,
                    "support_cleanup": support_cleanup,
                    "dimensions_count_before": dimensions_before,
                    "dimensions_count_after_rollback": dimensions_after_rollback,
                    "geometry_count_before": geometry_before,
                    "geometry_count_after_rollback": geometry_after_rollback,
                    "rollback_verified": rollback_verified,
                    "model_modified": not rollback_verified,
                    "document_save_required": not rollback_verified,
                },
                warnings=warnings,
                status="error" if rollback_verified else "partial_success",
            )

    names.append("catia_add_linear_dimension")

    @mcp.tool()
    def catia_create_section_view(
        parent_view_name: str,
        cut_line_points: list[Tuple[float, float]],
        offset_direction: Tuple[float, float] = (100.0, 0.0),
        section_type: str = "section_view",
        profile_type: str = "offset",
        side_to_draw: int = 0,
        view_name: str = "",
        coordinate_mode: str = "parent_local",
        profile_marshalling: str = "explicit_variant",
        require_profile_intersection: bool = True,
    ) -> dict[str, Any]:
        """Create and verify a generative section view.

        ``cut_line_points`` are interpreted according to ``coordinate_mode``:

        - ``parent_local`` (default): coordinates are already expressed in the
          parent drawing-view axis system required by DefineSectionView.
        - ``sheet``: coordinates are expressed in the drawing-sheet axis system
          and are converted using parent.xAxisData/yAxisData and parent.Angle.

        ``offset_direction`` remains a sheet-coordinate placement delta for the
        new child view. ``profile_marshalling`` defaults to an explicit
        CATSafeArrayVariant (VT_ARRAY | VT_VARIANT); ``python_tuple`` is retained
        only as a diagnostic compatibility strategy.  v7 resolves the root
        Application through conn.connect(visible=True), validates parent/child
        extents through CATIA-side SystemService.Evaluate, and makes a disabled
        profile-intersection check a true bypass.
        """
        section = None
        views = None
        created_index: Optional[int] = None
        actual_name = ""
        before_count: Optional[int] = None
        warnings: list[str] = []
        profile_precheck: Optional[dict[str, Any]] = None
        application: Any = None
        application_resolution: Optional[dict[str, Any]] = None

        try:
            _, input_profile_points = _normalise_profile(cut_line_points)
            coordinate_mode_resolved = _normalise_coordinate_mode(coordinate_mode)
            marshalling_resolved = _normalise_profile_marshalling(profile_marshalling)
            offset = _point2(offset_direction, "offset_direction")
            if offset == (0.0, 0.0):
                raise ValueError("offset_direction cannot be [0, 0].")
            if isinstance(side_to_draw, bool) or side_to_draw not in (0, 1):
                raise ValueError("side_to_draw must be 0 (left) or 1 (right).")

            actual_section_type = _normalise_choice(
                section_type,
                "section_type",
                {
                    "section_view": "SectionView",
                    "view": "SectionView",
                    "section_cut": "SectionCut",
                    "cut": "SectionCut",
                },
            )
            actual_profile_type = _normalise_choice(
                profile_type,
                "profile_type",
                {"offset": "Offset", "aligned": "Aligned"},
            )

            document = _active_drawing_document(ctx)
            sheet = _active_sheet(document)
            parent = _drawing_view(sheet, parent_view_name)

            try:
                application, application_resolution = _catia_application(
                    ctx,
                    document,
                )
            except AnnotationOperationError as application_error:
                inherited = (
                    application_error.data
                    if isinstance(application_error.data, dict)
                    else {}
                )
                application_resolution = inherited.get(
                    "application_resolution",
                    {
                        "resolved": False,
                        "method": None,
                        "system_service_available": False,
                        "attempts": [],
                    },
                )
                if require_profile_intersection:
                    raise AnnotationOperationError(
                        "Cannot perform the required parent-view BoundingBox precheck "
                        "because the live CATIA Application could not be resolved.",
                        data={
                            "failure_stage": "C_parent_view_bbox_read",
                            "application_resolution": application_resolution,
                        },
                    ) from application_error
                warnings.append(
                    "Live CATIA Application resolution failed, but "
                    "require_profile_intersection=false so profile BoundingBox "
                    "precheck was bypassed. Final DrawingView.Size verification "
                    "will use any available direct-COM fallback and may remain "
                    "unverified."
                )

            parent_generative = _is_generative(parent)
            if parent_generative is not True:
                raise ValueError(
                    f"Parent view '{parent_view_name}' must be a verified generative view."
                )
            parent_behavior = parent.GenerativeBehavior
            parent_link_precheck = _parent_generative_link_precheck(parent)

            profile_precheck = _prepare_section_profile(
                application,
                parent,
                input_profile_points,
                coordinate_mode_resolved,
                bool(require_profile_intersection),
            )
            profile = tuple(
                float(value)
                for value in profile_precheck["flattened_parent_local_profile"]
            )
            profile_points = profile_precheck["parent_local_points_mm"]

            views = sheet.Views
            before_count = int(views.Count)
            requested_name = (
                _nonempty_text(view_name, "view_name")
                if str(view_name).strip()
                else f"Section_{_safe_attr(parent, 'Name', parent_view_name)}"
            )
            requested_name = _unique_view_name(views, requested_name)
            try:
                section = views.Add(requested_name)
            except Exception as exc:
                raise AnnotationOperationError(
                    "CATIA could not add the new DrawingView to the active sheet.",
                    data={
                        "failure_stage": "D_DrawingView_add",
                        "views_count_before": before_count,
                        "requested_view_name": requested_name,
                        "profile_precheck": profile_precheck,
                        "application_resolution": application_resolution,
                        "error": str(exc),
                    },
                ) from exc
            created_index = int(views.Count)
            actual_name = str(_safe_attr(section, "Name", requested_name))

            try:
                parent.Activate()
            except Exception as exc:
                warnings.append(
                    f"Parent view could not be activated before section definition: {exc}"
                )

            behavior = section.GenerativeBehavior
            definition_call = _define_section_view_bstr(
                behavior,
                profile,
                actual_section_type,
                actual_profile_type,
                side_to_draw,
                parent_behavior,
                marshalling_resolved,
            )

            generative_links = _copy_parent_generative_links(parent, section)
            definition_reference = _section_reference_summary(section, parent)

            parent_x, parent_y = _view_position(parent)
            target_x = parent_x + offset[0]
            target_y = parent_y + offset[1]
            position_result = _set_view_position(section, target_x, target_y)

            generation_wait = _wait_for_section_generation(
                application,
                sheet,
                document,
                section,
                parent,
                target_x,
                target_y,
            )
            generative_links["after_generation"] = _generative_link_state(
                parent, section
            )
            generative_links["verified_after_generation"] = bool(
                generative_links["after_generation"]["verified"]
                or generation_wait["generation_verified"]
            )
            after_count = int(views.Count)
            if after_count != before_count + 1:
                raise AnnotationOperationError(
                    "DrawingViews.Add returned without increasing the collection count.",
                    data={
                        "failure_stage": "D_DrawingView_add",
                        "views_count_before": before_count,
                        "views_count_after": after_count,
                        "profile_precheck": profile_precheck,
                        "application_resolution": application_resolution,
                        "parent_link_precheck": parent_link_precheck,
                        "definition_call": definition_call,
                        "generative_links": generative_links,
                        "definition_reference": definition_reference,
                        "generation_wait": generation_wait,
                    },
                )
            if not generation_wait["generation_verified"]:
                links_verified = bool(generative_links.get("verified_after_generation"))
                raise AnnotationOperationError(
                    "Section definition returned, but CATIA did not create a non-empty "
                    "generative child view within the timeout.",
                    data={
                        "failure_stage": (
                            "G_GenerativeBehavior_generation_or_bbox_verification"
                            if links_verified
                            else "F_GenerativeLinks_copy_verification"
                        ),
                        "views_count_before": before_count,
                        "views_count_after": after_count,
                        "profile_precheck": profile_precheck,
                        "application_resolution": application_resolution,
                        "parent_link_precheck": parent_link_precheck,
                        "definition_call": definition_call,
                        "generative_links": generative_links,
                        "definition_reference": definition_reference,
                        "generation_wait": generation_wait,
                    },
                )

            final_position_result = _set_view_position(section, target_x, target_y)
            try:
                behavior.Update()
            except Exception as exc:
                warnings.append(
                    f"GenerativeBehavior.Update after final positioning failed: {exc}"
                )
            final_snapshot = _section_snapshot(
                application,
                section,
                parent,
                target_x,
                target_y,
            )
            generation_wait["final_position_result"] = final_position_result
            generation_wait["final_snapshot_after_reposition"] = final_snapshot
            generation_wait["verified"] = final_snapshot["verified"]
            if not final_snapshot["position_verified"]:
                raise AnnotationOperationError(
                    "Section geometry was generated, but the final view position could "
                    "not be verified after CATIA automatic placement.",
                    data={
                        "failure_stage": "H_position_verification",
                        "views_count_before": before_count,
                        "views_count_after": after_count,
                        "profile_precheck": profile_precheck,
                        "application_resolution": application_resolution,
                        "parent_link_precheck": parent_link_precheck,
                        "definition_call": definition_call,
                        "generative_links": generative_links,
                        "definition_reference": definition_reference,
                        "generation_wait": generation_wait,
                    },
                )
            return _success(
                {
                    "view_name": actual_name,
                    "view_index": after_count,
                    "parent_view_name": str(_safe_attr(parent, "Name", parent_view_name)),
                    "section_type": actual_section_type,
                    "section_type_argument": actual_section_type,
                    "section_type_argument_type": "CATBSTR",
                    "profile_type": actual_profile_type,
                    "profile_type_argument": actual_profile_type,
                    "profile_type_argument_type": "CATBSTR",
                    "side_to_draw": side_to_draw,
                    "coordinate_mode": coordinate_mode_resolved,
                    "profile_marshalling_requested": marshalling_resolved,
                    "profile_precheck": profile_precheck,
                    "application_resolution": application_resolution,
                    "parent_link_precheck": parent_link_precheck,
                    "definition_call": definition_call,
                    "generative_links": generative_links,
                    "definition_reference_immediately_after_call": definition_reference,
                    "cut_line_points_input_mm": input_profile_points,
                    "cut_line_points_parent_local_mm": profile_points,
                    "placement_offset_sheet_mm": list(offset),
                    "position_mm": [target_x, target_y],
                    "position_result": position_result,
                    "is_generative": final_snapshot["is_generative"],
                    "geometry_nonempty": final_snapshot["geometry_nonempty"],
                    "bounding_box": final_snapshot["bounding_box"],
                    "geometric_elements_count": final_snapshot["geometric_elements_count"],
                    "parent_reference": final_snapshot["parent_reference"],
                    "generation_wait": generation_wait,
                    "creation_verified": True,
                    "views_count_before": before_count,
                    "views_count_after": after_count,
                    "model_modified": True,
                    "document_save_required": True,
                },
                warnings,
            )
        except Exception as exc:
            cleanup = {
                "attempted": False,
                "succeeded": True,
                "verified": True,
            }
            if section is not None and views is not None:
                cleanup = _remove_view(views, actual_name, created_index)
            after_cleanup = _safe_count(views) if views is not None else None
            cleanup_verified = bool(
                cleanup.get("verified", True)
                and (
                    before_count is None
                    or after_cleanup is None
                    or after_cleanup == before_count
                )
            )
            inherited_data = exc.data if isinstance(exc, AnnotationOperationError) else None
            return _error(
                "catia_create_section_view",
                exc,
                data={
                    "failure_stage": (
                        inherited_data.get("failure_stage")
                        if isinstance(inherited_data, dict)
                        else (
                            "E_DefineSectionView_call"
                            if section is not None
                            else "A_input_or_parent_precheck"
                        )
                    ),
                    "profile_precheck": profile_precheck,
                    "application_resolution": application_resolution,
                    "failure_evidence": inherited_data,
                    "view_cleanup": cleanup,
                    "views_count_before": before_count,
                    "views_count_after_cleanup": after_cleanup,
                    "rollback_verified": cleanup_verified,
                    "model_modified": not cleanup_verified,
                    "document_save_required": not cleanup_verified,
                },
                warnings=warnings,
                status="error" if cleanup_verified else "partial_success",
            )

    names.append("catia_create_section_view")

    @mcp.tool()
    def catia_create_gdt_frame(
        view_name: str,
        symbol_type: str,
        tolerance_value: str,
        datum_refs: Optional[list[str]] = None,
        position_xy: Tuple[float, float] = (50.0, 50.0),
        leader_xy: Optional[Tuple[float, float]] = None,
    ) -> dict[str, Any]:
        """Create and verify a native two-dimensional DrawingGDT frame.

        This tool creates a classic drafting annotation. It does not create a
        semantic three-dimensional FT&A tolerance.
        """
        gdt = None
        gdts = None
        created_index: Optional[int] = None
        before_count: Optional[int] = None
        warnings: list[str] = []

        try:
            canonical_symbol, symbol_code = _gdt_symbol(symbol_type)
            tolerance = _nonempty_text(tolerance_value, "tolerance_value")
            if "|" in tolerance:
                raise ValueError("tolerance_value cannot contain '|'.")
            datums = _datum_references(datum_refs)
            frame_position = _point2(position_xy, "position_xy")
            leader_position = (
                _point2(leader_xy, "leader_xy")
                if leader_xy is not None
                else frame_position
            )

            document = _active_drawing_document(ctx)
            sheet = _active_sheet(document)
            view = _drawing_view(sheet, view_name)
            try:
                gdts = view.GDTs
                before_count = int(gdts.Count)
            except Exception as exc:
                raise CapabilityUnavailableError(
                    "This CATIA installation/view does not expose DrawingView.GDTs. "
                    "The tool will not silently substitute a plain DrawingText."
                ) from exc

            encoded_text = "|".join([tolerance, *datums])
            try:
                gdt = gdts.Add(
                    leader_position[0],
                    leader_position[1],
                    frame_position[0],
                    frame_position[1],
                    symbol_code,
                    encoded_text,
                )
            except Exception as exc:
                raise CapabilityUnavailableError(
                    "DrawingGDTs.Add failed. Confirm a classic CATDrawing view and "
                    "the required Drafting license are active."
                ) from exc

            created_index = int(gdts.Count)
            if created_index != before_count + 1:
                raise AnnotationOperationError(
                    "DrawingGDTs.Add returned without increasing the collection count.",
                    data={
                        "gdts_count_before": before_count,
                        "gdts_count_after": created_index,
                    },
                )

            try:
                gdt.SetToleranceType(1, symbol_code)
            except Exception as exc:
                warnings.append(f"DrawingGDT.SetToleranceType verification failed: {exc}")

            actual_symbol_code: Optional[int]
            try:
                actual_symbol_code = int(gdt.GetToleranceType(1))
                if actual_symbol_code != symbol_code:
                    raise AnnotationOperationError(
                        "DrawingGDT symbol readback mismatch.",
                        data={
                            "requested_symbol_code": symbol_code,
                            "actual_symbol_code": actual_symbol_code,
                        },
                    )
            except AnnotationOperationError:
                raise
            except Exception as exc:
                actual_symbol_code = None
                warnings.append(f"DrawingGDT symbol could not be read back: {exc}")

            leaders = _safe_attr(gdt, "Leaders")
            leader_count_before_cleanup = _safe_count(leaders)
            leader_cleanup_attempts: list[dict[str, Any]] = []
            if leader_xy is None and leaders is not None:
                count = _safe_count(leaders) or 0
                for index in range(count, 0, -1):
                    try:
                        leaders.Remove(index)
                        leader_cleanup_attempts.append(
                            {"index": index, "succeeded": True, "error": None}
                        )
                    except Exception as exc:
                        leader_cleanup_attempts.append(
                            {"index": index, "succeeded": False, "error": str(exc)}
                        )
                if (_safe_count(leaders) or 0) > 0:
                    warnings.append(
                        "The automatically created GDT leader could not be fully removed."
                    )

            _save_view_edition(view, warnings)
            _update_drawing(sheet, document, warnings)

            leader_count_after_cleanup = _safe_count(leaders)
            leader_created = bool(
                leader_xy is not None
                or (leader_count_after_cleanup is not None and leader_count_after_cleanup > 0)
            )
            return _success(
                {
                    "annotation_type": "DrawingGDT",
                    "capability": "native_2d_drafting_gdt",
                    "semantic_3d_fta": False,
                    "gdt_index": created_index,
                    "gdts_count_before": before_count,
                    "gdts_count_after": created_index,
                    "view_name": str(_safe_attr(view, "Name", view_name)),
                    "symbol_type": canonical_symbol,
                    "symbol_code": symbol_code,
                    "actual_symbol_code": actual_symbol_code,
                    "symbol_readback_verified": (
                        actual_symbol_code == symbol_code
                        if actual_symbol_code is not None
                        else None
                    ),
                    "tolerance_value": tolerance,
                    "datum_references": datums,
                    "position_mm": list(frame_position),
                    "leader_requested": leader_xy is not None,
                    "leader_created": leader_created,
                    "leader_position_mm": (
                        list(leader_position) if leader_created else None
                    ),
                    "leader_count_before_cleanup": leader_count_before_cleanup,
                    "leader_count_after_cleanup": leader_count_after_cleanup,
                    "leader_cleanup_attempts": leader_cleanup_attempts,
                    "encoded_text": encoded_text,
                    "native_drawing_gdt_created": True,
                    "model_modified": True,
                    "document_save_required": True,
                },
                warnings,
            )
        except Exception as exc:
            cleanup = {
                "attempted": False,
                "succeeded": True,
                "verified": True,
            }
            if gdts is not None:
                current_count = _safe_count(gdts)
                if created_index is None and before_count is not None and current_count is not None:
                    if current_count > before_count:
                        created_index = current_count
                if created_index is not None:
                    cleanup = _remove_collection_item(
                        gdts,
                        index=created_index,
                        name=str(_safe_attr(gdt, "Name", "")),
                    )
            count_after_cleanup = _safe_count(gdts) if gdts is not None else None
            cleanup_verified = bool(
                cleanup.get("verified", True)
                and (
                    before_count is None
                    or count_after_cleanup is None
                    or count_after_cleanup == before_count
                )
            )
            inherited_data = exc.data if isinstance(exc, AnnotationOperationError) else None
            return _error(
                "catia_create_gdt_frame",
                exc,
                data={
                    "failure_evidence": inherited_data,
                    "gdt_cleanup": cleanup,
                    "gdts_count_before": before_count,
                    "gdts_count_after_cleanup": count_after_cleanup,
                    "rollback_verified": cleanup_verified,
                    "model_modified": not cleanup_verified,
                    "document_save_required": not cleanup_verified,
                },
                warnings=warnings,
                status=(
                    "capability_unavailable"
                    if isinstance(exc, CapabilityUnavailableError)
                    and cleanup_verified
                    else ("error" if cleanup_verified else "partial_success")
                ),
            )

    names.append("catia_create_gdt_frame")

    return names
