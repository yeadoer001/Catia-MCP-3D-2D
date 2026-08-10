"""Safe CATIA V5 drafting annotation layout tools.

The public CATIA V5 Automation API does *not* expose
``DrawingDimension.GetTextPosition/SetTextPosition`` or
``DrawingText.GetAnchorPosition/SetAnchorPosition``.  Dimension value bounds
and movement are handled with ``GetBoundaryBox`` and ``MoveValue``.  Drawing
text positions are handled with the ``x``/``y`` properties, whose Automation
unit is metres; all MCP inputs and outputs from this module use millimetres.

Text extents are not exposed by the public Drafting Automation API.  Their
boxes are therefore conservative estimates based on text, font size, wrapping,
anchor and rotation.  Every response identifies exact versus estimated boxes.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import logging
import math
import unicodedata
from statistics import median
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

logger = logging.getLogger(__name__)

IMPLEMENTATION_VERSION = "smart-annotation-layout-fixed-2026-08-05-v3"
_M_TO_MM = 1000.0
_MM_TO_M = 0.001
_EPSILON_MM = 1.0e-6
_CATVB_SCRIPT_LANGUAGE = 0


def _success(data: Any, warnings: Optional[List[str]] = None) -> Dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": True,
        "status": "success_with_warnings" if warning_list else "success",
        "data": data,
        "warnings": warning_list,
    }


def _error(
    message: str,
    *,
    data: Any = None,
    warnings: Optional[List[str]] = None,
    status: str = "error",
) -> Dict[str, Any]:
    result: Dict[str, Any] = {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": False,
        "status": status,
        "error": str(message),
        "warnings": list(warnings or []),
    }
    if data is not None:
        result["data"] = data
    return result


class ToolOperationError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        data: Any = None,
        warnings: Optional[List[str]] = None,
        status: str = "error",
    ) -> None:
        super().__init__(message)
        self.data = data
        self.warnings = list(warnings or [])
        self.status = status


def _format_com_error(exc: BaseException) -> str:
    details = getattr(exc, "excepinfo", None)
    if details and len(details) >= 3 and details[2]:
        return str(details[2])
    hresult = getattr(exc, "hresult", None)
    if hresult is not None:
        return f"{exc} (HRESULT 0x{int(hresult) & 0xFFFFFFFF:08X})"
    return str(exc)


def _finite_float(value: Any, name: str) -> float:
    if isinstance(value, bool):
        raise ValueError(f"{name} must be a finite number.")
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{name} must be a finite number.") from exc
    if not math.isfinite(result):
        raise ValueError(f"{name} must be a finite number.")
    return result


def _positive_int(value: Any, name: str, maximum: int = 1000) -> int:
    if isinstance(value, bool):
        raise ValueError(f"{name} must be a positive integer.")
    try:
        result = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{name} must be a positive integer.") from exc
    if result < 1 or result > maximum or result != value:
        raise ValueError(f"{name} must be an integer in 1..{maximum}.")
    return result


def _safe_attr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _document_saved(document: Any) -> Optional[bool]:
    try:
        return bool(document.Saved)
    except Exception:
        return None


@dataclass
class AnnotationBoundingBox:
    """Axis-aligned bounds in the target DrawingView coordinate system."""

    x_min: float
    y_min: float
    x_max: float
    y_max: float
    item_id: str = ""
    index: int = 0
    kind: str = "text"
    name: str = ""
    text_content: str = ""
    boundary_method: str = "estimated"
    boundary_verified: bool = False
    boundary_attempts: List[Dict[str, Any]] = field(default_factory=list)

    @property
    def width(self) -> float:
        return self.x_max - self.x_min

    @property
    def height(self) -> float:
        return self.y_max - self.y_min

    @property
    def center_x(self) -> float:
        return (self.x_min + self.x_max) / 2.0

    @property
    def center_y(self) -> float:
        return (self.y_min + self.y_max) / 2.0

    def translate(self, dx_mm: float, dy_mm: float) -> None:
        self.x_min += dx_mm
        self.x_max += dx_mm
        self.y_min += dy_mm
        self.y_max += dy_mm

    def conflict(
        self,
        other: "AnnotationBoundingBox",
        minimum_gap_mm: float = 0.0,
    ) -> Tuple[bool, float, str]:
        """Return conflict, required clearance movement and best move axis."""
        gap = max(0.0, float(minimum_gap_mm))
        penetration_x = (
            min(self.x_max, other.x_max)
            - max(self.x_min, other.x_min)
            + gap
        )
        penetration_y = (
            min(self.y_max, other.y_max)
            - max(self.y_min, other.y_min)
            + gap
        )
        if penetration_x <= _EPSILON_MM or penetration_y <= _EPSILON_MM:
            return False, 0.0, "none"
        if penetration_x <= penetration_y:
            return True, penetration_x, "x"
        return True, penetration_y, "y"

    def overlaps(
        self,
        other: "AnnotationBoundingBox",
        tolerance: float = 0.0,
    ) -> Tuple[bool, float]:
        conflicting, depth, _ = self.conflict(other, tolerance)
        return conflicting, depth if conflicting else -_box_gap(self, other)

    def as_dict(self) -> Dict[str, Any]:
        return {
            "id": self.item_id,
            "index": self.index,
            "kind": self.kind,
            "name": self.name,
            "text": self.text_content,
            "x_min_mm": round(self.x_min, 6),
            "y_min_mm": round(self.y_min, 6),
            "x_max_mm": round(self.x_max, 6),
            "y_max_mm": round(self.y_max, 6),
            "width_mm": round(self.width, 6),
            "height_mm": round(self.height, 6),
            "center_mm": [round(self.center_x, 6), round(self.center_y, 6)],
            "boundary_method": self.boundary_method,
            "boundary_verified": self.boundary_verified,
            "boundary_attempts": self.boundary_attempts,
        }

@dataclass
class LayoutResult:
    total_annotations: int = 0
    moved_count: int = 0
    overlaps_detected: int = 0
    overlaps_fixed: int = 0
    overlaps_remaining: int = 0
    details: List[Dict[str, Any]] = field(default_factory=list)
    document_unit: str = "mm"


@dataclass
class _TextRecord:
    index: int
    obj: Any
    anchor_x_mm: float
    anchor_y_mm: float
    box: AnnotationBoundingBox


@dataclass
class _DimensionRecord:
    index: int
    obj: Any
    box: AnnotationBoundingBox
    dim_type: Optional[int]


def _box_gap(a: AnnotationBoundingBox, b: AnnotationBoundingBox) -> float:
    dx = max(a.x_min - b.x_max, b.x_min - a.x_max, 0.0)
    dy = max(a.y_min - b.y_max, b.y_min - a.y_max, 0.0)
    return math.hypot(dx, dy)


def _get_drawing_view(
    catia_app: Any,
    view_name: Optional[str] = None,
    sheet_index: Optional[int] = None,
) -> Tuple[Any, Any, Any, Dict[str, Any]]:
    try:
        document = catia_app.ActiveDocument
    except Exception as exc:
        raise RuntimeError("CATIA has no accessible active document.") from exc
    if document is None:
        raise RuntimeError("No active document. Open or create a CATDrawing first.")
    try:
        sheets = document.Sheets
        sheet_count = int(sheets.Count)
    except Exception as exc:
        raise RuntimeError("The active document is not a CATDrawing.") from exc

    if sheet_index is None:
        sheet = sheets.ActiveSheet
        resolved_sheet_index: Optional[int] = None
        for candidate in range(1, sheet_count + 1):
            try:
                if sheets.Item(candidate).Name == sheet.Name:
                    resolved_sheet_index = candidate
                    break
            except Exception:
                continue
    else:
        resolved_sheet_index = _positive_int(sheet_index, "sheet_index", sheet_count)
        sheet = sheets.Item(resolved_sheet_index)

    views = sheet.Views
    if view_name is None:
        view = views.ActiveView
        selection_method = "DrawingViews.ActiveView"
    else:
        clean_name = str(view_name).strip()
        if not clean_name:
            raise ValueError("view_name cannot be empty.")
        try:
            view = views.Item(clean_name)
        except Exception as exc:
            raise RuntimeError(f"DrawingView '{clean_name}' was not found.") from exc
        selection_method = f"DrawingViews.Item({clean_name!r})"

    return view, sheet, document, {
        "document_name": str(_safe_attr(document, "Name", "")),
        "sheet_name": str(_safe_attr(sheet, "Name", "")),
        "sheet_index": resolved_sheet_index,
        "sheet_count": sheet_count,
        "view_name": str(_safe_attr(view, "Name", "")),
        "view_type_code": _safe_int(_safe_attr(view, "ViewType", None)),
        "selection_method": selection_method,
    }


def _safe_int(value: Any) -> Optional[int]:
    try:
        return int(value)
    except (TypeError, ValueError, OverflowError):
        return None


def _flatten_numbers(value: Any) -> List[float]:
    result: List[float] = []
    if isinstance(value, (list, tuple)):
        for child in value:
            result.extend(_flatten_numbers(child))
        return result
    try:
        number = float(value)
    except (TypeError, ValueError):
        return result
    if math.isfinite(number):
        result.append(number)
    return result


def _evaluate(
    application: Any,
    script: str,
    function_name: str,
    parameters: List[Any],
) -> Any:
    try:
        service = application.SystemService
    except Exception as exc:
        raise RuntimeError(f"Cannot access CATIA SystemService: {_format_com_error(exc)}") from exc
    try:
        return service.Evaluate(
            script,
            _CATVB_SCRIPT_LANGUAGE,
            function_name,
            parameters,
        )
    except Exception as exc:
        raise RuntimeError(
            f"SystemService.Evaluate failed for {function_name}: {_format_com_error(exc)}"
        ) from exc


def _numeric_sequence(
    value: Any,
    expected_length: int,
    label: str,
) -> List[float]:
    try:
        sequence = list(value)
    except Exception as exc:
        raise RuntimeError(f"{label} did not return an array.") from exc
    if len(sequence) != expected_length:
        raise RuntimeError(
            f"{label} returned {len(sequence)} values; {expected_length} are required."
        )
    result: List[float] = []
    for index, item in enumerate(sequence):
        number = _finite_float(item, f"{label}[{index}]")
        result.append(number)
    return result


def _build_dimension_box(
    values: Sequence[float],
    index: int,
    dim: Any,
    method: str,
    attempts: List[Dict[str, Any]],
) -> AnnotationBoundingBox:
    numbers = _numeric_sequence(values, 8, method)
    xs = numbers[0::2]
    ys = numbers[1::2]
    width = max(xs) - min(xs)
    height = max(ys) - min(ys)
    if width <= _EPSILON_MM or height <= _EPSILON_MM:
        raise RuntimeError(
            f"{method} returned a zero-area dimension value box."
        )
    name = str(_safe_attr(dim, "Name", f"Dimension.{index}"))
    return AnnotationBoundingBox(
        min(xs),
        min(ys),
        max(xs),
        max(ys),
        item_id=f"dimension:{index}",
        index=index,
        kind="dimension",
        name=name,
        boundary_method=method,
        boundary_verified=True,
        boundary_attempts=list(attempts),
    )


def _dimension_boundary(
    application: Any,
    dim: Any,
    index: int,
) -> AnnotationBoundingBox:
    """Read DrawingDimension.GetBoundaryBox with verified output-array handling."""
    attempts: List[Dict[str, Any]] = []
    script = (
        "Public Function MCP_GetDrawingDimensionBoundary(dimObject)\n"
        "    Dim values(7)\n"
        "    dimObject.GetBoundaryBox values\n"
        "    MCP_GetDrawingDimensionBoundary = Array("
        "CDbl(values(0)), CDbl(values(1)), CDbl(values(2)), CDbl(values(3)), "
        "CDbl(values(4)), CDbl(values(5)), CDbl(values(6)), CDbl(values(7)))\n"
        "End Function"
    )
    try:
        values = _numeric_sequence(
            _evaluate(
                application,
                script,
                "MCP_GetDrawingDimensionBoundary",
                [dim],
            ),
            8,
            "SystemService.Evaluate.DrawingDimension.GetBoundaryBox",
        )
        attempts.append({
            "method": "SystemService.Evaluate.DrawingDimension.GetBoundaryBox",
            "succeeded": True,
            "values": values,
            "error": None,
        })
        box = _build_dimension_box(
            values,
            index,
            dim,
            "SystemService.Evaluate.DrawingDimension.GetBoundaryBox",
            attempts,
        )
        box.boundary_attempts = list(attempts)
        return box
    except Exception as exc:
        attempts.append({
            "method": "SystemService.Evaluate.DrawingDimension.GetBoundaryBox",
            "succeeded": False,
            "values": None,
            "error": _format_com_error(exc),
        })

    try:
        raw = dim.GetBoundaryBox()
        values = _numeric_sequence(
            raw, 8, "DrawingDimension.GetBoundaryBox_return_value"
        )
        attempts.append({
            "method": "DrawingDimension.GetBoundaryBox_return_value",
            "succeeded": True,
            "values": values,
            "error": None,
        })
        box = _build_dimension_box(
            values,
            index,
            dim,
            "DrawingDimension.GetBoundaryBox_return_value",
            attempts,
        )
        box.boundary_attempts = list(attempts)
        return box
    except Exception as exc:
        attempts.append({
            "method": "DrawingDimension.GetBoundaryBox_return_value",
            "succeeded": False,
            "values": None,
            "error": _format_com_error(exc),
        })

    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore

        payload = VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
            [0.0] * 8,
        )
        dim.GetBoundaryBox(payload)
        values = _numeric_sequence(
            payload.value,
            8,
            "DrawingDimension.GetBoundaryBox_typed_BYREF_VARIANT",
        )
        attempts.append({
            "method": "DrawingDimension.GetBoundaryBox_typed_BYREF_VARIANT",
            "succeeded": True,
            "values": values,
            "error": None,
        })
        box = _build_dimension_box(
            values,
            index,
            dim,
            "DrawingDimension.GetBoundaryBox_typed_BYREF_VARIANT",
            attempts,
        )
        box.boundary_attempts = list(attempts)
        return box
    except Exception as exc:
        attempts.append({
            "method": "DrawingDimension.GetBoundaryBox_typed_BYREF_VARIANT",
            "succeeded": False,
            "values": None,
            "error": _format_com_error(exc),
        })

    # A mutable Python list is retained only as a diagnostic fallback.  An
    # unchanged all-zero list is ambiguous and is never accepted as a box.
    try:
        values = [0.0] * 8
        returned = dim.GetBoundaryBox(values)
        candidate = values if returned is None else returned
        numbers = _numeric_sequence(
            candidate,
            8,
            "DrawingDimension.GetBoundaryBox_mutable_array",
        )
        attempts.append({
            "method": "DrawingDimension.GetBoundaryBox_mutable_array",
            "succeeded": False,
            "values": numbers,
            "error": "Mutable-array output is not trusted unless it produces a non-zero box.",
        })
        box = _build_dimension_box(
            numbers,
            index,
            dim,
            "DrawingDimension.GetBoundaryBox_mutable_array",
            attempts,
        )
        box.boundary_attempts = list(attempts)
        return box
    except Exception as exc:
        attempts.append({
            "method": "DrawingDimension.GetBoundaryBox_mutable_array",
            "succeeded": False,
            "values": None,
            "error": _format_com_error(exc),
        })

    raise RuntimeError(
        "DrawingDimension.GetBoundaryBox could not be read reliably. "
        f"Attempts: {attempts}"
    )

def _collect_dimensions(
    application: Any,
    view: Any,
) -> Tuple[List[_DimensionRecord], List[str], int]:
    warnings: List[str] = []
    try:
        dimensions = view.Dimensions
        total = int(dimensions.Count)
    except Exception as exc:
        return [], [f"DrawingView.Dimensions unavailable: {_format_com_error(exc)}"], 0
    records: List[_DimensionRecord] = []
    for index in range(1, total + 1):
        try:
            obj = dimensions.Item(index)
            records.append(_DimensionRecord(
                index=index,
                obj=obj,
                box=_dimension_boundary(application, obj, index),
                dim_type=_safe_int(_safe_attr(obj, "DimType", None)),
            ))
        except Exception as exc:
            warnings.append(
                f"Dimension {index} skipped: {_format_com_error(exc)}"
            )
    return records, warnings, total

_ANCHOR_LEFT = {1, 2, 3, 11, 12, 13}
_ANCHOR_CENTER = {4, 5, 6, 14, 15, 16}
_ANCHOR_RIGHT = {7, 8, 9, 17, 18, 19}
_ANCHOR_TOP = {1, 4, 7, 11, 14, 17}
_ANCHOR_MIDDLE = {2, 5, 8, 12, 15, 18}
_ANCHOR_BOTTOM = {3, 6, 9, 13, 16, 19}


def _font_size_mm(text: Any) -> Tuple[float, str]:
    try:
        size = float(text.GetFontSize(0, 0))
        if math.isfinite(size) and size > 0.0:
            return size, "DrawingText.GetFontSize"
    except Exception:
        pass
    try:
        size = float(text.TextProperties.FontSize)
        if math.isfinite(size) and size > 0.0:
            return size, "DrawingText.TextProperties.FontSize"
    except Exception:
        pass
    return 3.5, "default_3.5mm"


def _glyph_width_units(content: str) -> float:
    """Conservative width model supporting Latin, tabs and full-width glyphs."""
    units = 0.0
    for character in content:
        if character == "\t":
            units += 4.0 * 0.62
        elif unicodedata.combining(character):
            continue
        elif unicodedata.east_asian_width(character) in {"W", "F"}:
            units += 1.0
        elif character.isspace():
            units += 0.45
        else:
            units += 0.62
    return units


def _estimated_text_boundary(
    text_obj: Any,
    index: int,
) -> Tuple[AnnotationBoundingBox, float, float, Dict[str, Any]]:
    # DrawingText.x/y properties are metres; MCP values are millimetres.
    x_mm = _finite_float(text_obj.x, "DrawingText.x") * _M_TO_MM
    y_mm = _finite_float(text_obj.y, "DrawingText.y") * _M_TO_MM
    content = str(_safe_attr(text_obj, "Text", ""))
    font_mm, font_method = _font_size_mm(text_obj)
    logical_lines = content.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    if not logical_lines:
        logical_lines = [""]
    line_widths = [_glyph_width_units(line) * font_mm for line in logical_lines]
    width_mm = max(1.0, max(line_widths, default=0.0))
    line_count = len(logical_lines)

    wrapping_width = _safe_attr(text_obj, "WrappingWidth", 0.0)
    try:
        wrapping_mm = float(wrapping_width)
    except (TypeError, ValueError):
        wrapping_mm = 0.0
    if math.isfinite(wrapping_mm) and wrapping_mm > 0.0 and width_mm > wrapping_mm:
        wrapped_line_count = sum(
            max(1, int(math.ceil(line_width / wrapping_mm)))
            for line_width in line_widths
        )
        line_count = max(line_count, wrapped_line_count)
        width_mm = wrapping_mm

    height_mm = max(font_mm * 1.25, line_count * font_mm * 1.25)
    frame_type = _safe_int(_safe_attr(text_obj, "FrameType", 0)) or 0
    if frame_type != 0:
        padding = max(0.5, font_mm * 0.25)
        width_mm += 2.0 * padding
        height_mm += 2.0 * padding

    anchor = _safe_int(_safe_attr(text_obj, "AnchorPosition", None))
    if anchor in _ANCHOR_LEFT:
        x0, x1 = 0.0, width_mm
    elif anchor in _ANCHOR_RIGHT:
        x0, x1 = -width_mm, 0.0
    else:
        x0, x1 = -width_mm / 2.0, width_mm / 2.0
    if anchor in _ANCHOR_TOP:
        y0, y1 = -height_mm, 0.0
    elif anchor in _ANCHOR_BOTTOM:
        y0, y1 = 0.0, height_mm
    else:
        y0, y1 = -height_mm / 2.0, height_mm / 2.0

    angle = _safe_attr(text_obj, "Angle", 0.0)
    try:
        angle_rad = float(angle)
        if not math.isfinite(angle_rad):
            angle_rad = 0.0
    except (TypeError, ValueError):
        angle_rad = 0.0
    cos_a, sin_a = math.cos(angle_rad), math.sin(angle_rad)
    rotated = [
        (x_mm + px * cos_a - py * sin_a, y_mm + px * sin_a + py * cos_a)
        for px, py in ((x0, y0), (x1, y0), (x1, y1), (x0, y1))
    ]
    xs = [point[0] for point in rotated]
    ys = [point[1] for point in rotated]
    name = str(_safe_attr(text_obj, "Name", "")) or content[:40] or f"Text.{index}"
    box = AnnotationBoundingBox(
        min(xs),
        min(ys),
        max(xs),
        max(ys),
        item_id=f"text:{index}",
        index=index,
        kind="text",
        name=name,
        text_content=content,
        boundary_method="estimated_from_unicode_font_anchor_rotation",
        boundary_verified=False,
    )
    return box, x_mm, y_mm, {
        "font_size_mm": font_mm,
        "font_read_method": font_method,
        "anchor_position_code": anchor,
        "angle_rad": angle_rad,
        "frame_type_code": frame_type,
        "wrapping_width_mm": wrapping_mm if wrapping_mm > 0.0 else None,
        "line_count_estimate": line_count,
        "glyph_width_model": "Unicode East Asian width + conservative Latin factors",
    }

def _collect_texts(
    view: Any,
) -> Tuple[List[_TextRecord], List[str], int, Dict[int, Dict[str, Any]]]:
    warnings: List[str] = []
    metrics: Dict[int, Dict[str, Any]] = {}
    try:
        texts = view.Texts
        total = int(texts.Count)
    except Exception as exc:
        return [], [f"DrawingView.Texts unavailable: {_format_com_error(exc)}"], 0, metrics
    records: List[_TextRecord] = []
    for index in range(1, total + 1):
        try:
            obj = texts.Item(index)
            box, x_mm, y_mm, info = _estimated_text_boundary(obj, index)
            records.append(_TextRecord(index, obj, x_mm, y_mm, box))
            metrics[index] = info
        except Exception as exc:
            warnings.append(f"DrawingText {index} skipped: {_format_com_error(exc)}")
    return records, warnings, total, metrics


def _conflicts(
    movable_texts: Sequence[_TextRecord],
    dimensions: Sequence[_DimensionRecord],
    gap_mm: float,
    *,
    include_dimension_pairs: bool = False,
) -> List[Dict[str, Any]]:
    pairs: List[Tuple[AnnotationBoundingBox, AnnotationBoundingBox, bool]] = []
    for left in range(len(movable_texts)):
        for right in range(left + 1, len(movable_texts)):
            pairs.append((movable_texts[left].box, movable_texts[right].box, True))
    for text in movable_texts:
        for dimension in dimensions:
            pairs.append((text.box, dimension.box, False))
    if include_dimension_pairs:
        for left in range(len(dimensions)):
            for right in range(left + 1, len(dimensions)):
                pairs.append((dimensions[left].box, dimensions[right].box, False))

    result: List[Dict[str, Any]] = []
    for first, second, both_movable in pairs:
        conflict, depth, axis = first.conflict(second, gap_mm)
        if not conflict:
            continue
        severity = "minor" if depth < 2.0 else ("moderate" if depth < 5.0 else "severe")
        result.append({
            "pair_id": "|".join(sorted((first.item_id, second.item_id))),
            "item_a": first.item_id,
            "item_b": second.item_id,
            "required_clearance_move_mm": round(depth, 6),
            "minimum_separation_mm": round(depth, 6),
            "preferred_move_axis": axis,
            "severity": severity,
            "both_movable": both_movable,
        })
    return result

def _update_drawing(document: Any, sheet: Any) -> Tuple[List[Dict[str, Any]], List[str]]:
    attempts: List[Dict[str, Any]] = []
    warnings: List[str] = []
    for name, action in (
        ("DrawingSheet.Update", lambda: sheet.Update()),
        ("DrawingDocument.Update", lambda: document.Update()),
    ):
        try:
            action()
            attempts.append({"method": name, "succeeded": True, "error": None})
        except Exception as exc:
            message = _format_com_error(exc)
            attempts.append({"method": name, "succeeded": False, "error": message})
            warnings.append(f"{name} failed: {message}")
    return attempts, warnings


def _clear_selection(document: Any) -> Optional[str]:
    try:
        document.Selection.Clear()
        return None
    except Exception as exc:
        return _format_com_error(exc)


def analyze_annotation_layout(
    catia_app: Any,
    view_name: Optional[str] = None,
    sheet_index: Optional[int] = None,
    tolerance_mm: float = 1.5,
    include_dimensions: bool = True,
) -> Dict[str, Any]:
    """Return text/dimension boundaries and conflicts without changing CATIA."""
    gap = _finite_float(tolerance_mm, "tolerance_mm")
    if gap < 0.0:
        raise ValueError("tolerance_mm cannot be negative.")
    if not isinstance(include_dimensions, bool):
        raise ValueError("include_dimensions must be true or false.")
    view, sheet, document, target = _get_drawing_view(
        catia_app, view_name, sheet_index
    )
    texts, text_warnings, text_total, text_metrics = _collect_texts(view)
    dimensions: List[_DimensionRecord] = []
    dimension_warnings: List[str] = []
    dimension_total = 0
    if include_dimensions:
        dimensions, dimension_warnings, dimension_total = _collect_dimensions(
            catia_app, view
        )
    warnings = text_warnings + dimension_warnings
    conflicts = _conflicts(
        texts,
        dimensions,
        gap,
        include_dimension_pairs=include_dimensions,
    )
    if texts:
        warnings.append(
            "DrawingText has no exact boundary API; text boxes are conservative estimates."
        )
    return _success({
        "operation": "analyze_annotation_layout",
        "target": target,
        "minimum_gap_mm": gap,
        "include_dimensions": include_dimensions,
        "texts_total_in_collection": text_total,
        "texts_readable": len(texts),
        "text_readback_complete": len(texts) == text_total,
        "dimensions_total_in_collection": dimension_total,
        "dimensions_readable": len(dimensions),
        "dimension_readback_complete": len(dimensions) == dimension_total,
        "text_boxes": [record.box.as_dict() for record in texts],
        "text_metrics": {str(key): value for key, value in text_metrics.items()},
        "dimension_value_boxes": [record.box.as_dict() for record in dimensions],
        "conflicts": conflicts,
        "conflict_count": len(conflicts),
        "conflict_scope": "text-text, text-dimension and dimension-dimension",
        "coordinate_system": "DrawingView local coordinates",
        "mcp_unit": "mm",
        "drawing_text_position_com_unit": "m",
        "model_modified": False,
        "document_save_required": False,
        "document_saved": _document_saved(document),
    }, warnings)

def _normalize_alignment(
    direction: str,
    align_to: str,
) -> Tuple[str, Optional[str]]:
    if direction == "horizontal":
        if align_to not in {"bottom", "top", "center"}:
            raise ValueError(
                "For horizontal layout, align_to must be 'bottom', 'top' or 'center'."
            )
        return align_to, None
    if align_to in {"left", "right", "center"}:
        return align_to, None
    if align_to == "bottom":
        return "left", "For vertical layout, align_to='bottom' is a deprecated alias for 'left'."
    if align_to == "top":
        return "right", "For vertical layout, align_to='top' is a deprecated alias for 'right'."
    raise ValueError(
        "For vertical layout, align_to must be 'left', 'right' or 'center'."
    )


def _alignment_targets(
    dimensions: Sequence[_DimensionRecord],
    spacing_mm: float,
    direction: str,
    align_to: str,
) -> Tuple[List[Tuple[_DimensionRecord, float, float]], str]:
    if direction == "horizontal":
        ordered = sorted(dimensions, key=lambda item: (item.box.center_x, item.index))
        if align_to == "bottom":
            reference = min(item.box.y_min for item in ordered)
            y_for = lambda item: reference + item.box.height / 2.0
            alignment_basis = "common bottom edge"
        elif align_to == "top":
            reference = max(item.box.y_max for item in ordered)
            y_for = lambda item: reference - item.box.height / 2.0
            alignment_basis = "common top edge"
        else:
            reference = median(item.box.center_y for item in ordered)
            y_for = lambda item: reference
            alignment_basis = "common vertical center"
        cursor = ordered[0].box.x_min
        targets = []
        for item in ordered:
            center_x = cursor + item.box.width / 2.0
            targets.append((item, center_x, y_for(item)))
            cursor += item.box.width + spacing_mm
        return targets, alignment_basis

    ordered = sorted(dimensions, key=lambda item: (item.box.center_y, item.index))
    if align_to == "left":
        reference = min(item.box.x_min for item in ordered)
        x_for = lambda item: reference + item.box.width / 2.0
        alignment_basis = "common left edge"
    elif align_to == "right":
        reference = max(item.box.x_max for item in ordered)
        x_for = lambda item: reference - item.box.width / 2.0
        alignment_basis = "common right edge"
    else:
        reference = median(item.box.center_x for item in ordered)
        x_for = lambda item: reference
        alignment_basis = "common horizontal center"
    cursor = ordered[0].box.y_min
    targets = []
    for item in ordered:
        center_y = cursor + item.box.height / 2.0
        targets.append((item, x_for(item), center_y))
        cursor += item.box.height + spacing_mm
    return targets, alignment_basis


def _dimension_layout_violations(
    dimensions: Sequence[_DimensionRecord],
    spacing_mm: float,
    direction: str,
    align_to: str,
    tolerance_mm: float,
) -> List[Dict[str, Any]]:
    violations: List[Dict[str, Any]] = []
    if not dimensions:
        return violations
    if direction == "horizontal":
        ordered = sorted(dimensions, key=lambda item: (item.box.center_x, item.index))
        if align_to == "bottom":
            values = [item.box.y_min for item in ordered]
        elif align_to == "top":
            values = [item.box.y_max for item in ordered]
        else:
            values = [item.box.center_y for item in ordered]
        for item, value in zip(ordered, values):
            if abs(value - median(values)) > tolerance_mm:
                violations.append({
                    "type": "alignment",
                    "dimension_index": item.index,
                    "deviation_mm": abs(value - median(values)),
                })
        for first, second in zip(ordered, ordered[1:]):
            actual_gap = second.box.x_min - first.box.x_max
            if actual_gap < spacing_mm - tolerance_mm:
                violations.append({
                    "type": "edge_gap",
                    "first_dimension_index": first.index,
                    "second_dimension_index": second.index,
                    "actual_gap_mm": actual_gap,
                    "required_gap_mm": spacing_mm,
                })
        return violations

    ordered = sorted(dimensions, key=lambda item: (item.box.center_y, item.index))
    if align_to == "left":
        values = [item.box.x_min for item in ordered]
    elif align_to == "right":
        values = [item.box.x_max for item in ordered]
    else:
        values = [item.box.center_x for item in ordered]
    for item, value in zip(ordered, values):
        if abs(value - median(values)) > tolerance_mm:
            violations.append({
                "type": "alignment",
                "dimension_index": item.index,
                "deviation_mm": abs(value - median(values)),
            })
    for first, second in zip(ordered, ordered[1:]):
        actual_gap = second.box.y_min - first.box.y_max
        if actual_gap < spacing_mm - tolerance_mm:
            violations.append({
                "type": "edge_gap",
                "first_dimension_index": first.index,
                "second_dimension_index": second.index,
                "actual_gap_mm": actual_gap,
                "required_gap_mm": spacing_mm,
            })
    return violations


def _move_dimension_center_verified(
    application: Any,
    record: _DimensionRecord,
    target_x_mm: float,
    target_y_mm: float,
    document: Any,
    sheet: Any,
    tolerance_mm: float,
    maximum_attempts: int = 3,
) -> Dict[str, Any]:
    """Move a dimension value using readback feedback, not a center-point assumption."""
    command_x = target_x_mm
    command_y = target_y_mm
    attempts: List[Dict[str, Any]] = []
    for attempt_index in range(1, maximum_attempts + 1):
        try:
            record.obj.MoveValue(command_x, command_y, 1, 1)
            update_attempts, update_warnings = _update_drawing(document, sheet)
            box = _dimension_boundary(application, record.obj, record.index)
            error_x = box.center_x - target_x_mm
            error_y = box.center_y - target_y_mm
            error_mm = math.hypot(error_x, error_y)
            verified = error_mm <= tolerance_mm
            attempts.append({
                "attempt": attempt_index,
                "command_position_mm": [command_x, command_y],
                "actual_box_center_mm": [box.center_x, box.center_y],
                "target_box_center_mm": [target_x_mm, target_y_mm],
                "error_vector_mm": [error_x, error_y],
                "error_mm": error_mm,
                "verified": verified,
                "update_attempts": update_attempts,
                "update_warnings": update_warnings,
                "error": None,
            })
            if verified:
                return {
                    "verified": True,
                    "attempts": attempts,
                    "actual_box": box.as_dict(),
                    "error": None,
                }
            command_x -= error_x
            command_y -= error_y
        except Exception as exc:
            attempts.append({
                "attempt": attempt_index,
                "command_position_mm": [command_x, command_y],
                "verified": False,
                "error": _format_com_error(exc),
            })
            break
    return {
        "verified": False,
        "attempts": attempts,
        "actual_box": None,
        "error": "Dimension value did not reach the requested box center.",
    }


def _rollback_dimensions(
    application: Any,
    moved: Sequence[Tuple[_DimensionRecord, float, float]],
    document: Any,
    sheet: Any,
    tolerance_mm: float,
) -> Dict[str, Any]:
    details: List[Dict[str, Any]] = []
    all_verified = True
    for record, old_x, old_y in reversed(list(moved)):
        result = _move_dimension_center_verified(
            application,
            record,
            old_x,
            old_y,
            document,
            sheet,
            tolerance_mm,
        )
        all_verified = all_verified and bool(result["verified"])
        details.append({
            "dimension_index": record.index,
            "target_original_center_mm": [old_x, old_y],
            "verified": result["verified"],
            "attempts": result["attempts"],
            "error": result["error"],
        })
    return {
        "attempted": bool(moved),
        "succeeded": all_verified,
        "details": details,
        "verification": details,
        "update_attempts": [],
        "warnings": [],
    }


def auto_arrange_dimensions(
    catia_app: Any,
    view_name: Optional[str] = None,
    sheet_index: Optional[int] = None,
    spacing_mm: float = 8.0,
    tolerance_mm: float = 2.0,
    direction: str = "horizontal",
    align_to: str = "bottom",
    readback_tolerance_mm: float = 0.05,
    atomic: bool = True,
    strict_readback: bool = True,
) -> Dict[str, Any]:
    """Arrange dimension values with exact bounds and transactional verification."""
    spacing = _finite_float(spacing_mm, "spacing_mm")
    movement_tolerance = _finite_float(tolerance_mm, "tolerance_mm")
    readback_tolerance = _finite_float(
        readback_tolerance_mm, "readback_tolerance_mm"
    )
    if spacing < 0.0:
        raise ValueError("spacing_mm cannot be negative.")
    if movement_tolerance < 0.0:
        raise ValueError("tolerance_mm cannot be negative.")
    if readback_tolerance <= 0.0:
        raise ValueError("readback_tolerance_mm must be greater than zero.")
    if not isinstance(atomic, bool):
        raise ValueError("atomic must be true or false.")
    if not isinstance(strict_readback, bool):
        raise ValueError("strict_readback must be true or false.")

    direction = str(direction).strip().lower()
    align_to_requested = str(align_to).strip().lower()
    if direction not in {"horizontal", "vertical"}:
        raise ValueError("direction must be 'horizontal' or 'vertical'.")
    align_to_resolved, alias_warning = _normalize_alignment(
        direction, align_to_requested
    )

    view, sheet, document, target = _get_drawing_view(
        catia_app, view_name, sheet_index
    )
    saved_before = _document_saved(document)
    selection_warning = _clear_selection(document)
    warnings: List[str] = []
    if alias_warning:
        warnings.append(alias_warning)
    if selection_warning:
        warnings.append(f"Selection.Clear before operation failed: {selection_warning}")
    try:
        dimensions, collect_warnings, total = _collect_dimensions(catia_app, view)
        warnings.extend(collect_warnings)
        if total == 0:
            return _success({
                "operation": "auto_arrange_dimensions",
                "result": "no_dimensions",
                "target": target,
                "total_dimensions": 0,
                "model_modified": False,
                "document_save_required": False,
                "document_saved": _document_saved(document),
            }, warnings)
        if not dimensions:
            raise ToolOperationError(
                "Dimensions exist, but none returned a verified GetBoundaryBox result.",
                data={"target": target, "total_dimensions": total},
                warnings=warnings,
            )
        if strict_readback and len(dimensions) != total:
            raise ToolOperationError(
                "Dimension readback is incomplete; no layout changes were made.",
                data={
                    "target": target,
                    "total_dimensions": total,
                    "readable_dimensions": len(dimensions),
                    "strict_readback": True,
                },
                warnings=warnings,
            )

        targets, alignment_basis = _alignment_targets(
            dimensions, spacing, direction, align_to_resolved
        )
        details: List[Dict[str, Any]] = []
        moved: List[Tuple[_DimensionRecord, float, float]] = []
        expected_center_by_index: Dict[int, Tuple[float, float]] = {}
        failed_count = 0
        for record, target_x, target_y in targets:
            old_x, old_y = record.box.center_x, record.box.center_y
            displacement = math.hypot(target_x - old_x, target_y - old_y)
            if displacement <= movement_tolerance:
                expected_center_by_index[record.index] = (old_x, old_y)
                details.append({
                    "dimension_index": record.index,
                    "action": "kept_within_movement_tolerance",
                    "from_center_mm": [round(old_x, 6), round(old_y, 6)],
                    "target_center_mm": [round(target_x, 6), round(target_y, 6)],
                    "displacement_mm": round(displacement, 6),
                })
                continue
            move_result = _move_dimension_center_verified(
                catia_app,
                record,
                target_x,
                target_y,
                document,
                sheet,
                readback_tolerance,
            )
            if move_result["verified"]:
                moved.append((record, old_x, old_y))
                expected_center_by_index[record.index] = (target_x, target_y)
                details.append({
                    "dimension_index": record.index,
                    "action": "moved_and_verified",
                    "from_center_mm": [round(old_x, 6), round(old_y, 6)],
                    "target_center_mm": [round(target_x, 6), round(target_y, 6)],
                    "displacement_mm": round(displacement, 6),
                    "move_api": "DrawingDimension.MoveValue + feedback readback",
                    "feedback_attempts": move_result["attempts"],
                    "error": None,
                })
            else:
                failed_count += 1
                details.append({
                    "dimension_index": record.index,
                    "action": "move_failed_or_unverified",
                    "from_center_mm": [round(old_x, 6), round(old_y, 6)],
                    "target_center_mm": [round(target_x, 6), round(target_y, 6)],
                    "move_api": "DrawingDimension.MoveValue + feedback readback",
                    "feedback_attempts": move_result["attempts"],
                    "error": move_result["error"],
                })
                expected_center_by_index[record.index] = (old_x, old_y)
                if atomic:
                    break

        update_attempts, update_warnings = _update_drawing(document, sheet)
        warnings.extend(update_warnings)
        final_dimensions, final_warnings, _ = _collect_dimensions(catia_app, view)
        warnings.extend(final_warnings)
        final_by_index = {item.index: item.box.as_dict() for item in final_dimensions}
        final_records = {item.index: item for item in final_dimensions}
        readback_details: List[Dict[str, Any]] = []
        readback_failures = 0
        target_by_index = {
            record.index: (target_x, target_y)
            for record, target_x, target_y in targets
        }
        for record in dimensions:
            final = final_records.get(record.index)
            expected_xy = expected_center_by_index.get(
                record.index,
                (record.box.center_x, record.box.center_y),
            )
            requested_target_xy = target_by_index[record.index]
            if final is None:
                readback_failures += 1
                readback_details.append({
                    "dimension_index": record.index,
                    "verified": False,
                    "error": "Dimension boundary unavailable after layout.",
                })
                continue
            error_mm = math.hypot(
                final.box.center_x - expected_xy[0],
                final.box.center_y - expected_xy[1],
            )
            verified = error_mm <= readback_tolerance
            if not verified:
                readback_failures += 1
            readback_details.append({
                "dimension_index": record.index,
                "verified": verified,
                "requested_target_center_mm": [
                    requested_target_xy[0], requested_target_xy[1]
                ],
                "expected_readback_center_mm": [expected_xy[0], expected_xy[1]],
                "actual_center_mm": [final.box.center_x, final.box.center_y],
                "error_mm": error_mm,
            })

        layout_violations = _dimension_layout_violations(
            final_dimensions,
            spacing,
            direction,
            align_to_resolved,
            max(readback_tolerance, movement_tolerance),
        )
        needs_rollback = bool(
            atomic and (failed_count or readback_failures or layout_violations)
        )
        rollback = {
            "attempted": False,
            "succeeded": None,
            "details": [],
            "verification": [],
        }
        if needs_rollback:
            rollback = _rollback_dimensions(
                catia_app,
                moved,
                document,
                sheet,
                readback_tolerance,
            )
            warnings.extend(rollback.get("warnings", []))

        committed_count = 0 if needs_rollback and rollback.get("succeeded") else len(moved)
        data = {
            "operation": "auto_arrange_dimensions",
            "target": target,
            "direction": direction,
            "align_to_requested": align_to_requested,
            "align_to_resolved": align_to_resolved,
            "alignment_basis": alignment_basis,
            "requested_edge_gap_mm": spacing,
            "movement_tolerance_mm": movement_tolerance,
            "readback_tolerance_mm": readback_tolerance,
            "atomic": atomic,
            "strict_readback": strict_readback,
            "total_dimensions": total,
            "readable_dimensions": len(dimensions),
            "move_calls_succeeded": len(moved),
            "committed_moved_count": committed_count,
            "move_call_failures": failed_count,
            "readback_failures": readback_failures,
            "layout_violations": layout_violations,
            "details": details,
            "readback_details": readback_details,
            "final_dimension_value_boxes_before_optional_rollback": final_by_index,
            "rollback": rollback,
            "update_attempts": update_attempts,
            "coordinate_system": "DrawingView local coordinates",
            "mcp_unit": "mm",
            "model_modified": committed_count > 0,
            "document_save_required": committed_count > 0,
            "document_saved_before": saved_before,
            "document_saved_after": _document_saved(document),
        }
        if failed_count or readback_failures or layout_violations:
            return _error(
                "Dimension layout failed exact readback or final spacing/alignment validation.",
                data=data,
                warnings=warnings,
                status="rolled_back" if rollback.get("succeeded") else (
                    "partial_success" if committed_count else "error"
                ),
            )
        return _success(data, warnings)
    finally:
        final_clear_error = _clear_selection(document)
        if final_clear_error:
            logger.warning("Selection.Clear after dimension layout failed: %s", final_clear_error)

def _repel(
    first: AnnotationBoundingBox,
    second: AnnotationBoundingBox,
    gap_mm: float,
    max_step_mm: float,
    move_second: bool,
) -> bool:
    conflict, depth, axis = first.conflict(second, gap_mm)
    if not conflict:
        return False
    amount = min(max_step_mm, depth + _EPSILON_MM)
    if axis == "x":
        direction = 1.0 if second.center_x > first.center_x else -1.0
        if abs(second.center_x - first.center_x) <= _EPSILON_MM:
            direction = 1.0 if second.item_id > first.item_id else -1.0
        if move_second:
            first.translate(-direction * amount / 2.0, 0.0)
            second.translate(direction * amount / 2.0, 0.0)
        else:
            first.translate(-direction * amount, 0.0)
    else:
        direction = 1.0 if second.center_y > first.center_y else -1.0
        if abs(second.center_y - first.center_y) <= _EPSILON_MM:
            direction = 1.0 if second.item_id > first.item_id else -1.0
        if move_second:
            first.translate(0.0, -direction * amount / 2.0)
            second.translate(0.0, direction * amount / 2.0)
        else:
            first.translate(0.0, -direction * amount)
    return True


def _clamp_box_displacement(
    box: AnnotationBoundingBox,
    original: AnnotationBoundingBox,
    maximum_mm: float,
) -> bool:
    dx = box.center_x - original.center_x
    dy = box.center_y - original.center_y
    distance = math.hypot(dx, dy)
    if distance <= maximum_mm + _EPSILON_MM:
        return False
    scale = maximum_mm / distance
    allowed_x = dx * scale
    allowed_y = dy * scale
    box.translate(
        original.center_x + allowed_x - box.center_x,
        original.center_y + allowed_y - box.center_y,
    )
    return True


def _solve_text_layout(
    texts: Sequence[_TextRecord],
    dimensions: Sequence[_DimensionRecord],
    gap_mm: float,
    max_iterations: int,
    move_step_mm: float,
    original_boxes: Dict[int, AnnotationBoundingBox],
    max_total_displacement_mm: float,
) -> Tuple[int, bool, int]:
    iterations_used = 0
    converged = False
    clamped_events = 0
    for iteration in range(1, max_iterations + 1):
        changed = False
        for left in range(len(texts)):
            for right in range(left + 1, len(texts)):
                changed = _repel(
                    texts[left].box,
                    texts[right].box,
                    gap_mm,
                    move_step_mm,
                    True,
                ) or changed
        for text in texts:
            for dimension in dimensions:
                changed = _repel(
                    text.box,
                    dimension.box,
                    gap_mm,
                    move_step_mm,
                    False,
                ) or changed
        for text in texts:
            if _clamp_box_displacement(
                text.box,
                original_boxes[text.index],
                max_total_displacement_mm,
            ):
                clamped_events += 1
        iterations_used = iteration
        remaining = _conflicts(
            texts,
            dimensions,
            gap_mm,
            include_dimension_pairs=False,
        )
        if not remaining:
            converged = True
            break
        if not changed:
            break
    return iterations_used, converged, clamped_events

def _set_text_anchor_mm(
    record: _TextRecord,
    target_x_mm: float,
    target_y_mm: float,
    verify_tolerance_mm: float = 0.01,
) -> Dict[str, Any]:
    obj = record.obj
    old_x_m = float(obj.x)
    old_y_m = float(obj.y)
    rollback = {"attempted": False, "succeeded": None, "error": None}
    error = None
    try:
        obj.x = target_x_mm * _MM_TO_M
        obj.y = target_y_mm * _MM_TO_M
        actual_x_mm = float(obj.x) * _M_TO_MM
        actual_y_mm = float(obj.y) * _M_TO_MM
        verified = math.hypot(
            actual_x_mm - target_x_mm, actual_y_mm - target_y_mm
        ) <= verify_tolerance_mm
    except Exception as exc:
        actual_x_mm = actual_y_mm = None
        verified = False
        error = _format_com_error(exc)
    if not verified:
        rollback["attempted"] = True
        try:
            obj.x = old_x_m
            obj.y = old_y_m
            rollback["succeeded"] = bool(
                abs(float(obj.x) - old_x_m) <= 1.0e-9
                and abs(float(obj.y) - old_y_m) <= 1.0e-9
            )
        except Exception as exc:
            rollback["succeeded"] = False
            rollback["error"] = _format_com_error(exc)
    return {
        "verified": verified,
        "actual_anchor_mm": (
            [round(actual_x_mm, 6), round(actual_y_mm, 6)]
            if actual_x_mm is not None and actual_y_mm is not None else None
        ),
        "rollback": rollback,
        "error": error,
    }


def _rollback_texts(
    moved: Sequence[_TextRecord],
    document: Any,
    sheet: Any,
) -> Dict[str, Any]:
    details: List[Dict[str, Any]] = []
    succeeded = True
    for record in reversed(list(moved)):
        try:
            record.obj.x = record.anchor_x_mm * _MM_TO_M
            record.obj.y = record.anchor_y_mm * _MM_TO_M
            verified = (
                abs(float(record.obj.x) * _M_TO_MM - record.anchor_x_mm) <= 0.01
                and abs(float(record.obj.y) * _M_TO_MM - record.anchor_y_mm) <= 0.01
            )
            succeeded = succeeded and verified
            details.append({
                "text_index": record.index,
                "verified": verified,
                "error": None,
            })
        except Exception as exc:
            succeeded = False
            details.append({
                "text_index": record.index,
                "verified": False,
                "error": _format_com_error(exc),
            })
    update_attempts, update_warnings = _update_drawing(document, sheet)
    return {
        "attempted": bool(moved),
        "succeeded": succeeded,
        "details": details,
        "update_attempts": update_attempts,
        "warnings": update_warnings,
    }


def fix_overlapping_annotations(
    catia_app: Any,
    view_name: Optional[str] = None,
    sheet_index: Optional[int] = None,
    tolerance_mm: float = 1.5,
    max_iterations: int = 20,
    move_step_mm: float = 5.0,
    include_dimensions: bool = True,
    max_total_displacement_mm: float = 40.0,
    atomic: bool = True,
    strict_readback: bool = True,
) -> Dict[str, Any]:
    """Move drawing texts with displacement caps and transactional readback."""
    gap = _finite_float(tolerance_mm, "tolerance_mm")
    step = _finite_float(move_step_mm, "move_step_mm")
    maximum_displacement = _finite_float(
        max_total_displacement_mm, "max_total_displacement_mm"
    )
    iterations = _positive_int(max_iterations, "max_iterations", 500)
    if gap < 0.0:
        raise ValueError("tolerance_mm cannot be negative.")
    if step <= 0.0:
        raise ValueError("move_step_mm must be greater than zero.")
    if maximum_displacement <= 0.0:
        raise ValueError("max_total_displacement_mm must be greater than zero.")
    if not isinstance(include_dimensions, bool):
        raise ValueError("include_dimensions must be true or false.")
    if not isinstance(atomic, bool):
        raise ValueError("atomic must be true or false.")
    if not isinstance(strict_readback, bool):
        raise ValueError("strict_readback must be true or false.")

    view, sheet, document, target = _get_drawing_view(
        catia_app, view_name, sheet_index
    )
    saved_before = _document_saved(document)
    warnings: List[str] = []
    selection_warning = _clear_selection(document)
    if selection_warning:
        warnings.append(f"Selection.Clear before operation failed: {selection_warning}")
    try:
        texts, text_warnings, text_total, _ = _collect_texts(view)
        warnings.extend(text_warnings)
        dimensions: List[_DimensionRecord] = []
        dimension_total = 0
        if include_dimensions:
            dimensions, dim_warnings, dimension_total = _collect_dimensions(
                catia_app, view
            )
            warnings.extend(dim_warnings)
        if text_total < 1:
            return _success({
                "operation": "fix_overlapping_annotations",
                "result": "no_text_annotations",
                "target": target,
                "model_modified": False,
                "document_save_required": False,
                "document_saved": _document_saved(document),
            }, warnings)
        if not texts:
            raise ToolOperationError(
                "Drawing texts exist, but no text position could be read safely.",
                data={"target": target, "text_count": text_total},
                warnings=warnings,
            )
        if strict_readback and len(texts) != text_total:
            raise ToolOperationError(
                "Text readback is incomplete; no annotation moves were made.",
                data={
                    "target": target,
                    "texts_total": text_total,
                    "texts_readable": len(texts),
                },
                warnings=warnings,
            )
        if strict_readback and include_dimensions and len(dimensions) != dimension_total:
            raise ToolOperationError(
                "Dimension obstacle readback is incomplete; no annotation moves were made.",
                data={
                    "target": target,
                    "dimensions_total": dimension_total,
                    "dimensions_readable": len(dimensions),
                },
                warnings=warnings,
            )

        warnings.append(
            "DrawingText has no exact boundary API; avoidance uses conservative estimated boxes."
        )
        warnings.append(
            "The solver caps total displacement but does not independently clip text boxes to the paper boundary."
        )
        original_boxes = {
            item.index: AnnotationBoundingBox(**item.box.__dict__) for item in texts
        }
        initial_conflicts = _conflicts(
            texts,
            dimensions,
            gap,
            include_dimension_pairs=False,
        )
        if not initial_conflicts:
            return _success({
                "operation": "fix_overlapping_annotations",
                "result": "no_conflicts",
                "target": target,
                "texts_total_in_collection": text_total,
                "texts_readable": len(texts),
                "dimensions_total_in_collection": dimension_total,
                "dimensions_readable": len(dimensions),
                "overlaps_detected": 0,
                "overlaps_fixed": 0,
                "overlaps_remaining": 0,
                "model_modified": False,
                "document_save_required": False,
                "document_saved": _document_saved(document),
            }, warnings)

        iterations_used, solver_converged, clamped_events = _solve_text_layout(
            texts,
            dimensions,
            gap,
            iterations,
            step,
            original_boxes,
            maximum_displacement,
        )
        simulated_remaining = _conflicts(
            texts,
            dimensions,
            gap,
            include_dimension_pairs=False,
        )
        if atomic and simulated_remaining:
            return _error(
                "The estimated text layout could not resolve all conflicts within the displacement cap.",
                data={
                    "operation": "fix_overlapping_annotations",
                    "target": target,
                    "initial_conflicts": initial_conflicts,
                    "simulated_remaining_conflicts": simulated_remaining,
                    "iterations_used": iterations_used,
                    "solver_converged": solver_converged,
                    "clamped_events": clamped_events,
                    "model_modified": False,
                    "document_save_required": False,
                },
                warnings=warnings,
                status="no_change",
            )

        move_details: List[Dict[str, Any]] = []
        moved_records: List[_TextRecord] = []
        move_failures = 0
        for record in texts:
            original = original_boxes[record.index]
            dx = record.box.x_min - original.x_min
            dy = record.box.y_min - original.y_min
            if math.hypot(dx, dy) <= _EPSILON_MM:
                continue
            target_x = record.anchor_x_mm + dx
            target_y = record.anchor_y_mm + dy
            set_result = _set_text_anchor_mm(record, target_x, target_y)
            if set_result["verified"]:
                moved_records.append(record)
            else:
                move_failures += 1
            move_details.append({
                "text_index": record.index,
                "text_id": record.box.item_id,
                "from_anchor_mm": [
                    round(record.anchor_x_mm, 6), round(record.anchor_y_mm, 6)
                ],
                "target_anchor_mm": [round(target_x, 6), round(target_y, 6)],
                "delta_mm": [round(dx, 6), round(dy, 6)],
                "total_displacement_mm": round(math.hypot(dx, dy), 6),
                "position_api": "DrawingText.x/y (metres)",
                "readback_verified": set_result["verified"],
                "actual_anchor_mm": set_result["actual_anchor_mm"],
                "rollback": set_result["rollback"],
                "error": set_result["error"],
            })
            if atomic and not set_result["verified"]:
                break

        update_attempts, update_warnings = _update_drawing(document, sheet)
        warnings.extend(update_warnings)
        final_texts, final_text_warnings, _, _ = _collect_texts(view)
        warnings.extend(final_text_warnings)
        final_dimensions: List[_DimensionRecord] = []
        if include_dimensions:
            final_dimensions, final_dim_warnings, _ = _collect_dimensions(
                catia_app, view
            )
            warnings.extend(final_dim_warnings)
        remaining = _conflicts(
            final_texts,
            final_dimensions,
            gap,
            include_dimension_pairs=False,
        )
        initial_ids = {item["pair_id"] for item in initial_conflicts}
        remaining_ids = {item["pair_id"] for item in remaining}
        fixed_count = len(initial_ids - remaining_ids)

        needs_rollback = bool(atomic and (remaining or move_failures))
        rollback = {
            "attempted": False,
            "succeeded": None,
            "details": [],
        }
        if needs_rollback:
            rollback = _rollback_texts(moved_records, document, sheet)
            warnings.extend(rollback.get("warnings", []))

        committed_count = (
            0 if needs_rollback and rollback.get("succeeded") else len(moved_records)
        )
        data = {
            "operation": "fix_overlapping_annotations",
            "target": target,
            "minimum_gap_mm": gap,
            "move_step_mm": step,
            "max_iterations": iterations,
            "max_total_displacement_mm": maximum_displacement,
            "iterations_used": iterations_used,
            "solver_converged_before_com_write": solver_converged,
            "clamped_events": clamped_events,
            "atomic": atomic,
            "strict_readback": strict_readback,
            "include_dimensions": include_dimensions,
            "texts_total_in_collection": text_total,
            "texts_readable": len(texts),
            "dimensions_total_in_collection": dimension_total,
            "dimensions_readable": len(dimensions),
            "move_calls_verified": len(moved_records),
            "committed_moved_count": committed_count,
            "move_failures": move_failures,
            "overlaps_detected": len(initial_conflicts),
            "overlaps_fixed_before_optional_rollback": fixed_count,
            "overlaps_remaining_before_optional_rollback": len(remaining),
            "initial_conflicts": initial_conflicts,
            "simulated_remaining_conflicts": simulated_remaining,
            "remaining_conflicts_before_optional_rollback": remaining,
            "move_details": move_details,
            "final_text_boxes_before_optional_rollback": [
                record.box.as_dict() for record in final_texts
            ],
            "rollback": rollback,
            "update_attempts": update_attempts,
            "coordinate_system": "DrawingView local coordinates",
            "mcp_unit": "mm",
            "drawing_text_position_com_unit": "m",
            "model_modified": committed_count > 0,
            "document_save_required": committed_count > 0,
            "document_saved_before": saved_before,
            "document_saved_after": _document_saved(document),
        }
        if remaining or move_failures:
            return _error(
                "Annotation avoidance completed with unresolved conflicts or failed moves.",
                data=data,
                warnings=warnings,
                status="rolled_back" if rollback.get("succeeded") else (
                    "partial_success" if committed_count else "error"
                ),
            )
        return _success(data, warnings)
    finally:
        final_clear_error = _clear_selection(document)
        if final_clear_error:
            logger.warning("Selection.Clear after annotation layout failed: %s", final_clear_error)

# Retained as machine-readable metadata for integrations that inspect modules.
SMART_ANNOTATION_MCP_TOOLS = [
    {"name": "analyze_annotation_layout", "handler": analyze_annotation_layout},
    {"name": "auto_arrange_dimensions", "handler": auto_arrange_dimensions},
    {"name": "fix_overlapping_annotations", "handler": fix_overlapping_annotations},
]




def _attach_runtime_evidence(result: Dict[str, Any]) -> Dict[str, Any]:
    data = result.get("data")
    if isinstance(data, dict):
        data.setdefault(
            "application_resolution",
            {
                "method": "conn.connect(visible=True)",
                "system_service_expected": True,
            },
        )
    return result

def register_tools(mcp: Any, ctx: Any) -> list[str]:
    """Register tools using the same decorator pattern as the server modules."""
    conn = ctx.conn
    names: list[str] = []
    analyze_impl = globals()["analyze_annotation_layout"]
    arrange_impl = globals()["auto_arrange_dimensions"]
    fix_impl = globals()["fix_overlapping_annotations"]

    @mcp.tool()
    def analyze_annotation_layout(
        view_name: Optional[str] = None,
        sheet_index: Optional[int] = None,
        tolerance_mm: float = 1.5,
        include_dimensions: bool = True,
    ) -> Dict[str, Any]:
        """Preview estimated text and exact dimension-value conflicts."""
        try:
            app = conn.connect(visible=True)
            return _attach_runtime_evidence(analyze_impl(
                app, view_name, sheet_index, tolerance_mm, include_dimensions
            ))
        except ToolOperationError as exc:
            return _error(
                str(exc), data=exc.data, warnings=exc.warnings, status=exc.status
            )
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("analyze_annotation_layout")

    @mcp.tool()
    def auto_arrange_dimensions(
        view_name: Optional[str] = None,
        sheet_index: Optional[int] = None,
        spacing_mm: float = 8.0,
        tolerance_mm: float = 2.0,
        direction: str = "horizontal",
        align_to: str = "bottom",
        readback_tolerance_mm: float = 0.05,
        atomic: bool = True,
        strict_readback: bool = True,
    ) -> Dict[str, Any]:
        """Arrange dimension values with exact readback and atomic rollback."""
        try:
            app = conn.connect(visible=True)
            return _attach_runtime_evidence(arrange_impl(
                app,
                view_name,
                sheet_index,
                spacing_mm,
                tolerance_mm,
                direction,
                align_to,
                readback_tolerance_mm,
                atomic,
                strict_readback,
            ))
        except ToolOperationError as exc:
            return _error(
                str(exc), data=exc.data, warnings=exc.warnings, status=exc.status
            )
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("auto_arrange_dimensions")

    @mcp.tool()
    def fix_overlapping_annotations(
        view_name: Optional[str] = None,
        sheet_index: Optional[int] = None,
        tolerance_mm: float = 1.5,
        max_iterations: int = 20,
        move_step_mm: float = 5.0,
        include_dimensions: bool = True,
        max_total_displacement_mm: float = 40.0,
        atomic: bool = True,
        strict_readback: bool = True,
    ) -> Dict[str, Any]:
        """Move texts with displacement caps, exact readback and rollback."""
        try:
            app = conn.connect(visible=True)
            return _attach_runtime_evidence(fix_impl(
                app,
                view_name,
                sheet_index,
                tolerance_mm,
                max_iterations,
                move_step_mm,
                include_dimensions,
                max_total_displacement_mm,
                atomic,
                strict_readback,
            ))
        except ToolOperationError as exc:
            return _error(
                str(exc), data=exc.data, warnings=exc.warnings, status=exc.status
            )
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("fix_overlapping_annotations")
    return names
