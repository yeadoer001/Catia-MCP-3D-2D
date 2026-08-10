"""
drawing_template_tools.py
Version: drawing-template-tools-fixed-2026-08-05-v7

Verified CATIA V5 title-block, projection-standard and drawing-frame tools.

v7 adds managed default-template resources, verified clean-template generation,
template-copy drawing creation, table-aware title fields and file transactions.
"""

from __future__ import annotations

import hashlib
import logging
import math
import os
import re
import shutil
import uuid
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)


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
    result = {
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
    return str(exc)


def _safe_attr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _safe_count(collection: Any) -> Optional[int]:
    if collection is None:
        return None
    try:
        return int(collection.Count)
    except Exception:
        return None


def _finite_float(value: Any, name: str) -> float:
    if isinstance(value, bool):
        raise ValueError(f"{name} must be finite.")
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{name} must be finite.") from exc
    if not math.isfinite(result):
        raise ValueError(f"{name} must be finite.")
    return result


def _document_saved(document: Any) -> Optional[bool]:
    try:
        return bool(document.Saved)
    except Exception:
        return None


def _normalised_path_key(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def _evaluate(
    catia_app: Any,
    script: str,
    function_name: str,
    parameters: List[Any],
) -> Any:
    try:
        return catia_app.SystemService.Evaluate(
            script,
            _CATVB_SCRIPT_LANGUAGE,
            function_name,
            parameters,
        )
    except Exception as exc:
        raise RuntimeError(
            f"SystemService.Evaluate failed for {function_name}: "
            f"{_format_com_error(exc)}"
        ) from exc

# ---------------------------------------------------------------------------
# Constants & Enums
# ---------------------------------------------------------------------------

# v6: paste-first, non-destructive replacement with verified rollback paths.

IMPLEMENTATION_VERSION = "drawing-template-tools-fixed-2026-08-05-v7"
# CATScriptLanguage is zero based: CATVBScriptLanguage=0,
# CATVBALanguage=1.  Evaluate executes inline VBScript, not a VBA project.
_CATVB_SCRIPT_LANGUAGE = 0

PROJECTION_FIRST_ANGLE = "first_angle"
PROJECTION_THIRD_ANGLE = "third_angle"
PROJECTION_UNKNOWN = "unknown"

PAPER_SIZE_NAMES = {
    0: "LETTER",
    1: "LEGAL",
    2: "A0",
    3: "A1",
    4: "A2",
    5: "A3",
    6: "A4",
    13: "USER",
}

ORIENTATION_NAMES = {
    0: "portrait",
    1: "landscape",
}

STANDARD_PAPER_PORTRAIT_MM = {
    "A0": (841.0, 1189.0),
    "A1": (594.0, 841.0),
    "A2": (420.0, 594.0),
    "A3": (297.0, 420.0),
    "A4": (210.0, 297.0),
    "LETTER": (215.9, 279.4),
    "LEGAL": (215.9, 355.6),
}


# Standard Title Block property keys
TITLE_BLOCK_KEYS = [
    "Material", "Weight", "Drafter", "Checked", "Approved", "Date",
    "Scale", "PartNumber", "Description", "SheetFormat", "Revision",
    "Units", "Projection", "Sheet",
]

DEFAULT_TEMPLATE_FILENAME = "TAISHAN_STANDARD_A3_LANDSCAPE.CATDrawing"
DEFAULT_TEMPLATE_ENV = "CATIA_MCP_DRAWING_TEMPLATE"

# ---------------------------------------------------------------------------
# Core Helper Functions
# ---------------------------------------------------------------------------

def _get_drawing_document(catia_app):
    """Return the active CATDrawing, including unsaved Drawing documents."""
    doc = catia_app.ActiveDocument
    if doc is None:
        raise RuntimeError(
            "No active document. Open or create a CATDrawing first."
        )
    try:
        sheets = doc.Sheets
        _ = sheets.Count
    except Exception as exc:
        raise RuntimeError(
            f"Active document '{_safe_attr(doc, 'Name', '')}' "
            "is not a CATDrawing."
        ) from exc
    return doc


def _get_active_sheet(drawing_doc) -> Any:
    sheets = drawing_doc.Sheets
    return sheets.ActiveSheet


def _get_background_view(sheet) -> Any:
    """Return CATIA's system Background View by collection position."""
    views = sheet.Views
    if int(views.Count) < 2:
        raise RuntimeError(
            "Sheet does not contain Main View and Background View."
        )
    view = views.Item(2)
    # DrawingSheet.reorder_Views documents that the first two system views
    # are Main View and Background View.  When ViewType is exposed, verify
    # the second item as catViewBackground (the first enum value, code 0).
    view_type = _safe_attr(view, "ViewType", None)
    if view_type is not None:
        try:
            if int(view_type) != 0:
                raise RuntimeError(
                    "DrawingSheet.Views.Item(2) is not the Background View "
                    f"(ViewType={int(view_type)})."
                )
        except (TypeError, ValueError, OverflowError):
            pass
    return view






def _detect_projection_method(
    catia_app: Any,
    sheet: Any,
) -> Dict[str, Any]:
    """Detect first/third angle using CATIA-side enum constants."""
    script = (
        "Public Function MCP_DetectProjection(sheetObject)\n"
        " If sheetObject.ProjectionMethod = catFirstAngle Then\n"
        "  MCP_DetectProjection = \"first_angle\"\n"
        " ElseIf sheetObject.ProjectionMethod = catThirdAngle Then\n"
        "  MCP_DetectProjection = \"third_angle\"\n"
        " Else\n"
        "  MCP_DetectProjection = \"unknown\"\n"
        " End If\n"
        "End Function"
    )
    actual_code = None
    try:
        actual_code = int(sheet.ProjectionMethod)
    except Exception:
        pass

    # CatSheetProjectionMethod is zero based in the Automation type library.
    # Prefer the direct property, and use CATIA-side symbolic constants only
    # when late binding cannot return a numeric enum value.
    direct_mapping = {
        0: PROJECTION_FIRST_ANGLE,
        1: PROJECTION_THIRD_ANGLE,
    }
    method = direct_mapping.get(actual_code, PROJECTION_UNKNOWN)
    read_method = (
        "DrawingSheet.ProjectionMethod"
        if method != PROJECTION_UNKNOWN else None
    )
    error = None
    if method == PROJECTION_UNKNOWN:
        try:
            evaluated = str(
                _evaluate(
                    catia_app,
                    script,
                    "MCP_DetectProjection",
                    [sheet],
                )
            ).strip().lower()
            if evaluated in {PROJECTION_FIRST_ANGLE, PROJECTION_THIRD_ANGLE}:
                method = evaluated
                read_method = "SystemService.Evaluate_CATVBScriptLanguage"
            else:
                error = f"CATIA returned unrecognised projection value: {evaluated!r}."
        except Exception as exc:
            error = str(exc)

    labels = {
        PROJECTION_FIRST_ANGLE: "First Angle (ISO)",
        PROJECTION_THIRD_ANGLE: "Third Angle (ANSI)",
        PROJECTION_UNKNOWN: "Unknown",
    }
    descriptions = {
        PROJECTION_FIRST_ANGLE: (
            "Top view below Front; right-side view left of Front."
        ),
        PROJECTION_THIRD_ANGLE: (
            "Top view above Front; right-side view right of Front."
        ),
        PROJECTION_UNKNOWN: (
            "Projection method could not be classified."
        ),
    }
    return {
        "method": method,
        "label": labels.get(method, "Unknown"),
        "description": descriptions.get(
            method,
            descriptions[PROJECTION_UNKNOWN],
        ),
        "actual_code": actual_code,
        "read_method": read_method,
        "is_first_angle": method == PROJECTION_FIRST_ANGLE,
        "is_third_angle": method == PROJECTION_THIRD_ANGLE,
        "verified": method in {
            PROJECTION_FIRST_ANGLE,
            PROJECTION_THIRD_ANGLE,
        },
        "read_error": error,
    }



_FIELD_ALIASES = {
    "material": "Material",
    "weight": "Weight",
    "mass": "Weight",
    "drafter": "Drafter",
    "drawnby": "Drafter",
    "designer": "Drafter",
    "date": "Date",
    "scale": "Scale",
    "partnumber": "PartNumber",
    "partno": "PartNumber",
    "drawingnumber": "PartNumber",
    "description": "Description",
    "title": "Description",
    "sheetformat": "SheetFormat",
    "format": "SheetFormat",
    "papersize": "SheetFormat",
    "revision": "Revision",
    "rev": "Revision",
    "checked": "Checked",
    "checker": "Checked",
    "approved": "Approved",
    "approver": "Approved",
    "units": "Units",
    "unit": "Units",
    "projection": "Projection",
    "projectangle": "Projection",
    "projectionangle": "Projection",
    "sheet": "Sheet",
    "sheetnumber": "Sheet",
}


def _resolve_sheet(
    drawing_doc: Any,
    sheet_index: Optional[int],
) -> Tuple[Any, Dict[str, Any]]:
    sheets = drawing_doc.Sheets
    count = int(sheets.Count)
    if sheet_index is None:
        sheet = sheets.ActiveSheet
        index = None
        for i in range(1, count + 1):
            try:
                if sheets.Item(i).Name == sheet.Name:
                    index = i
                    break
            except Exception:
                continue
        method = "DrawingSheets.ActiveSheet"
    else:
        if isinstance(sheet_index, bool):
            raise ValueError("sheet_index must be a positive integer.")
        index = int(sheet_index)
        if index < 1 or index > count:
            raise IndexError(
                f"sheet_index={index} is outside 1..{count}."
            )
        sheet = sheets.Item(index)
        method = f"DrawingSheets.Item({index})"
    return sheet, {
        "selection_method": method,
        "sheet_index": index,
        "sheet_name": str(sheet.Name),
        "sheet_count": count,
    }


def _background_info(view: Any) -> Dict[str, Any]:
    view_type = _safe_attr(view, "ViewType", None)
    return {
        "selection_method": (
            "DrawingSheet.Views.Item(2)_official_system_order"
        ),
        "index": 2,
        "name": str(_safe_attr(view, "Name", "")),
        "view_type_code": (
            int(view_type) if view_type is not None else None
        ),
    }


def _normalise_label(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def _canonical_field(value: str) -> Optional[str]:
    if value in TITLE_BLOCK_KEYS:
        return value
    return _FIELD_ALIASES.get(_normalise_label(value))


def _split_label_value(
    content: str,
) -> Tuple[str, str, Optional[str]]:
    match = re.match(
        r"^\s*([^:=：]+?)\s*([:=：])\s*(.*?)\s*$",
        content,
        flags=re.DOTALL,
    )
    if match:
        return (
            match.group(1).strip(),
            match.group(3).strip(),
            match.group(2),
        )
    return content.strip(), "", None


def _read_text_records(
    background_view: Any,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    warnings: List[str] = []
    records: List[Dict[str, Any]] = []
    texts = background_view.Texts

    for index in range(1, int(texts.Count) + 1):
        try:
            obj = texts.Item(index)
            content = str(obj.Text)
            object_name = str(_safe_attr(obj, "Name", "")).strip()
            label, value, separator = _split_label_value(content)
            canonical_from_content = _canonical_field(label)
            canonical_from_name = _canonical_field(object_name)
            match_mode = None
            canonical_field = None
            if separator is not None and canonical_from_content is not None:
                canonical_field = canonical_from_content
                match_mode = "label_value_text"
            elif canonical_from_name is not None:
                canonical_field = canonical_from_name
                match_mode = "drawing_text_name"
                label = object_name
                value = content
            elif canonical_from_content is not None:
                canonical_field = canonical_from_content
                match_mode = "label_only_text"
            x = _safe_attr(obj, "x", None)
            y = _safe_attr(obj, "y", None)
            linked_count = _safe_attr(obj, "NbLink", None)
            records.append({
                "index": index,
                "object_name": object_name,
                "text": content,
                "label": label,
                "value": value,
                "separator": separator,
                "canonical_field": canonical_field,
                "match_mode": match_mode,
                # DrawingText.x/y are metres in the Automation API.
                "x_mm": float(x) * 1000.0 if x is not None else None,
                "y_mm": float(y) * 1000.0 if y is not None else None,
                "linked_parameter_count": (
                    int(linked_count) if linked_count is not None else None
                ),
            })
        except Exception as exc:
            warnings.append(
                f"DrawingText index {index} could not be read: "
                f"{_format_com_error(exc)}"
            )
    return records, warnings


def _group_text_records(
    records: List[Dict[str, Any]],
) -> Dict[str, List[Dict[str, Any]]]:
    grouped = {key: [] for key in TITLE_BLOCK_KEYS}
    for record in records:
        field_name = record["canonical_field"]
        if field_name in grouped:
            grouped[field_name].append(record)
    return grouped


def _canonicalise_properties(
    properties: Dict[str, str],
) -> Dict[str, str]:
    if not isinstance(properties, dict) or not properties:
        raise ValueError("properties must be a non-empty object.")
    result: Dict[str, str] = {}
    unsupported: List[str] = []
    for raw_key, raw_value in properties.items():
        key = str(raw_key).strip()
        canonical = _canonical_field(key)
        if canonical is None:
            unsupported.append(key)
        else:
            result[canonical] = str(raw_value)
    if unsupported:
        raise ValueError(
            "Unsupported properties: "
            + ", ".join(unsupported)
            + ". Supported: "
            + ", ".join(TITLE_BLOCK_KEYS)
        )
    return result


def _update_drawing(
    drawing_doc: Any,
    sheet: Any,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    attempts: List[Dict[str, Any]] = []
    warnings: List[str] = []
    for name, action in (
        ("DrawingSheet.ForceUpdate", lambda: sheet.ForceUpdate()),
        ("DrawingDocument.Update", lambda: drawing_doc.Update()),
    ):
        try:
            action()
            attempts.append({
                "method": name,
                "succeeded": True,
                "error": None,
            })
        except Exception as exc:
            message = _format_com_error(exc)
            attempts.append({
                "method": name,
                "succeeded": False,
                "error": message,
            })
            warnings.append(f"{name} failed: {message}")
    return attempts, warnings


def _delete_object(
    drawing_doc: Any,
    obj: Any,
) -> Dict[str, Any]:
    selection = drawing_doc.Selection
    try:
        selection.Clear()
        selection.Add(obj)
        selection.Delete()
        selection.Clear()
        return {
            "attempted": True,
            "succeeded": True,
            "error": None,
        }
    except Exception as exc:
        try:
            selection.Clear()
        except Exception:
            pass
        return {
            "attempted": True,
            "succeeded": False,
            "error": _format_com_error(exc),
        }



def _read_sheet_property(
    catia_app: Any,
    sheet: Any,
    property_name: str,
) -> Dict[str, Any]:
    """Read a DrawingSheet scalar property without leaking COM errors."""
    supported = {
        "PaperSize": ("MCP_ReadPaperSize", "CLng(sheetObject.PaperSize)"),
        "Orientation": ("MCP_ReadOrientation", "CLng(sheetObject.Orientation)"),
        "ProjectionMethod": (
            "MCP_ReadProjectionMethod",
            "CLng(sheetObject.ProjectionMethod)",
        ),
    }
    if property_name not in supported:
        raise ValueError(
            f"Unsupported DrawingSheet property: {property_name}."
        )

    attempts: List[Dict[str, Any]] = []
    try:
        value = int(getattr(sheet, property_name))
        attempts.append({
            "method": f"DrawingSheet.{property_name}",
            "succeeded": True,
            "value": value,
            "error": None,
        })
        return {
            "value": value,
            "read_method": attempts[-1]["method"],
            "attempts": attempts,
            "verified": True,
        }
    except Exception as exc:
        attempts.append({
            "method": f"DrawingSheet.{property_name}",
            "succeeded": False,
            "value": None,
            "error": _format_com_error(exc),
        })

    function_name, expression = supported[property_name]
    script = (
        f"Public Function {function_name}(sheetObject)\n"
        f" {function_name} = {expression}\n"
        "End Function"
    )
    try:
        value = int(
            _evaluate(
                catia_app,
                script,
                function_name,
                [sheet],
            )
        )
        attempts.append({
            "method": (
                "SystemService.Evaluate."
                f"DrawingSheet.{property_name}"
            ),
            "succeeded": True,
            "value": value,
            "error": None,
        })
        return {
            "value": value,
            "read_method": attempts[-1]["method"],
            "attempts": attempts,
            "verified": True,
        }
    except Exception as exc:
        attempts.append({
            "method": (
                "SystemService.Evaluate."
                f"DrawingSheet.{property_name}"
            ),
            "succeeded": False,
            "value": None,
            "error": _format_com_error(exc),
        })

    return {
        "value": None,
        "read_method": None,
        "attempts": attempts,
        "verified": False,
    }


def _set_sheet_property(
    catia_app: Any,
    sheet: Any,
    property_name: str,
    value: int,
) -> Dict[str, Any]:
    """Set and verify a DrawingSheet scalar property with CATIA fallback."""
    supported = {
        "PaperSize": "MCP_SetPaperSize",
        "Orientation": "MCP_SetOrientation",
        "ProjectionMethod": "MCP_SetProjectionMethodCode",
    }
    if property_name not in supported:
        raise ValueError(
            f"Unsupported DrawingSheet property: {property_name}."
        )

    requested = int(value)
    attempts: List[Dict[str, Any]] = []

    try:
        setattr(sheet, property_name, requested)
        attempts.append({
            "method": f"DrawingSheet.{property_name}=value",
            "succeeded": True,
            "error": None,
        })
    except Exception as exc:
        attempts.append({
            "method": f"DrawingSheet.{property_name}=value",
            "succeeded": False,
            "error": _format_com_error(exc),
        })

    readback = _read_sheet_property(
        catia_app,
        sheet,
        property_name,
    )
    if readback["verified"] and readback["value"] == requested:
        return {
            "property": property_name,
            "requested": requested,
            "selected_method": attempts[-1]["method"],
            "attempts": attempts,
            "readback": readback,
            "verified": True,
        }

    function_name = supported[property_name]
    script = (
        f"Public Function {function_name}(sheetObject, propertyValue)\n"
        f" sheetObject.{property_name} = CLng(propertyValue)\n"
        f" {function_name} = CLng(sheetObject.{property_name})\n"
        "End Function"
    )
    try:
        actual = int(
            _evaluate(
                catia_app,
                script,
                function_name,
                [sheet, requested],
            )
        )
        attempts.append({
            "method": (
                "SystemService.Evaluate."
                f"DrawingSheet.{property_name}=value"
            ),
            "succeeded": True,
            "actual": actual,
            "error": None,
        })
    except Exception as exc:
        attempts.append({
            "method": (
                "SystemService.Evaluate."
                f"DrawingSheet.{property_name}=value"
            ),
            "succeeded": False,
            "actual": None,
            "error": _format_com_error(exc),
        })

    readback = _read_sheet_property(
        catia_app,
        sheet,
        property_name,
    )
    verified = bool(
        readback["verified"]
        and readback["value"] == requested
    )
    selected_method = None
    for attempt in reversed(attempts):
        if attempt["succeeded"]:
            selected_method = attempt["method"]
            break

    return {
        "property": property_name,
        "requested": requested,
        "selected_method": selected_method,
        "attempts": attempts,
        "readback": readback,
        "verified": verified,
    }


def _standard_paper_dimensions(
    paper_size_code: Optional[int],
    orientation_code: Optional[int],
) -> Tuple[Optional[float], Optional[float]]:
    paper_name = PAPER_SIZE_NAMES.get(paper_size_code)
    orientation = ORIENTATION_NAMES.get(orientation_code)
    values = STANDARD_PAPER_PORTRAIT_MM.get(paper_name)
    if values is None or orientation is None:
        return None, None

    portrait_width, portrait_height = values
    if orientation == "landscape":
        return (
            max(portrait_width, portrait_height),
            min(portrait_width, portrait_height),
        )
    return (
        min(portrait_width, portrait_height),
        max(portrait_width, portrait_height),
    )


def _read_paper_dimensions_from_api(
    catia_app: Any,
    sheet: Any,
) -> Dict[str, Any]:
    attempts: List[Dict[str, Any]] = []

    try:
        width = float(sheet.GetPaperWidth())
        height = float(sheet.GetPaperHeight())
        if (
            not math.isfinite(width)
            or not math.isfinite(height)
            or width <= 0.0
            or height <= 0.0
        ):
            raise RuntimeError(
                f"Invalid paper dimensions: {width} x {height}."
            )
        attempts.append({
            "method": (
                "DrawingSheet.GetPaperWidth/"
                "GetPaperHeight"
            ),
            "succeeded": True,
            "width_mm": width,
            "height_mm": height,
            "error": None,
        })
        return {
            "width_mm": width,
            "height_mm": height,
            "read_method": attempts[-1]["method"],
            "attempts": attempts,
            "verified": True,
        }
    except Exception as exc:
        attempts.append({
            "method": (
                "DrawingSheet.GetPaperWidth/"
                "GetPaperHeight"
            ),
            "succeeded": False,
            "width_mm": None,
            "height_mm": None,
            "error": _format_com_error(exc),
        })

    script = (
        "Public Function MCP_ReadPaperDimensions(sheetObject)\n"
        " Dim paperWidth\n"
        " Dim paperHeight\n"
        " paperWidth = sheetObject.GetPaperWidth()\n"
        " paperHeight = sheetObject.GetPaperHeight()\n"
        " MCP_ReadPaperDimensions = Array("
        "CDbl(paperWidth), CDbl(paperHeight))\n"
        "End Function"
    )
    try:
        result = list(
            _evaluate(
                catia_app,
                script,
                "MCP_ReadPaperDimensions",
                [sheet],
            )
        )
        if len(result) != 2:
            raise RuntimeError(
                f"CATIA returned {len(result)} paper values."
            )
        width = float(result[0])
        height = float(result[1])
        if (
            not math.isfinite(width)
            or not math.isfinite(height)
            or width <= 0.0
            or height <= 0.0
        ):
            raise RuntimeError(
                f"Invalid paper dimensions: {width} x {height}."
            )
        attempts.append({
            "method": (
                "SystemService.Evaluate."
                "GetPaperWidth/GetPaperHeight"
            ),
            "succeeded": True,
            "width_mm": width,
            "height_mm": height,
            "error": None,
        })
        return {
            "width_mm": width,
            "height_mm": height,
            "read_method": attempts[-1]["method"],
            "attempts": attempts,
            "verified": True,
        }
    except Exception as exc:
        attempts.append({
            "method": (
                "SystemService.Evaluate."
                "GetPaperWidth/GetPaperHeight"
            ),
            "succeeded": False,
            "width_mm": None,
            "height_mm": None,
            "error": _format_com_error(exc),
        })

    return {
        "width_mm": None,
        "height_mm": None,
        "read_method": None,
        "attempts": attempts,
        "verified": False,
    }


def _paper_dimensions(
    catia_app: Any,
    sheet: Any,
) -> Dict[str, Any]:
    """Return paper metadata without allowing COM property errors to escape."""
    paper_size = _read_sheet_property(
        catia_app,
        sheet,
        "PaperSize",
    )
    orientation = _read_sheet_property(
        catia_app,
        sheet,
        "Orientation",
    )
    dimensions = _read_paper_dimensions_from_api(
        catia_app,
        sheet,
    )

    fallback_used = False
    if not dimensions["verified"]:
        width, height = _standard_paper_dimensions(
            paper_size["value"],
            orientation["value"],
        )
        if width is not None and height is not None:
            dimensions = {
                **dimensions,
                "width_mm": width,
                "height_mm": height,
                "read_method": (
                    "standard_dimensions_from_verified_"
                    "PaperSize_and_Orientation"
                ),
                "verified": True,
            }
            fallback_used = True

    paper_code = paper_size["value"]
    orientation_code = orientation["value"]
    warnings = []
    if not paper_size["verified"]:
        warnings.append(
            "PaperSize could not be read through direct COM or "
            "CATIA-side Evaluate."
        )
    if not orientation["verified"]:
        warnings.append(
            "Orientation could not be read through direct COM or "
            "CATIA-side Evaluate."
        )
    if not dimensions["verified"]:
        warnings.append(
            "Paper width and height could not be determined."
        )

    return {
        "paper_size_code": paper_code,
        "paper_size": PAPER_SIZE_NAMES.get(
            paper_code,
            None,
        ),
        "orientation_code": orientation_code,
        "orientation": ORIENTATION_NAMES.get(
            orientation_code,
            None,
        ),
        "width_mm": dimensions["width_mm"],
        "height_mm": dimensions["height_mm"],
        "paper_size_read": paper_size,
        "orientation_read": orientation,
        "dimension_read": dimensions,
        "fallback_used": fallback_used,
        "warnings": warnings,
        "verified": bool(
            paper_size["verified"]
            and orientation["verified"]
            and dimensions["verified"]
        ),
    }


# ---------------------------------------------------------------------------
# Public Tool Functions
# ---------------------------------------------------------------------------

def get_title_block_properties(
    catia_app,
    sheet_index: Optional[int] = None,
) -> Dict[str, Any]:
    """Read title-block properties from DrawingText and DrawingTable cells."""
    doc = _get_drawing_document(catia_app)
    sheet, sheet_info = _resolve_sheet(doc, sheet_index)
    background = _get_background_view(sheet)
    evidence = _read_title_field_evidence(background)

    properties: Dict[str, str] = {}
    field_evidence: Dict[str, Any] = {}
    warnings = list(evidence.get("warnings", []))
    for field_name in TITLE_BLOCK_KEYS:
        matches = evidence["fields"].get(field_name, [])
        properties[field_name] = (
            str(matches[0].get("current_value", "")) if matches else ""
        )
        field_evidence[field_name] = {
            "match_count": len(matches),
            "ambiguous": len(matches) > 1,
            "matches": matches,
        }
        if len(matches) > 1:
            warnings.append(
                f"Multiple template objects matched '{field_name}'; "
                "the first value is returned."
            )

    text_records, text_warnings = _read_text_records(background)
    table_records, table_warnings = _read_table_records(background)
    warnings.extend(text_warnings)
    warnings.extend(table_warnings)
    return _success({
        "operation": "get_title_block_properties",
        "properties": properties,
        "field_evidence": field_evidence,
        "all_background_texts": text_records,
        "all_background_table_cells": table_records,
        "unrecognised_texts": [
            item for item in text_records
            if item["canonical_field"] is None
        ],
        "sheet": sheet_info,
        "background_view": _background_info(background),
        "projection": _detect_projection_method(catia_app, sheet),
        "paper": _paper_dimensions(catia_app, sheet),
        "coordinate_unit": "mm",
        "model_modified": False,
        "document_save_required": False,
        "document_saved": _document_saved(doc),
    }, warnings)



def set_title_block_properties(
    catia_app,
    properties: Dict[str, str],
    sheet_index: Optional[int] = None,
    create_if_missing: bool = False,
    require_unique: bool = True,
    base_x_mm: Optional[float] = None,
    base_y_mm: Optional[float] = None,
    row_spacing_mm: float = 7.0,
    sheet_margin_mm: float = 5.0,
) -> Dict[str, Any]:
    """Update/create title-block fields as one verified transaction.

    Ambiguous fields and out-of-sheet creation coordinates are rejected before
    any mutation.  If a later COM write fails, all earlier writes/creations from
    the same call are rolled back in reverse order.
    """
    requested = _canonicalise_properties(properties)
    doc = _get_drawing_document(catia_app)
    sheet, sheet_info = _resolve_sheet(doc, sheet_index)
    background = _get_background_view(sheet)
    texts = background.Texts
    records, warnings = _read_text_records(background)
    grouped = _group_text_records(records)
    saved_before = _document_saved(doc)

    # v7: CATIA title blocks are frequently DrawingTables rather than
    # independent DrawingText objects. If a requested field is represented by
    # a table cell, use the table-aware atomic implementation for the complete
    # request so text/table updates cannot diverge.
    table_targets, table_warnings = _discover_table_field_targets(background)
    warnings.extend(table_warnings)
    requested_table_fields = {
        item["field"] for item in table_targets
        if item["field"] in requested
    }
    if requested_table_fields:
        table_aware = _set_template_title_fields(
            doc, sheet, requested, require_unique=require_unique
        )
        update_attempts, update_warnings = _update_drawing(doc, sheet)
        warnings.extend(update_warnings)
        data = {
            "operation": "set_title_block_properties",
            "write_strategy": "DrawingText_and_DrawingTable_atomic",
            "requested_properties": requested,
            "table_aware_result": table_aware,
            "sheet": sheet_info,
            "background_view": _background_info(background),
            "paper": _paper_dimensions(catia_app, sheet),
            "update_attempts": update_attempts,
            "model_modified": bool(table_aware["verified"]),
            "document_save_required": bool(table_aware["verified"]),
            "document_saved_before": saved_before,
            "document_saved_after": _document_saved(doc),
            "preflight_verified": table_aware["preflight_verified"],
        }
        if not table_aware["verified"]:
            return _error(
                "One or more title-block table/text fields failed verified "
                "update. Changes were rolled back where possible.",
                data=data,
                warnings=warnings + table_aware.get("warnings", []),
                status="error",
            )
        return _success(
            data, warnings + table_aware.get("warnings", [])
        )

    paper = _paper_dimensions(catia_app, sheet)
    width = float(paper.get("width_mm") or 420.0)
    height = float(paper.get("height_mm") or 297.0)
    margin = _finite_float(sheet_margin_mm, "sheet_margin_mm")
    if margin < 0.0:
        raise ValueError("sheet_margin_mm must be zero or greater.")
    if width <= 2.0 * margin or height <= 2.0 * margin:
        raise ValueError(
            "sheet_margin_mm leaves no usable title-field placement area."
        )

    x_default = (
        width * 0.62
        if base_x_mm is None
        else _finite_float(base_x_mm, "base_x_mm")
    )
    y_default = (
        height * 0.10
        if base_y_mm is None
        else _finite_float(base_y_mm, "base_y_mm")
    )
    spacing = _finite_float(row_spacing_mm, "row_spacing_mm")
    if spacing <= 0:
        raise ValueError("row_spacing_mm must be greater than zero.")

    ambiguous = {
        field: grouped[field]
        for field in requested
        if len(grouped[field]) > 1
    }
    if require_unique and ambiguous:
        return _error(
            "One or more requested title-block fields are ambiguous. "
            "No properties were modified.",
            data={
                "operation": "set_title_block_properties",
                "requested_properties": requested,
                "ambiguous_fields": {
                    field: {
                        "match_count": len(matches),
                        "matches": matches,
                    }
                    for field, matches in ambiguous.items()
                },
                "sheet": sheet_info,
                "background_view": _background_info(background),
                "paper": paper,
                "model_modified": False,
                "document_save_required": False,
                "document_saved_before": saved_before,
                "document_saved_after": _document_saved(doc),
                "preflight_verified": False,
            },
        )

    creation_plan: Dict[str, Dict[str, Any]] = {}
    creation_index = 0
    for field in requested:
        if grouped[field] or not create_if_missing:
            continue
        x = float(x_default)
        y = float(y_default + creation_index * spacing)
        creation_index += 1
        within = bool(
            margin <= x <= width - margin
            and margin <= y <= height - margin
        )
        creation_plan[field] = {
            "x_mm": x,
            "y_mm": y,
            "within_sheet_margin": within,
            "left_clearance_mm": x,
            "right_clearance_mm": width - x,
            "bottom_clearance_mm": y,
            "top_clearance_mm": height - y,
        }

    invalid_positions = {
        field: item
        for field, item in creation_plan.items()
        if not item["within_sheet_margin"]
    }
    if invalid_positions:
        return _error(
            "One or more missing title-block fields would be created outside "
            "the selected sheet margin. No properties were modified.",
            data={
                "operation": "set_title_block_properties",
                "requested_properties": requested,
                "creation_plan": creation_plan,
                "invalid_creation_positions": invalid_positions,
                "sheet_margin_mm": margin,
                "sheet": sheet_info,
                "background_view": _background_info(background),
                "paper": paper,
                "model_modified": False,
                "document_save_required": False,
                "document_saved_before": saved_before,
                "document_saved_after": _document_saved(doc),
                "preflight_verified": False,
            },
        )

    results: Dict[str, Any] = {}
    undo_stack: List[Dict[str, Any]] = []
    failed = False

    for field_name, value in requested.items():
        matches = grouped[field_name]
        if matches:
            match = matches[0]
            obj = texts.Item(match["index"])
            old_text = str(obj.Text)
            label, _, separator = _split_label_value(old_text)
            new_text = f"{label}{separator or ':'} {value}"
            error = None
            try:
                obj.Text = new_text
                actual_text = str(obj.Text)
                verified = actual_text == new_text
            except Exception as exc:
                actual_text = old_text
                verified = False
                error = _format_com_error(exc)

            if verified:
                undo_stack.append({
                    "kind": "updated_text",
                    "field": field_name,
                    "object": obj,
                    "old_text": old_text,
                })
            else:
                failed = True
            results[field_name] = {
                "action": "updated",
                "ok": verified,
                "target_index": match["index"],
                "match_count": len(matches),
                "old_text": old_text,
                "requested_text": new_text,
                "actual_text": actual_text,
                "readback_verified": verified,
                "error": error,
                "model_modified": verified,
            }
            if failed:
                break
            continue

        if not create_if_missing:
            results[field_name] = {
                "action": "skipped_not_found",
                "ok": True,
                "match_count": 0,
                "model_modified": False,
            }
            warnings.append(
                f"'{field_name}' was not found and create_if_missing=false."
            )
            continue

        placement = creation_plan[field_name]
        x = placement["x_mm"]
        y = placement["y_mm"]
        before_count = int(texts.Count)
        requested_text = f"{field_name}: {value}"
        obj = None
        error = None
        try:
            obj = texts.Add(requested_text, x, y)
            after_count = int(texts.Count)
            actual_text = str(obj.Text)
            verified = bool(
                after_count == before_count + 1
                and actual_text == requested_text
            )
        except Exception as exc:
            after_count = int(texts.Count)
            actual_text = ""
            verified = False
            error = _format_com_error(exc)

        if verified and obj is not None:
            undo_stack.append({
                "kind": "created_text",
                "field": field_name,
                "object": obj,
            })
        else:
            failed = True
        results[field_name] = {
            "action": "created",
            "ok": verified,
            "requested_text": requested_text,
            "actual_text": actual_text,
            "x_mm": x,
            "y_mm": y,
            "placement": placement,
            "texts_count_before": before_count,
            "texts_count_after": after_count,
            "readback_verified": verified,
            "error": error,
            "model_modified": verified,
        }
        if failed:
            break

    rollback_results: List[Dict[str, Any]] = []
    if failed:
        for item in reversed(undo_stack):
            if item["kind"] == "updated_text":
                try:
                    item["object"].Text = item["old_text"]
                    verified = str(item["object"].Text) == item["old_text"]
                    error = None
                except Exception as exc:
                    verified = False
                    error = _format_com_error(exc)
                rollback_results.append({
                    "field": item["field"],
                    "kind": item["kind"],
                    "succeeded": verified,
                    "error": error,
                })
            else:
                deletion = _delete_object(doc, item["object"])
                rollback_results.append({
                    "field": item["field"],
                    "kind": item["kind"],
                    "succeeded": bool(deletion["succeeded"]),
                    "error": deletion.get("error"),
                    "deletion": deletion,
                })

    update_attempts, update_warnings = _update_drawing(doc, sheet)
    warnings.extend(update_warnings)
    final_records, final_warnings = _read_text_records(background)
    warnings.extend(final_warnings)
    final_grouped = _group_text_records(final_records)
    final_values = {
        name: final_grouped[name][0]["value"] if final_grouped[name] else ""
        for name in TITLE_BLOCK_KEYS
    }

    rollback_verified = bool(
        failed
        and all(item["succeeded"] for item in rollback_results)
    ) if rollback_results else bool(failed and not undo_stack)
    modified = bool(undo_stack) and not rollback_verified
    data = {
        "operation": "set_title_block_properties",
        "requested_properties": requested,
        "results": results,
        "final_values": final_values,
        "creation_plan": creation_plan,
        "sheet_margin_mm": margin,
        "sheet": sheet_info,
        "background_view": _background_info(background),
        "paper": paper,
        "update_attempts": update_attempts,
        "transaction": {
            "atomic": True,
            "failed": failed,
            "undo_item_count": len(undo_stack),
            "rollback_attempted": bool(failed and undo_stack),
            "rollback_results": rollback_results,
            "rollback_verified": rollback_verified,
        },
        "model_modified": modified if failed else bool(undo_stack),
        "document_save_required": modified if failed else bool(undo_stack),
        "document_saved_before": saved_before,
        "document_saved_after": _document_saved(doc),
        "preflight_verified": True,
    }

    if failed:
        return _error(
            "One or more title-block properties failed verified update or "
            "creation. Earlier changes were rolled back where possible.",
            data=data,
            warnings=warnings,
            status="partial_success" if modified else "error",
        )
    return _success(data, warnings)




_BACKGROUND_COLLECTIONS = (
    "GeometricElements",
    "Components",
    "Texts",
    "Tables",
    "Pictures",
    "Dimensions",
    "Leaders",
    "Arrows",
    "GDTs",
    "Weldings",
    "Threads",
)

_STRUCTURAL_GEOMETRY_NAMES = {
    "absoluteaxis",
    "axis2d",
    "absolute axis",
    "绝对轴",
}


def _object_descriptor(
    collection_name: str,
    index: int,
    obj: Any,
) -> Dict[str, Any]:
    name = str(_safe_attr(obj, "Name", "")).strip()
    python_type = type(obj).__name__
    normalised_name = re.sub(r"\s+", " ", name.lower())
    normalised_type = python_type.lower()
    structural_reason = None

    if collection_name == "GeometricElements":
        if normalised_name in _STRUCTURAL_GEOMETRY_NAMES:
            structural_reason = "CATIA sketch AbsoluteAxis"
        elif "axis2d" in normalised_type:
            structural_reason = "CATIA Axis2D automation object"

    return {
        "collection": collection_name,
        "index": index,
        "name": name,
        "python_type": python_type,
        "is_structural": structural_reason is not None,
        "structural_reason": structural_reason,
    }


def _collection_items_with_descriptors(
    collection_name: str,
    collection: Any,
) -> List[Dict[str, Any]]:
    count = _safe_count(collection)
    if count is None:
        return []
    result: List[Dict[str, Any]] = []
    for index in range(1, count + 1):
        try:
            obj = collection.Item(index)
            result.append({
                "object": obj,
                "descriptor": _object_descriptor(
                    collection_name,
                    index,
                    obj,
                ),
            })
        except Exception:
            continue
    return result


def _content_inventory(
    background_view: Any,
) -> Dict[str, Any]:
    raw_counts: Dict[str, Optional[int]] = {}
    meaningful_counts: Dict[str, int] = {}
    structural_counts: Dict[str, int] = {}
    items: Dict[str, List[Dict[str, Any]]] = {}
    raw_total = 0
    meaningful_total = 0
    structural_total = 0

    for name in _BACKGROUND_COLLECTIONS:
        collection = _safe_attr(background_view, name, None)
        raw_count = _safe_count(collection)
        raw_counts[name] = raw_count
        if raw_count is not None:
            raw_total += raw_count

        entries = _collection_items_with_descriptors(
            name,
            collection,
        )
        descriptors = [
            entry["descriptor"] for entry in entries
        ]
        items[name] = descriptors
        structural_count = sum(
            1 for item in descriptors
            if item["is_structural"]
        )
        meaningful_count = sum(
            1 for item in descriptors
            if not item["is_structural"]
        )
        structural_counts[name] = structural_count
        meaningful_counts[name] = meaningful_count
        structural_total += structural_count
        meaningful_total += meaningful_count

    return {
        **raw_counts,
        "raw_counts": raw_counts,
        "meaningful_counts": meaningful_counts,
        "structural_counts": structural_counts,
        "raw_total_objects": raw_total,
        "structural_total_objects": structural_total,
        # Keep the legacy field name, but define it as user/template content.
        "total_known_objects": meaningful_total,
        "meaningful_total_objects": meaningful_total,
        "has_meaningful_content": meaningful_total > 0,
        "items": items,
        "structural_policy": (
            "exclude AbsoluteAxis/Axis2D from template content"
        ),
    }


def _content_counts(
    background_view: Any,
) -> Dict[str, Any]:
    return _content_inventory(background_view)


def _select_background_contents(
    drawing_doc: Any,
    background_view: Any,
) -> Dict[str, Any]:
    selection = drawing_doc.Selection
    selection.Clear()
    selected_total = 0
    selected_by_collection: Dict[str, int] = {}
    skipped_structural: List[Dict[str, Any]] = []
    errors: List[str] = []

    for name in _BACKGROUND_COLLECTIONS:
        selected = 0
        collection = _safe_attr(
            background_view,
            name,
            None,
        )
        for entry in _collection_items_with_descriptors(
            name,
            collection,
        ):
            descriptor = entry["descriptor"]
            if descriptor["is_structural"]:
                skipped_structural.append(descriptor)
                continue
            try:
                selection.Add(entry["object"])
                selected += 1
                selected_total += 1
            except Exception as exc:
                errors.append(
                    f"{name}[{descriptor['index']}]: "
                    f"{_format_com_error(exc)}"
                )
        selected_by_collection[name] = selected

    return {
        "selected_total": selected_total,
        "selected_by_collection": selected_by_collection,
        "skipped_structural": skipped_structural,
        "selection_errors": errors,
        "selection_count_readback": _safe_count(selection),
    }


def _delete_single_background_object(
    drawing_doc: Any,
    collection: Any,
    obj: Any,
    descriptor: Dict[str, Any],
) -> Dict[str, Any]:
    selection = drawing_doc.Selection
    attempts: List[Dict[str, Any]] = []
    before_count = _safe_count(collection)

    try:
        selection.Clear()
        selection.Add(obj)
        selection.Delete()
        selection.Clear()
        after_selection_count = _safe_count(collection)
        selection_verified = bool(
            before_count is not None
            and after_selection_count is not None
            and after_selection_count < before_count
        )
        attempts.append({
            "method": "document.Selection.Delete_single",
            "succeeded": True,
            "count_before": before_count,
            "count_after": after_selection_count,
            "verified_removed": selection_verified,
            "error": None,
        })
        if selection_verified:
            return {
                "descriptor": descriptor,
                "succeeded": True,
                "selected_strategy": (
                    "document.Selection.Delete_single"
                ),
                "attempts": attempts,
            }
    except Exception as exc:
        attempts.append({
            "method": "document.Selection.Delete_single",
            "succeeded": False,
            "verified_removed": False,
            "error": _format_com_error(exc),
        })
        try:
            selection.Clear()
        except Exception:
            pass

    # Many Drafting collections expose Remove(index/name). This is
    # especially important for DrawingComponents and Pictures.
    for remove_key in (
        descriptor.get("index"),
        descriptor.get("name") or None,
    ):
        if remove_key is None:
            continue
        try:
            current_before = _safe_count(collection)
            collection.Remove(remove_key)
            current_after = _safe_count(collection)
            verified = bool(
                current_before is not None
                and current_after is not None
                and current_after < current_before
            )
            attempts.append({
                "method": f"collection.Remove({remove_key!r})",
                "succeeded": True,
                "count_before": current_before,
                "count_after": current_after,
                "verified_removed": verified,
                "error": None,
            })
            if verified:
                return {
                    "descriptor": descriptor,
                    "succeeded": True,
                    "selected_strategy": (
                        f"collection.Remove({remove_key!r})"
                    ),
                    "attempts": attempts,
                }
        except Exception as exc:
            attempts.append({
                "method": f"collection.Remove({remove_key!r})",
                "succeeded": False,
                "verified_removed": False,
                "error": _format_com_error(exc),
            })

    return {
        "descriptor": descriptor,
        "succeeded": False,
        "selected_strategy": None,
        "attempts": attempts,
    }


def _delete_remaining_background_contents(
    drawing_doc: Any,
    background_view: Any,
) -> Dict[str, Any]:
    results: List[Dict[str, Any]] = []
    # Reverse iteration protects collection indices as objects disappear.
    for name in _BACKGROUND_COLLECTIONS:
        collection = _safe_attr(background_view, name, None)
        entries = _collection_items_with_descriptors(
            name,
            collection,
        )
        for entry in reversed(entries):
            descriptor = entry["descriptor"]
            if descriptor["is_structural"]:
                continue
            results.append(
                _delete_single_background_object(
                    drawing_doc,
                    collection,
                    entry["object"],
                    descriptor,
                )
            )
    return {
        "attempted": bool(results),
        "results": results,
        "all_calls_succeeded": all(
            result["succeeded"] for result in results
        ) if results else True,
    }


def _clear_background_contents(
    drawing_doc: Any,
    background_view: Any,
) -> Dict[str, Any]:
    before = _content_inventory(background_view)
    selection_result = _select_background_contents(
        drawing_doc,
        background_view,
    )
    if selection_result["selected_total"] == 0:
        verified = not before["has_meaningful_content"]
        return {
            "attempted": False,
            "succeeded": verified,
            "reason": (
                "No meaningful user/template content. "
                "CATIA structural axis is intentionally preserved."
                if verified else
                "Meaningful content exists but could not be selected."
            ),
            "before": before,
            "after": before,
            "selection": selection_result,
            "individual_fallback": {
                "attempted": False,
                "results": [],
            },
        }

    selection = drawing_doc.Selection
    error = None
    try:
        selection.Delete()
        selection.Clear()
        call_succeeded = True
    except Exception as exc:
        call_succeeded = False
        error = _format_com_error(exc)
        try:
            selection.Clear()
        except Exception:
            pass

    after_batch = _content_inventory(background_view)
    individual_fallback = {
        "attempted": False,
        "results": [],
        "all_calls_succeeded": True,
    }
    if after_batch["has_meaningful_content"]:
        individual_fallback = _delete_remaining_background_contents(
            drawing_doc,
            background_view,
        )

    after = _content_inventory(background_view)
    verified = not after["has_meaningful_content"]
    return {
        "attempted": True,
        "succeeded": verified,
        "delete_call_succeeded": call_succeeded,
        "error": error,
        "before": before,
        "after_batch_delete": after_batch,
        "after": after,
        "selection": selection_result,
        "individual_fallback": individual_fallback,
        "verification_basis": (
            "meaningful user/template objects removed; "
            "AbsoluteAxis/Axis2D preserved"
        ),
    }


def _copy_background_contents(
    template_doc: Any,
    template_background: Any,
) -> Dict[str, Any]:
    inventory = _content_inventory(template_background)
    if not inventory["has_meaningful_content"]:
        raise ToolOperationError(
            "Template Background View contains no meaningful "
            "user/template content. CATIA's structural AbsoluteAxis "
            "does not count as a drawing frame.",
            data={
                "template_preflight": {
                    "verified_nonempty": False,
                    "inventory": inventory,
                },
                "model_modified": False,
                "document_save_required": False,
            },
        )

    selection_result = _select_background_contents(
        template_doc,
        template_background,
    )
    if selection_result["selected_total"] == 0:
        raise ToolOperationError(
            "Template contains meaningful content, but no template "
            "objects could be selected for Copy.",
            data={
                "template_preflight": {
                    "verified_nonempty": True,
                    "inventory": inventory,
                    "selection": selection_result,
                },
                "model_modified": False,
                "document_save_required": False,
            },
        )

    selection = template_doc.Selection
    try:
        selection.Copy()
        succeeded = True
        error = None
    except Exception as exc:
        succeeded = False
        error = _format_com_error(exc)
    finally:
        try:
            selection.Clear()
        except Exception:
            pass

    if not succeeded:
        raise ToolOperationError(
            f"Template Copy failed: {error}",
            data={
                "template_preflight": {
                    "verified_nonempty": True,
                    "inventory": inventory,
                    "selection": selection_result,
                },
                "model_modified": False,
                "document_save_required": False,
            },
        )
    return {
        "succeeded": True,
        "template_preflight": {
            "verified_nonempty": True,
            "inventory": inventory,
        },
        "selection": selection_result,
    }


def _paste_background_contents(
    target_doc: Any,
    target_background: Any,
) -> Dict[str, Any]:
    selection = target_doc.Selection
    selection.Clear()
    selection.Add(target_background)
    try:
        selection.Paste()
        succeeded = True
        error = None
    except Exception as exc:
        succeeded = False
        error = _format_com_error(exc)
    finally:
        try:
            selection.Clear()
        except Exception:
            pass
    return {
        "succeeded": succeeded,
        "attempts": [{
            "method": "target_document.Selection.Paste",
            "succeeded": succeeded,
            "error": error,
        }],
    }


def _activate_document(document: Any) -> Dict[str, Any]:
    try:
        document.Activate()
        return {
            "attempted": True,
            "succeeded": True,
            "error": None,
        }
    except Exception as exc:
        return {
            "attempted": True,
            "succeeded": False,
            "error": _format_com_error(exc),
        }


def _activate_view(view: Any) -> Dict[str, Any]:
    try:
        view.Activate()
        return {
            "attempted": True,
            "succeeded": True,
            "error": None,
        }
    except Exception as exc:
        return {
            "attempted": True,
            "succeeded": False,
            "error": _format_com_error(exc),
        }



def _close_document(
    document: Any,
    *,
    discard_unsaved_changes: bool = False,
) -> Dict[str, Any]:
    """Close a document without allowing a modal save prompt to block MCP.

    Template documents opened by this tool are expected to remain unchanged, so
    normal cleanup uses ``discard_unsaved_changes=False``.  The optional alert
    suppression is available for rollback/temporary documents and always
    restores the caller's previous CATIA setting.
    """
    details: Dict[str, Any] = {
        "attempted": document is not None,
        "succeeded": None,
        "error": None,
        "discard_unsaved_changes": bool(discard_unsaved_changes),
        "file_alerts_temporarily_disabled": False,
        "file_alerts_disable_error": None,
        "file_alerts_restore_error": None,
    }
    if document is None:
        return details

    application = _safe_attr(document, "Application", None)
    previous_alerts: Optional[bool] = None
    try:
        if discard_unsaved_changes and application is not None:
            try:
                previous_alerts = bool(application.DisplayFileAlerts)
                application.DisplayFileAlerts = False
                details["file_alerts_temporarily_disabled"] = True
            except Exception as exc:
                details["file_alerts_disable_error"] = _format_com_error(exc)
        document.Close()
        details["succeeded"] = True
    except Exception as exc:
        details["succeeded"] = False
        details["error"] = _format_com_error(exc)
    finally:
        if previous_alerts is not None:
            try:
                application.DisplayFileAlerts = previous_alerts
            except Exception as exc:
                details["file_alerts_restore_error"] = _format_com_error(exc)
    return details



def _find_open_document(
    catia_app: Any,
    path: str,
) -> Any:
    target = _normalised_path_key(path)
    documents = catia_app.Documents
    for index in range(1, int(documents.Count) + 1):
        document = documents.Item(index)
        full_name = str(
            _safe_attr(document, "FullName", "")
        )
        if (
            full_name
            and _normalised_path_key(full_name) == target
        ):
            return document
    return None


def _sheet_scale(sheet: Any) -> Optional[float]:
    for name in ("Scale", "Scale2"):
        try:
            return float(getattr(sheet, name))
        except Exception:
            continue
    return None


def _apply_template_settings(
    catia_app: Any,
    target_sheet: Any,
    template_sheet: Any,
) -> Dict[str, Any]:
    requested_reads = {
        name: _read_sheet_property(
            catia_app,
            template_sheet,
            name,
        )
        for name in (
            "PaperSize",
            "Orientation",
            "ProjectionMethod",
        )
    }
    before_reads = {
        name: _read_sheet_property(
            catia_app,
            target_sheet,
            name,
        )
        for name in (
            "PaperSize",
            "Orientation",
            "ProjectionMethod",
        )
    }
    requested_scale = _sheet_scale(template_sheet)
    before_scale = _sheet_scale(target_sheet)

    missing_requested = [
        name
        for name, result in requested_reads.items()
        if not result["verified"]
    ]
    if missing_requested:
        return {
            "requested": {
                **{
                    name: result["value"]
                    for name, result in requested_reads.items()
                },
                "Scale": requested_scale,
            },
            "before": {
                **{
                    name: result["value"]
                    for name, result in before_reads.items()
                },
                "Scale": before_scale,
            },
            "requested_reads": requested_reads,
            "before_reads": before_reads,
            "after_reads": {},
            "attempts": [],
            "scale_set_method": None,
            "verified": False,
            "model_modified": False,
            "error": (
                "Template sheet settings could not be read: "
                + ", ".join(missing_requested)
            ),
        }

    attempts: List[Dict[str, Any]] = []
    property_results: Dict[str, Any] = {}
    for name in (
        "PaperSize",
        "Orientation",
        "ProjectionMethod",
    ):
        result = _set_sheet_property(
            catia_app,
            target_sheet,
            name,
            requested_reads[name]["value"],
        )
        property_results[name] = result
        attempts.extend(result["attempts"])

    scale_method = None
    scale_verified = requested_scale is None
    scale_attempts: List[Dict[str, Any]] = []
    if requested_scale is not None:
        for name in ("Scale", "Scale2"):
            try:
                setattr(
                    target_sheet,
                    name,
                    requested_scale,
                )
                actual_scale = _sheet_scale(target_sheet)
                verified = bool(
                    actual_scale is not None
                    and abs(actual_scale - requested_scale)
                    <= 1e-9
                )
                scale_attempts.append({
                    "property": name,
                    "succeeded": True,
                    "actual": actual_scale,
                    "verified": verified,
                    "error": None,
                })
                if verified:
                    scale_method = name
                    scale_verified = True
                    break
            except Exception as exc:
                scale_attempts.append({
                    "property": name,
                    "succeeded": False,
                    "actual": None,
                    "verified": False,
                    "error": _format_com_error(exc),
                })
    attempts.extend(scale_attempts)

    after_reads = {
        name: _read_sheet_property(
            catia_app,
            target_sheet,
            name,
        )
        for name in (
            "PaperSize",
            "Orientation",
            "ProjectionMethod",
        )
    }
    after_scale = _sheet_scale(target_sheet)

    property_verified = all(
        property_results[name]["verified"]
        and after_reads[name]["verified"]
        and after_reads[name]["value"]
        == requested_reads[name]["value"]
        for name in property_results
    )
    verified = bool(property_verified and scale_verified)

    changed = any(
        before_reads[name]["verified"]
        and after_reads[name]["verified"]
        and before_reads[name]["value"]
        != after_reads[name]["value"]
        for name in property_results
    )
    if (
        before_scale is not None
        and after_scale is not None
        and abs(after_scale - before_scale) > 1e-9
    ):
        changed = True

    return {
        "requested": {
            **{
                name: result["value"]
                for name, result in requested_reads.items()
            },
            "Scale": requested_scale,
        },
        "before": {
            **{
                name: result["value"]
                for name, result in before_reads.items()
            },
            "Scale": before_scale,
        },
        "after": {
            **{
                name: result["value"]
                for name, result in after_reads.items()
            },
            "Scale": after_scale,
        },
        "requested_reads": requested_reads,
        "before_reads": before_reads,
        "after_reads": after_reads,
        "property_results": property_results,
        "attempts": attempts,
        "scale_set_method": scale_method,
        "scale_verified": scale_verified,
        "verified": verified,
        "model_modified": changed,
        "error": None if verified else (
            "One or more sheet settings failed verified readback."
        ),
    }



def _capture_meaningful_background_objects(
    background_view: Any,
) -> Tuple[Dict[str, List[Dict[str, Any]]], Dict[str, Any]]:
    captured: Dict[str, List[Dict[str, Any]]] = {}
    descriptors: Dict[str, List[Dict[str, Any]]] = {}
    total = 0
    for name in _BACKGROUND_COLLECTIONS:
        collection = _safe_attr(background_view, name, None)
        entries = [
            entry
            for entry in _collection_items_with_descriptors(name, collection)
            if not entry["descriptor"]["is_structural"]
        ]
        captured[name] = entries
        descriptors[name] = [entry["descriptor"] for entry in entries]
        total += len(entries)
    return captured, {
        "captured_total": total,
        "captured_by_collection": {
            name: len(entries) for name, entries in captured.items()
        },
        "descriptors": descriptors,
    }


def _delete_captured_background_objects(
    drawing_doc: Any,
    background_view: Any,
    captured: Dict[str, List[Dict[str, Any]]],
) -> Dict[str, Any]:
    before = _content_inventory(background_view)
    results: List[Dict[str, Any]] = []
    for name in _BACKGROUND_COLLECTIONS:
        collection = _safe_attr(background_view, name, None)
        for entry in reversed(captured.get(name, [])):
            results.append(
                _delete_single_background_object(
                    drawing_doc,
                    collection,
                    entry["object"],
                    entry["descriptor"],
                )
            )
    after = _content_inventory(background_view)
    failed = [item for item in results if not item["succeeded"]]
    return {
        "attempted": bool(results),
        "succeeded": not failed,
        "strategy": "delete_preexisting_objects_after_verified_paste",
        "results": results,
        "failed_count": len(failed),
        "before": before,
        "after": after,
        "data_loss_policy": (
            "template is pasted and verified before any preexisting target "
            "object is deleted"
        ),
    }


def _paste_delta_verification(
    copy_result: Dict[str, Any],
    before: Dict[str, Any],
    after: Dict[str, Any],
) -> Dict[str, Any]:
    expected_by = dict(
        copy_result.get("selection", {}).get("selected_by_collection", {})
    )
    expected_total = int(
        copy_result.get("selection", {}).get("selected_total", 0) or 0
    )
    before_by = before.get("meaningful_counts", {})
    after_by = after.get("meaningful_counts", {})
    delta_by = {
        name: int(after_by.get(name, 0)) - int(before_by.get(name, 0))
        for name in _BACKGROUND_COLLECTIONS
    }
    total_delta = int(after.get("meaningful_total_objects", 0)) - int(
        before.get("meaningful_total_objects", 0)
    )
    collection_results = {
        name: {
            "expected_minimum": int(expected_by.get(name, 0) or 0),
            "actual_delta": int(delta_by.get(name, 0)),
            "verified": int(delta_by.get(name, 0)) >= int(expected_by.get(name, 0) or 0),
        }
        for name in _BACKGROUND_COLLECTIONS
        if int(expected_by.get(name, 0) or 0) > 0
    }
    verified = bool(
        expected_total > 0
        and total_delta >= expected_total
        and all(item["verified"] for item in collection_results.values())
    )
    return {
        "verified": verified,
        "expected_total": expected_total,
        "actual_total_delta": total_delta,
        "expected_by_collection": expected_by,
        "actual_delta_by_collection": delta_by,
        "collection_results": collection_results,
        "verification_policy": (
            "every copied collection must gain at least its selected object "
            "count; total delta must cover all copied objects"
        ),
    }


def _rollback_added_background_objects(
    drawing_doc: Any,
    background_view: Any,
    baseline: Dict[str, Any],
) -> Dict[str, Any]:
    results: List[Dict[str, Any]] = []
    baseline_raw = baseline.get("raw_counts", {})
    for name in _BACKGROUND_COLLECTIONS:
        collection = _safe_attr(background_view, name, None)
        current_count = _safe_count(collection)
        keep_count = int(baseline_raw.get(name) or 0)
        if current_count is None:
            continue
        for index in range(current_count, keep_count, -1):
            try:
                obj = collection.Item(index)
                descriptor = _object_descriptor(name, index, obj)
                if descriptor["is_structural"]:
                    continue
                results.append(
                    _delete_single_background_object(
                        drawing_doc,
                        collection,
                        obj,
                        descriptor,
                    )
                )
            except Exception as exc:
                results.append({
                    "descriptor": {"collection": name, "index": index},
                    "succeeded": False,
                    "selected_strategy": None,
                    "attempts": [],
                    "error": _format_com_error(exc),
                })
    after = _content_inventory(background_view)
    restored = all(
        int(after.get("meaningful_counts", {}).get(name, 0))
        == int(baseline.get("meaningful_counts", {}).get(name, 0))
        for name in _BACKGROUND_COLLECTIONS
    )
    return {
        "attempted": bool(results),
        "results": results,
        "rollback_verified": restored,
        "after": after,
    }


def _restore_sheet_settings(
    catia_app: Any,
    target_sheet: Any,
    before: Dict[str, Any],
) -> Dict[str, Any]:
    attempts: List[Dict[str, Any]] = []
    verified = True
    for name in ("PaperSize", "Orientation", "ProjectionMethod"):
        value = before.get(name)
        if value is None:
            continue
        result = _set_sheet_property(catia_app, target_sheet, name, int(value))
        attempts.append({"property": name, "result": result})
        verified = verified and bool(result["verified"])
    requested_scale = before.get("Scale")
    scale_result = {
        "attempted": requested_scale is not None,
        "verified": requested_scale is None,
        "method": None,
        "actual": _sheet_scale(target_sheet),
        "errors": [],
    }
    if requested_scale is not None:
        for name in ("Scale", "Scale2"):
            try:
                setattr(target_sheet, name, float(requested_scale))
                actual = _sheet_scale(target_sheet)
                ok = actual is not None and abs(actual - float(requested_scale)) <= 1e-9
                if ok:
                    scale_result.update({
                        "verified": True,
                        "method": name,
                        "actual": actual,
                    })
                    break
            except Exception as exc:
                scale_result["errors"].append(_format_com_error(exc))
        verified = verified and bool(scale_result["verified"])
    return {
        "attempted": True,
        "verified": verified,
        "property_attempts": attempts,
        "scale": scale_result,
    }



def load_drawing_frame(
    catia_app,
    template_path: str,
    replace_existing: bool = True,
    sheet_index: Optional[int] = None,
    apply_sheet_settings: bool = False,
) -> Dict[str, Any]:
    """Copy a verified template frame with paste-first replacement safety."""
    raw_path = str(template_path).strip()
    if not raw_path:
        raise ValueError("template_path cannot be empty.")
    path = os.path.abspath(os.path.expanduser(raw_path))
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Template file not found: {path}")
    if not path.lower().endswith(".catdrawing"):
        raise ValueError("template_path must point to a .CATDrawing file.")

    target_doc = _get_drawing_document(catia_app)
    target_full_name = str(_safe_attr(target_doc, "FullName", ""))
    if (
        target_full_name
        and _normalised_path_key(target_full_name) == _normalised_path_key(path)
    ):
        raise ValueError("Template and active target CATDrawing must differ.")

    target_sheet, target_sheet_info = _resolve_sheet(target_doc, sheet_index)
    target_background = _get_background_view(target_sheet)
    target_counts_before = _content_inventory(target_background)
    target_capture, target_capture_summary = _capture_meaningful_background_objects(
        target_background
    )
    saved_before = _document_saved(target_doc)

    template_doc = _find_open_document(catia_app, path)
    opened_by_tool = template_doc is None
    if template_doc is None:
        template_doc = catia_app.Documents.Open(path)

    lifecycle = {
        "opened_by_tool": opened_by_tool,
        "policy": "close_if_opened_by_tool_preserve_if_preexisting",
        "cleanup": None,
        "target_restore": None,
    }
    warnings: List[str] = []

    try:
        template_activation = _activate_document(template_doc)
        template_sheet, template_sheet_info = _resolve_sheet(template_doc, 1)
        template_background = _get_background_view(template_sheet)
        template_view_activation = _activate_view(template_background)
        template_counts = _content_inventory(template_background)
        template_preflight = {
            "verified_nonempty": bool(template_counts["has_meaningful_content"]),
            "inventory": template_counts,
            "performed_before_target_mutation": True,
        }
        if not template_preflight["verified_nonempty"]:
            raise ToolOperationError(
                "Template Background View contains no meaningful user/template "
                "content. Target CATDrawing was not modified.",
                data={
                    "template_preflight": template_preflight,
                    "template_lifecycle": lifecycle,
                    "model_modified": False,
                    "document_save_required": False,
                },
            )

        template_projection = _detect_projection_method(catia_app, template_sheet)
        template_paper = _paper_dimensions(catia_app, template_sheet)
        copy_result = _copy_background_contents(template_doc, template_background)

        target_activation = _activate_document(target_doc)
        target_view_activation = _activate_view(target_background)
        counts_before_paste = _content_inventory(target_background)
        paste_result = _paste_background_contents(target_doc, target_background)
        if not paste_result["succeeded"]:
            counts_after_failed_paste = _content_inventory(target_background)
            partial_delta = (
                int(counts_after_failed_paste["meaningful_total_objects"])
                - int(counts_before_paste["meaningful_total_objects"])
            )
            paste_rollback = (
                _rollback_added_background_objects(
                    target_doc, target_background, counts_before_paste
                )
                if partial_delta > 0
                else {
                    "attempted": False,
                    "results": [],
                    "rollback_verified": True,
                    "after": counts_after_failed_paste,
                }
            )
            restored = bool(paste_rollback["rollback_verified"])
            raise ToolOperationError(
                "Template Copy succeeded, but Paste into target Background View "
                "failed before any existing target content was deleted.",
                data={
                    "copy_result": copy_result,
                    "paste_result": paste_result,
                    "counts_before_paste": counts_before_paste,
                    "counts_after_failed_paste": counts_after_failed_paste,
                    "partial_added_count": partial_delta,
                    "paste_rollback": paste_rollback,
                    "target_snapshot": target_capture_summary,
                    "template_lifecycle": lifecycle,
                    "model_modified": not restored,
                    "document_save_required": not restored,
                    "replacement_policy": "paste_then_delete_preexisting",
                },
                status="error" if restored else "partial_success",
            )

        paste_update_attempts, paste_update_warnings = _update_drawing(
            target_doc, target_sheet
        )
        warnings.extend(paste_update_warnings)
        counts_after_paste = _content_inventory(target_background)
        paste_verification = _paste_delta_verification(
            copy_result, counts_before_paste, counts_after_paste
        )
        if not paste_verification["verified"]:
            rollback = _rollback_added_background_objects(
                target_doc, target_background, counts_before_paste
            )
            raise ToolOperationError(
                "Paste returned without error, but the complete copied-content "
                "delta could not be verified.",
                data={
                    "template_counts": template_counts,
                    "target_counts_before": target_counts_before,
                    "counts_before_paste": counts_before_paste,
                    "counts_after_paste": counts_after_paste,
                    "paste_verification": paste_verification,
                    "copy_result": copy_result,
                    "paste_result": paste_result,
                    "paste_rollback": rollback,
                    "template_lifecycle": lifecycle,
                    "model_modified": not rollback["rollback_verified"],
                    "document_save_required": not rollback["rollback_verified"],
                    "replacement_policy": "paste_then_delete_preexisting",
                },
                status=(
                    "error" if rollback["rollback_verified"] else "partial_success"
                ),
            )

        settings_result = None
        settings_rollback = None
        if apply_sheet_settings:
            settings_result = _apply_template_settings(
                catia_app, target_sheet, template_sheet
            )
            if not settings_result["verified"]:
                paste_rollback = _rollback_added_background_objects(
                    target_doc, target_background, counts_before_paste
                )
                settings_rollback = _restore_sheet_settings(
                    catia_app, target_sheet, settings_result.get("before", {})
                )
                restored = bool(
                    paste_rollback["rollback_verified"]
                    and settings_rollback["verified"]
                )
                raise ToolOperationError(
                    "Template sheet settings failed readback verification. "
                    "Pasted content and changed settings were rolled back where "
                    "possible; existing target frame content was never deleted.",
                    data={
                        "sheet_settings": settings_result,
                        "sheet_settings_rollback": settings_rollback,
                        "paste_rollback": paste_rollback,
                        "template_lifecycle": lifecycle,
                        "model_modified": not restored,
                        "document_save_required": not restored,
                        "replacement_policy": "paste_then_delete_preexisting",
                    },
                    status="error" if restored else "partial_success",
                )

        clear_result = {
            "attempted": False,
            "succeeded": True,
            "reason": "replace_existing=false",
            "strategy": "none",
        }
        if replace_existing:
            clear_result = _delete_captured_background_objects(
                target_doc, target_background, target_capture
            )
            if not clear_result["succeeded"]:
                raise ToolOperationError(
                    "Template content was pasted and verified, but one or more "
                    "preexisting target objects could not be deleted. Existing "
                    "content was preserved rather than risking data loss.",
                    data={
                        "clear_result": clear_result,
                        "copy_result": copy_result,
                        "paste_result": paste_result,
                        "paste_verification": paste_verification,
                        "sheet_settings": settings_result,
                        "template_lifecycle": lifecycle,
                        "model_modified": True,
                        "document_save_required": True,
                        "replacement_policy": "paste_then_delete_preexisting",
                    },
                    status="partial_success",
                )

        final_update_attempts, final_update_warnings = _update_drawing(
            target_doc, target_sheet
        )
        warnings.extend(final_update_warnings)
        target_counts_after = _content_inventory(target_background)

        return _success({
            "operation": "load_drawing_frame",
            "template": {
                "path": path,
                "document_name": str(_safe_attr(template_doc, "Name", "")),
                "opened_by_tool": opened_by_tool,
                "activation": template_activation,
                "sheet": template_sheet_info,
                "background_view": _background_info(template_background),
                "background_activation": template_view_activation,
                "content_counts": template_counts,
                "paper": template_paper,
                "projection": template_projection,
                "has_title_block_text": bool(int(template_counts.get("Texts") or 0) > 0),
            },
            "target": {
                "document_name": str(_safe_attr(target_doc, "Name", "")),
                "document_full_name": target_full_name,
                "activation": target_activation,
                "sheet": target_sheet_info,
                "background_view": _background_info(target_background),
                "background_activation": target_view_activation,
                "counts_before_operation": target_counts_before,
                "counts_before_paste": counts_before_paste,
                "counts_after_paste": counts_after_paste,
                "counts_after": target_counts_after,
                "paper": _paper_dimensions(catia_app, target_sheet),
                "projection": _detect_projection_method(catia_app, target_sheet),
            },
            "replace_existing": bool(replace_existing),
            "apply_sheet_settings": bool(apply_sheet_settings),
            "sheet_settings": settings_result,
            "sheet_settings_rollback": settings_rollback,
            "template_preflight": template_preflight,
            "target_snapshot": target_capture_summary,
            "copy_result": copy_result,
            "paste_result": paste_result,
            "paste_verification": paste_verification,
            "paste_verified": True,
            "clear_result": clear_result,
            "replacement_policy": "paste_then_delete_preexisting",
            "existing_content_deleted_only_after_verified_paste": True,
            "paste_update_attempts": paste_update_attempts,
            "final_update_attempts": final_update_attempts,
            "template_lifecycle": lifecycle,
            "model_modified": True,
            "document_save_required": True,
            "document_saved_before": saved_before,
            "document_saved_after": _document_saved(target_doc),
        }, warnings)
    finally:
        if opened_by_tool:
            lifecycle["cleanup"] = _close_document(
                template_doc, discard_unsaved_changes=False
            )
        else:
            lifecycle["cleanup"] = {
                "attempted": False,
                "succeeded": True,
                "reason": "preexisting user template document preserved",
            }
        lifecycle["target_restore"] = _activate_document(target_doc)



def detect_projection_standard(
    catia_app,
    sheet_index: Optional[int] = None,
) -> Dict[str, Any]:
    """Detect projection standard for the selected sheet."""
    doc = _get_drawing_document(catia_app)
    sheet, sheet_info = _resolve_sheet(
        doc,
        sheet_index,
    )
    projection = _detect_projection_method(
        catia_app,
        sheet,
    )
    warnings: List[str] = []
    if not projection["verified"]:
        warnings.append(
            "Projection could not be classified as "
            "first-angle or third-angle."
        )
    return _success({
        "operation": "detect_projection_standard",
        "projection": projection,
        "sheet": sheet_info,
        "paper": _paper_dimensions(catia_app, sheet),
        "coordinate_unit": "mm",
        "model_modified": False,
        "document_save_required": False,
        "document_saved": _document_saved(doc),
    }, warnings)



# ---------------------------------------------------------------------------
# Managed CATDrawing template resources (v7)
# ---------------------------------------------------------------------------


def _sha256_file(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _default_template_path() -> str:
    return str(
        Path(__file__).resolve().parents[1]
        / "resources"
        / "drawing_templates"
        / DEFAULT_TEMPLATE_FILENAME
    )


def _resolve_template_resource(template_path: Optional[str]) -> str:
    explicit = str(template_path or "").strip()
    environment = str(os.getenv(DEFAULT_TEMPLATE_ENV, "")).strip()
    selected = explicit or environment or _default_template_path()
    path = os.path.abspath(os.path.expanduser(selected))
    if not os.path.isfile(path):
        raise FileNotFoundError(
            "CATDrawing template was not found: "
            f"{path}. Supply template_path, set {DEFAULT_TEMPLATE_ENV}, "
            "or deploy the default resource."
        )
    if not path.lower().endswith(".catdrawing"):
        raise ValueError("The selected template must be a .CATDrawing file.")
    return path


def _normalised_output_drawing_path(value: str) -> str:
    raw = str(value).strip()
    if not raw:
        raise ValueError("output_path cannot be empty.")
    path = os.path.abspath(os.path.expanduser(raw))
    if not path.lower().endswith(".catdrawing"):
        raise ValueError("output_path must end with .CATDrawing.")
    return path


def _sidecar_path(target_path: str, label: str) -> str:
    target = Path(target_path)
    return str(
        target.with_name(
            f".{target.name}.mcp_{label}_{uuid.uuid4().hex}"
        )
    )


def _begin_template_file_transaction(
    source_path: str,
    target_path: str,
    overwrite: bool,
) -> Dict[str, Any]:
    if not isinstance(overwrite, bool):
        raise ValueError("overwrite must be a boolean.")
    if _normalised_path_key(source_path) == _normalised_path_key(target_path):
        raise ValueError("Template source and output_path must differ.")

    Path(target_path).parent.mkdir(parents=True, exist_ok=True)
    existed_before = os.path.exists(target_path)
    if existed_before and not os.path.isfile(target_path):
        raise IsADirectoryError(f"Output target is not a file: {target_path}")
    if existed_before and not overwrite:
        raise FileExistsError(
            f"Output CATDrawing already exists: {target_path}. "
            "Set overwrite=true to replace it."
        )

    source_sha = _sha256_file(source_path)
    backup_path = _sidecar_path(target_path, "template_backup") if existed_before else None
    staging_path = _sidecar_path(target_path, "template_stage")
    transaction: Dict[str, Any] = {
        "source_path": source_path,
        "source_sha256": source_sha,
        "target_path": target_path,
        "existed_before": existed_before,
        "overwrite": overwrite,
        "backup_path": backup_path,
        "backup_created": False,
        "staging_path": staging_path,
        "staging_created": False,
        "copy_sha256": None,
        "copy_verified": False,
        "target_materialised": False,
        "committed": False,
        "rollback_attempted": False,
        "rollback_succeeded": None,
        "restored_sha256": None,
        "backup_exists_after_completion": None,
        "staging_exists_after_completion": None,
    }

    try:
        if existed_before:
            os.replace(target_path, backup_path)
            transaction["backup_created"] = True
        shutil.copy2(source_path, staging_path)
        transaction["staging_created"] = True
        copy_sha = _sha256_file(staging_path)
        transaction["copy_sha256"] = copy_sha
        transaction["copy_verified"] = copy_sha == source_sha
        if not transaction["copy_verified"]:
            raise RuntimeError("Template staging SHA256 does not match source.")
        os.replace(staging_path, target_path)
        transaction["target_materialised"] = True
        transaction["staging_exists_after_materialisation"] = os.path.exists(staging_path)
        return transaction
    except Exception:
        try:
            if os.path.isfile(staging_path):
                os.unlink(staging_path)
        except OSError:
            pass
        try:
            if os.path.isfile(target_path):
                os.unlink(target_path)
        except OSError:
            pass
        if backup_path and os.path.isfile(backup_path):
            os.replace(backup_path, target_path)
        raise


def _commit_template_file_transaction(transaction: Dict[str, Any]) -> None:
    backup = transaction.get("backup_path")
    if backup and os.path.isfile(backup):
        os.unlink(backup)
    staging = transaction.get("staging_path")
    if staging and os.path.isfile(staging):
        os.unlink(staging)
    transaction["committed"] = True
    transaction["target_sha256"] = _sha256_file(transaction["target_path"])
    transaction["backup_exists_after_completion"] = bool(
        backup and os.path.exists(backup)
    )
    transaction["staging_exists_after_completion"] = bool(
        staging and os.path.exists(staging)
    )


def _rollback_template_file_transaction(transaction: Dict[str, Any]) -> None:
    transaction["rollback_attempted"] = True
    target = transaction["target_path"]
    backup = transaction.get("backup_path")
    staging = transaction.get("staging_path")
    errors: List[str] = []
    try:
        if os.path.isfile(target):
            os.unlink(target)
    except OSError as exc:
        errors.append(str(exc))
    try:
        if staging and os.path.isfile(staging):
            os.unlink(staging)
    except OSError as exc:
        errors.append(str(exc))
    try:
        if backup and os.path.isfile(backup):
            os.replace(backup, target)
    except OSError as exc:
        errors.append(str(exc))

    restored = bool(
        (transaction["existed_before"] and os.path.isfile(target))
        or (not transaction["existed_before"] and not os.path.exists(target))
    )
    if transaction["existed_before"] and restored:
        transaction["restored_sha256"] = _sha256_file(target)
    transaction["rollback_errors"] = errors
    transaction["rollback_succeeded"] = bool(restored and not errors)
    transaction["backup_exists_after_completion"] = bool(
        backup and os.path.exists(backup)
    )
    transaction["staging_exists_after_completion"] = bool(
        staging and os.path.exists(staging)
    )


def _view_inventory(sheet: Any) -> Dict[str, Any]:
    views = sheet.Views
    count = int(views.Count)
    items: List[Dict[str, Any]] = []
    non_system = 0
    generative = 0
    link_count = 0
    for index in range(1, count + 1):
        try:
            view = views.Item(index)
            is_system = index <= 2
            view_type = _safe_attr(view, "ViewType", None)
            is_generative_raw = _safe_attr(view, "IsGenerative", None)
            is_generative = (
                bool(is_generative_raw)
                if is_generative_raw is not None else None
            )
            links = _safe_count(_safe_attr(view, "GenerativeLinks", None))
            if links is None:
                links = _safe_count(_safe_attr(view, "Links", None))
            if not is_system:
                non_system += 1
            if is_generative:
                generative += 1
            link_count += int(links or 0)
            items.append({
                "index": index,
                "name": str(_safe_attr(view, "Name", "")),
                "view_type_code": int(view_type) if view_type is not None else None,
                "is_system_view": is_system,
                "is_generative": is_generative,
                "generative_link_count": links,
                "dimensions_count": _safe_count(_safe_attr(view, "Dimensions", None)),
                "texts_count": _safe_count(_safe_attr(view, "Texts", None)),
            })
        except Exception as exc:
            items.append({
                "index": index,
                "read_error": _format_com_error(exc),
                "is_system_view": index <= 2,
            })
    return {
        "count": count,
        "items": items,
        "non_system_view_count": non_system,
        "generative_view_count": generative,
        "generative_link_count": link_count,
    }


def _read_table_records(background_view: Any) -> Tuple[List[Dict[str, Any]], List[str]]:
    records: List[Dict[str, Any]] = []
    warnings: List[str] = []
    tables = _safe_attr(background_view, "Tables", None)
    count = _safe_count(tables) or 0
    for table_index in range(1, count + 1):
        try:
            table = tables.Item(table_index)
            rows = int(_safe_attr(table, "NumberOfRows", 0) or 0)
            columns = int(_safe_attr(table, "NumberOfColumns", 0) or 0)
        except Exception as exc:
            warnings.append(
                f"DrawingTable {table_index} metadata could not be read: "
                f"{_format_com_error(exc)}"
            )
            continue
        for row in range(1, rows + 1):
            for column in range(1, columns + 1):
                try:
                    content = str(table.GetCellString(row, column))
                    label, value, separator = _split_label_value(content)
                    canonical = _canonical_field(label)
                    records.append({
                        "table_index": table_index,
                        "table_name": str(_safe_attr(table, "Name", "")),
                        "row": row,
                        "column": column,
                        "text": content,
                        "label": label,
                        "value": value,
                        "separator": separator,
                        "canonical_field": canonical,
                    })
                except Exception as exc:
                    warnings.append(
                        f"DrawingTable {table_index} cell ({row},{column}) "
                        f"could not be read: {_format_com_error(exc)}"
                    )
    return records, warnings


def _table_cell_text(table: Any, row: int, column: int) -> str:
    return str(table.GetCellString(int(row), int(column)))


def _set_table_cell_verified(
    table: Any,
    row: int,
    column: int,
    text: str,
) -> Dict[str, Any]:
    old_text = _table_cell_text(table, row, column)
    try:
        table.SetCellString(int(row), int(column), str(text))
        actual = _table_cell_text(table, row, column)
        verified = actual == str(text)
        error = None
    except Exception as exc:
        actual = old_text
        verified = False
        error = _format_com_error(exc)
    return {
        "old_text": old_text,
        "requested_text": str(text),
        "actual_text": actual,
        "verified": verified,
        "error": error,
    }


def _discover_table_field_targets(background_view: Any) -> Tuple[List[Dict[str, Any]], List[str]]:
    records, warnings = _read_table_records(background_view)
    by_table: Dict[int, List[Dict[str, Any]]] = {}
    for record in records:
        by_table.setdefault(record["table_index"], []).append(record)
    revision_table_indices = {
        record["table_index"]
        for record in records
        if "revisionhistory" in _normalise_label(record.get("text", ""))
    }

    targets: List[Dict[str, Any]] = []
    tables = _safe_attr(background_view, "Tables", None)
    for table_index, table_records in by_table.items():
        if table_index in revision_table_indices:
            continue
        table = tables.Item(table_index)
        rows = int(_safe_attr(table, "NumberOfRows", 0) or 0)
        columns = int(_safe_attr(table, "NumberOfColumns", 0) or 0)
        for record in table_records:
            field = record.get("canonical_field")
            if field is None:
                continue
            if record.get("separator") is not None:
                targets.append({
                    "field": field,
                    "table_index": table_index,
                    "table_name": record["table_name"],
                    "label_row": record["row"],
                    "label_column": record["column"],
                    "target_row": record["row"],
                    "target_column": record["column"],
                    "mode": "same_cell_label_value",
                    "label": record["label"],
                    "separator": record["separator"],
                    "current_value": record["value"],
                })
                continue

            candidates: List[Tuple[int, int, str]] = []
            if record["column"] < columns:
                candidates.append((record["row"], record["column"] + 1, "right_cell"))
            if record["row"] < rows:
                candidates.append((record["row"] + 1, record["column"], "below_cell"))
            selected = None
            for target_row, target_column, mode in candidates:
                try:
                    value = _table_cell_text(table, target_row, target_column)
                except Exception:
                    continue
                selected = (target_row, target_column, mode, value)
                if str(value).strip():
                    break
            if selected is not None:
                targets.append({
                    "field": field,
                    "table_index": table_index,
                    "table_name": record["table_name"],
                    "label_row": record["row"],
                    "label_column": record["column"],
                    "target_row": selected[0],
                    "target_column": selected[1],
                    "mode": selected[2],
                    "label": record["text"],
                    "separator": None,
                    "current_value": selected[3],
                })
    return targets, warnings


def _read_title_field_evidence(background_view: Any) -> Dict[str, Any]:
    text_records, text_warnings = _read_text_records(background_view)
    table_targets, table_warnings = _discover_table_field_targets(background_view)
    fields = {key: [] for key in TITLE_BLOCK_KEYS}
    for record in text_records:
        field = record.get("canonical_field")
        if field in fields:
            current_value = record.get("value", "")
            if record.get("match_mode") == "drawing_text_name":
                current_value = record.get("text", "")
            fields[field].append({
                "kind": "DrawingText",
                "index": record["index"],
                "object_name": record["object_name"],
                "match_mode": record["match_mode"],
                "label": record["label"],
                "separator": record["separator"],
                "current_value": current_value,
            })
    for target in table_targets:
        fields[target["field"]].append({"kind": "DrawingTableCell", **target})
    nonempty = [
        {"field": field, **item}
        for field, items in fields.items()
        for item in items
        if str(item.get("current_value", "")).strip()
    ]
    return {
        "fields": fields,
        "nonempty_dynamic_values": nonempty,
        "nonempty_dynamic_value_count": len(nonempty),
        "warnings": text_warnings + table_warnings,
    }


def _revision_history_inventory(background_view: Any) -> Dict[str, Any]:
    records, warnings = _read_table_records(background_view)
    tables = _safe_attr(background_view, "Tables", None)
    results: List[Dict[str, Any]] = []
    for table_index in sorted({item["table_index"] for item in records}):
        table_records = [item for item in records if item["table_index"] == table_index]
        has_title = any("revisionhistory" in _normalise_label(item["text"]) for item in table_records)
        if not has_title:
            continue
        header_row = None
        for row in sorted({item["row"] for item in table_records}):
            labels = {
                _normalise_label(item["text"])
                for item in table_records
                if item["row"] == row
            }
            header_hits = sum(
                any(token in label for label in labels)
                for token in ("rev", "description", "modified", "approved", "date")
            )
            if header_hits >= 2:
                header_row = row
                break
        table = tables.Item(table_index)
        rows = int(_safe_attr(table, "NumberOfRows", 0) or 0)
        columns = int(_safe_attr(table, "NumberOfColumns", 0) or 0)
        data_cells: List[Dict[str, Any]] = []
        if header_row is not None:
            for row in range(header_row + 1, rows + 1):
                for column in range(1, columns + 1):
                    try:
                        text = _table_cell_text(table, row, column)
                    except Exception:
                        continue
                    if str(text).strip():
                        data_cells.append({
                            "row": row,
                            "column": column,
                            "text": text,
                        })
        results.append({
            "table_index": table_index,
            "table_name": str(_safe_attr(table, "Name", "")),
            "header_row": header_row,
            "data_cells": data_cells,
            "data_cell_count": len(data_cells),
        })
    return {
        "tables": results,
        "data_cell_count": sum(item["data_cell_count"] for item in results),
        "warnings": warnings,
    }


def _delete_non_system_views(drawing_doc: Any, sheet: Any) -> Dict[str, Any]:
    before = _view_inventory(sheet)
    results: List[Dict[str, Any]] = []
    views = sheet.Views
    for index in range(int(views.Count), 2, -1):
        try:
            view = views.Item(index)
            descriptor = {
                "index": index,
                "name": str(_safe_attr(view, "Name", "")),
                "view_type_code": _safe_attr(view, "ViewType", None),
            }
            deletion = _delete_object(drawing_doc, view)
            results.append({"descriptor": descriptor, "deletion": deletion})
        except Exception as exc:
            results.append({
                "descriptor": {"index": index},
                "deletion": {
                    "attempted": True,
                    "succeeded": False,
                    "error": _format_com_error(exc),
                },
            })
    after = _view_inventory(sheet)
    return {
        "before": before,
        "after": after,
        "results": results,
        "verified": after["non_system_view_count"] == 0,
    }


def _clear_revision_history(background_view: Any) -> Dict[str, Any]:
    inventory = _revision_history_inventory(background_view)
    tables = _safe_attr(background_view, "Tables", None)
    results: List[Dict[str, Any]] = []
    for table_info in inventory["tables"]:
        table = tables.Item(table_info["table_index"])
        for cell in table_info["data_cells"]:
            result = _set_table_cell_verified(
                table, cell["row"], cell["column"], ""
            )
            results.append({
                "table_index": table_info["table_index"],
                "row": cell["row"],
                "column": cell["column"],
                "result": result,
            })
    after = _revision_history_inventory(background_view)
    return {
        "before": inventory,
        "results": results,
        "after": after,
        "verified": after["data_cell_count"] == 0,
    }


def _set_template_title_fields(
    drawing_doc: Any,
    sheet: Any,
    properties: Dict[str, str],
    *,
    require_unique: bool = True,
) -> Dict[str, Any]:
    requested = _canonicalise_properties(properties)
    background = _get_background_view(sheet)
    texts = background.Texts
    text_records, text_warnings = _read_text_records(background)
    table_targets, table_warnings = _discover_table_field_targets(background)
    grouped_texts = _group_text_records(text_records)
    tables = _safe_attr(background, "Tables", None)

    candidates: Dict[str, List[Dict[str, Any]]] = {key: [] for key in TITLE_BLOCK_KEYS}
    for field, records in grouped_texts.items():
        for record in records:
            candidates[field].append({"kind": "DrawingText", "record": record})
    for target in table_targets:
        candidates[target["field"]].append({"kind": "DrawingTableCell", "target": target})

    ambiguous = {
        field: items for field, items in candidates.items()
        if field in requested and len(items) > 1
    }
    missing = [field for field in requested if not candidates[field]]
    if require_unique and ambiguous:
        return {
            "verified": False,
            "preflight_verified": False,
            "requested": requested,
            "ambiguous_fields": {
                field: len(items) for field, items in ambiguous.items()
            },
            "missing_fields": missing,
            "results": {},
            "rollback": {"attempted": False, "verified": True},
            "warnings": text_warnings + table_warnings,
        }

    undo: List[Dict[str, Any]] = []
    results: Dict[str, Any] = {}
    failed = False
    for field, value in requested.items():
        field_candidates = candidates[field]
        if not field_candidates:
            results[field] = {
                "action": "not_found",
                "verified": False,
                "error": "No matching DrawingText or DrawingTable cell.",
            }
            failed = True
            break
        candidate = field_candidates[0]
        if candidate["kind"] == "DrawingText":
            record = candidate["record"]
            obj = texts.Item(record["index"])
            old_text = str(obj.Text)
            if record.get("separator") is not None:
                requested_text = (
                    f"{record['label']}{record['separator']} {value}"
                )
            elif record.get("match_mode") == "drawing_text_name":
                requested_text = str(value)
            else:
                requested_text = f"{record['label']}: {value}"
            try:
                obj.Text = requested_text
                actual = str(obj.Text)
                verified = actual == requested_text
                error = None
            except Exception as exc:
                actual = old_text
                verified = False
                error = _format_com_error(exc)
            if verified:
                undo.append({"kind": "text", "object": obj, "old_text": old_text})
            results[field] = {
                "kind": "DrawingText",
                "requested_text": requested_text,
                "actual_text": actual,
                "verified": verified,
                "error": error,
            }
        else:
            target = candidate["target"]
            table = tables.Item(target["table_index"])
            if target["mode"] == "same_cell_label_value":
                requested_text = (
                    f"{target['label']}{target['separator'] or ':'} {value}"
                )
            else:
                requested_text = str(value)
            result = _set_table_cell_verified(
                table,
                target["target_row"],
                target["target_column"],
                requested_text,
            )
            verified = bool(result["verified"])
            if verified:
                undo.append({
                    "kind": "table",
                    "table": table,
                    "row": target["target_row"],
                    "column": target["target_column"],
                    "old_text": result["old_text"],
                })
            results[field] = {
                "kind": "DrawingTableCell",
                "target": target,
                **result,
            }
        if not verified:
            failed = True
            break

    rollback_results: List[Dict[str, Any]] = []
    if failed:
        for item in reversed(undo):
            if item["kind"] == "text":
                try:
                    item["object"].Text = item["old_text"]
                    ok = str(item["object"].Text) == item["old_text"]
                    error = None
                except Exception as exc:
                    ok = False
                    error = _format_com_error(exc)
            else:
                restore = _set_table_cell_verified(
                    item["table"], item["row"], item["column"], item["old_text"]
                )
                ok = bool(restore["verified"])
                error = restore.get("error")
            rollback_results.append({"verified": ok, "error": error})

    return {
        "verified": not failed,
        "preflight_verified": not bool(require_unique and ambiguous),
        "requested": requested,
        "ambiguous_fields": {
            field: len(items) for field, items in ambiguous.items()
        },
        "missing_fields": missing,
        "results": results,
        "rollback": {
            "attempted": failed and bool(undo),
            "results": rollback_results,
            "verified": all(item["verified"] for item in rollback_results)
            if rollback_results else True,
        },
        "warnings": text_warnings + table_warnings,
    }


def _clear_template_title_fields(
    drawing_doc: Any,
    sheet: Any,
) -> Dict[str, Any]:
    background = _get_background_view(sheet)
    texts = background.Texts
    text_records, text_warnings = _read_text_records(background)
    table_targets, table_warnings = _discover_table_field_targets(background)
    tables = _safe_attr(background, "Tables", None)
    before = _read_title_field_evidence(background)

    undo: List[Dict[str, Any]] = []
    results: List[Dict[str, Any]] = []
    failed = False

    for record in text_records:
        if record.get("canonical_field") not in TITLE_BLOCK_KEYS:
            continue
        obj = texts.Item(record["index"])
        old_text = str(obj.Text)
        if record.get("separator") is not None:
            requested_text = (
                f"{record['label']}{record['separator']} "
            )
        elif record.get("match_mode") == "drawing_text_name":
            requested_text = ""
        else:
            requested_text = f"{record['label']}: "
        try:
            obj.Text = requested_text
            actual = str(obj.Text)
            verified = actual == requested_text
            error = None
        except Exception as exc:
            actual = old_text
            verified = False
            error = _format_com_error(exc)
        if verified:
            undo.append({"kind": "text", "object": obj, "old_text": old_text})
        results.append({
            "field": record.get("canonical_field"),
            "kind": "DrawingText",
            "index": record["index"],
            "old_text": old_text,
            "requested_text": requested_text,
            "actual_text": actual,
            "verified": verified,
            "error": error,
        })
        if not verified:
            failed = True
            break

    if not failed:
        seen_targets = set()
        for target in table_targets:
            key = (
                target["table_index"],
                target["target_row"],
                target["target_column"],
            )
            if key in seen_targets:
                continue
            seen_targets.add(key)
            table = tables.Item(target["table_index"])
            if target["mode"] == "same_cell_label_value":
                requested_text = (
                    f"{target['label']}{target['separator'] or ':'} "
                )
            else:
                requested_text = ""
            result = _set_table_cell_verified(
                table,
                target["target_row"],
                target["target_column"],
                requested_text,
            )
            if result["verified"]:
                undo.append({
                    "kind": "table",
                    "table": table,
                    "row": target["target_row"],
                    "column": target["target_column"],
                    "old_text": result["old_text"],
                })
            results.append({
                "field": target["field"],
                "kind": "DrawingTableCell",
                "target": target,
                **result,
            })
            if not result["verified"]:
                failed = True
                break

    rollback_results: List[Dict[str, Any]] = []
    if failed:
        for item in reversed(undo):
            if item["kind"] == "text":
                try:
                    item["object"].Text = item["old_text"]
                    ok = str(item["object"].Text) == item["old_text"]
                    error = None
                except Exception as exc:
                    ok = False
                    error = _format_com_error(exc)
            else:
                restore = _set_table_cell_verified(
                    item["table"], item["row"], item["column"], item["old_text"]
                )
                ok = bool(restore["verified"])
                error = restore.get("error")
            rollback_results.append({"verified": ok, "error": error})

    after = _read_title_field_evidence(background)
    # Revision-history rows are cleared separately before this function is
    # called. Any remaining recognised value therefore represents stale
    # title-block data or an unhandled duplicate field.
    verified = bool(
        not failed
        and after["nonempty_dynamic_value_count"] == 0
    )
    return {
        "attempted": bool(results),
        "verified": verified,
        "before": before,
        "results": results,
        "after": after,
        "rollback": {
            "attempted": failed and bool(undo),
            "results": rollback_results,
            "verified": all(item["verified"] for item in rollback_results)
            if rollback_results else True,
        },
        "warnings": text_warnings + table_warnings,
    }


def _sheet_template_audit(catia_app: Any, sheet: Any, index: int) -> Dict[str, Any]:
    background = _get_background_view(sheet)
    views = _view_inventory(sheet)
    title_fields = _read_title_field_evidence(background)
    revision = _revision_history_inventory(background)
    return {
        "sheet_index": index,
        "sheet_name": str(_safe_attr(sheet, "Name", "")),
        "paper": _paper_dimensions(catia_app, sheet),
        "projection": _detect_projection_method(catia_app, sheet),
        "scale": _sheet_scale(sheet),
        "views": views,
        "background_inventory": _content_inventory(background),
        "title_fields": title_fields,
        "revision_history": revision,
        "clean_for_reuse": bool(
            views["non_system_view_count"] == 0
            and title_fields["nonempty_dynamic_value_count"] == 0
            and revision["data_cell_count"] == 0
        ),
    }


def _audit_open_drawing_template(catia_app: Any, document: Any) -> Dict[str, Any]:
    sheets = document.Sheets
    sheet_results = [
        _sheet_template_audit(catia_app, sheets.Item(index), index)
        for index in range(1, int(sheets.Count) + 1)
    ]
    return {
        "document_name": str(_safe_attr(document, "Name", "")),
        "document_full_name": str(_safe_attr(document, "FullName", "")),
        "sheet_count": int(sheets.Count),
        "sheets": sheet_results,
        "non_system_view_count": sum(
            item["views"]["non_system_view_count"] for item in sheet_results
        ),
        "generative_link_count": sum(
            item["views"]["generative_link_count"] for item in sheet_results
        ),
        "nonempty_dynamic_value_count": sum(
            item["title_fields"]["nonempty_dynamic_value_count"]
            for item in sheet_results
        ),
        "revision_history_data_cell_count": sum(
            item["revision_history"]["data_cell_count"]
            for item in sheet_results
        ),
        "clean_for_reuse": all(item["clean_for_reuse"] for item in sheet_results),
    }


def audit_drawing_template(
    catia_app: Any,
    template_path: Optional[str] = None,
) -> Dict[str, Any]:
    path = _resolve_template_resource(template_path)
    sha_before = _sha256_file(path)
    existing = _find_open_document(catia_app, path)
    opened_by_tool = existing is None
    document = existing or catia_app.Documents.Open(path)
    try:
        audit = _audit_open_drawing_template(catia_app, document)
    finally:
        cleanup = (
            _close_document(document, discard_unsaved_changes=False)
            if opened_by_tool else
            {"attempted": False, "succeeded": True, "reason": "preexisting document preserved"}
        )
    sha_after = _sha256_file(path)
    return _success({
        "operation": "audit_drawing_template",
        "template_path": path,
        "template_sha256_before": sha_before,
        "template_sha256_after": sha_after,
        "template_source_protected": sha_before == sha_after,
        "audit": audit,
        "cleanup": cleanup,
        "model_modified": False,
        "document_save_required": False,
    })


def _sanitise_open_template_document(
    catia_app: Any,
    document: Any,
    *,
    remove_non_system_views: bool,
    clear_title_fields: bool,
    clear_revision_history: bool,
) -> Dict[str, Any]:
    sheets = document.Sheets
    results: List[Dict[str, Any]] = []
    verified = True
    modified = False
    document_activation = _activate_document(document)
    for index in range(1, int(sheets.Count) + 1):
        sheet = sheets.Item(index)
        try:
            sheet.Activate()
            sheet_activation = {"attempted": True, "succeeded": True, "error": None}
        except Exception as exc:
            sheet_activation = {
                "attempted": True,
                "succeeded": False,
                "error": _format_com_error(exc),
            }
        background = _get_background_view(sheet)
        background_activation = _activate_view(background)
        view_cleanup = (
            _delete_non_system_views(document, sheet)
            if remove_non_system_views else
            {"verified": True, "attempted": False, "reason": "disabled"}
        )
        revision_cleanup = (
            _clear_revision_history(background)
            if clear_revision_history else
            {"verified": True, "attempted": False, "reason": "disabled"}
        )
        title_cleanup = (
            _clear_template_title_fields(document, sheet)
            if clear_title_fields else
            {"verified": True, "attempted": False, "reason": "disabled"}
        )
        update_attempts, update_warnings = _update_drawing(document, sheet)
        main_view_restore = _activate_view(sheet.Views.Item(1))
        item_verified = bool(
            view_cleanup["verified"]
            and title_cleanup["verified"]
            and revision_cleanup["verified"]
        )
        verified = verified and item_verified
        modified = modified or bool(
            view_cleanup.get("results")
            or title_cleanup.get("attempted")
            or revision_cleanup.get("results")
        )
        results.append({
            "sheet_index": index,
            "sheet_name": str(_safe_attr(sheet, "Name", "")),
            "document_activation": document_activation,
            "sheet_activation": sheet_activation,
            "background_activation": background_activation,
            "main_view_restore": main_view_restore,
            "view_cleanup": view_cleanup,
            "title_cleanup": title_cleanup,
            "revision_cleanup": revision_cleanup,
            "update_attempts": update_attempts,
            "warnings": update_warnings,
            "verified": item_verified,
        })
    audit = _audit_open_drawing_template(catia_app, document)
    verified = bool(verified and audit["clean_for_reuse"])
    return {
        "verified": verified,
        "modified": modified,
        "sheets": results,
        "post_sanitisation_audit": audit,
    }


def create_clean_drawing_template(
    catia_app: Any,
    output_path: str,
    source_template_path: Optional[str] = None,
    overwrite: bool = False,
    remove_non_system_views: bool = True,
    clear_title_fields: bool = True,
    clear_revision_history: bool = True,
) -> Dict[str, Any]:
    source = _resolve_template_resource(source_template_path)
    target = _normalised_output_drawing_path(output_path)
    if _find_open_document(catia_app, target) is not None:
        raise RuntimeError("Output template is already open in CATIA.")
    source_sha_before = _sha256_file(source)
    transaction = _begin_template_file_transaction(source, target, overwrite)
    document = None
    try:
        document = catia_app.Documents.Open(target)
        _get_drawing_document_from_object(document)
        sanitisation = _sanitise_open_template_document(
            catia_app,
            document,
            remove_non_system_views=bool(remove_non_system_views),
            clear_title_fields=bool(clear_title_fields),
            clear_revision_history=bool(clear_revision_history),
        )
        if not sanitisation["verified"]:
            raise ToolOperationError(
                "Template sanitisation did not pass verified readback.",
                data={"sanitisation": sanitisation, "file_transaction": transaction},
            )
        document.Save()
        saved_verified = _document_saved(document) is True
        close_result = _close_document(document, discard_unsaved_changes=False)
        document = None
        if not close_result.get("succeeded"):
            raise RuntimeError(f"Clean template could not be closed: {close_result.get('error')}")
        final_audit_result = audit_drawing_template(catia_app, target)
        final_audit = final_audit_result["data"]["audit"]
        if not final_audit["clean_for_reuse"]:
            raise ToolOperationError(
                "Saved output template failed final clean-template audit.",
                data={
                    "sanitisation": sanitisation,
                    "final_audit": final_audit,
                    "file_transaction": transaction,
                },
            )
        _commit_template_file_transaction(transaction)
        source_sha_after = _sha256_file(source)
        return _success({
            "operation": "create_clean_drawing_template",
            "source_template_path": source,
            "source_sha256_before": source_sha_before,
            "source_sha256_after": source_sha_after,
            "source_template_protected": source_sha_before == source_sha_after,
            "output_path": target,
            "output_sha256": _sha256_file(target),
            "sanitisation": sanitisation,
            "final_audit": final_audit,
            "document_saved_verified": saved_verified,
            "close_result": close_result,
            "file_transaction": transaction,
            "model_modified": True,
            "document_save_required": False,
        })
    except Exception:
        if document is not None:
            _close_document(document, discard_unsaved_changes=True)
        _rollback_template_file_transaction(transaction)
        raise


def _get_drawing_document_from_object(document: Any) -> Any:
    try:
        _ = int(document.Sheets.Count)
    except Exception as exc:
        raise RuntimeError("Opened template copy is not a CATDrawing.") from exc
    return document


def create_drawing_from_template(
    catia_app: Any,
    output_path: str,
    template_path: Optional[str] = None,
    overwrite: bool = False,
    title_block_properties: Optional[Dict[str, str]] = None,
    sanitise_template_copy: bool = True,
    clear_revision_history: bool = True,
    require_clean_result: bool = True,
    close_after_create: bool = False,
) -> Dict[str, Any]:
    source = _resolve_template_resource(template_path)
    target = _normalised_output_drawing_path(output_path)
    if _find_open_document(catia_app, target) is not None:
        raise RuntimeError("Output CATDrawing is already open in CATIA.")
    source_sha_before = _sha256_file(source)
    transaction = _begin_template_file_transaction(source, target, overwrite)
    document = None
    try:
        document = catia_app.Documents.Open(target)
        _get_drawing_document_from_object(document)
        sanitisation = None
        if sanitise_template_copy:
            sanitisation = _sanitise_open_template_document(
                catia_app,
                document,
                remove_non_system_views=True,
                clear_title_fields=True,
                clear_revision_history=bool(clear_revision_history),
            )
            if require_clean_result and not sanitisation["verified"]:
                raise ToolOperationError(
                    "Template copy could not be sanitised to a verified clean drawing.",
                    data={"sanitisation": sanitisation, "file_transaction": transaction},
                )

        title_result = None
        if title_block_properties:
            sheet = document.Sheets.ActiveSheet
            title_result = _set_template_title_fields(
                document,
                sheet,
                title_block_properties,
                require_unique=True,
            )
            if not title_result["verified"]:
                raise ToolOperationError(
                    "One or more title-block fields could not be updated atomically.",
                    data={
                        "title_block_update": title_result,
                        "sanitisation": sanitisation,
                        "file_transaction": transaction,
                    },
                )

        document.Update()
        document.Save()
        saved_verified = _document_saved(document) is True
        audit = _audit_open_drawing_template(catia_app, document)
        # Once user-supplied title values are applied, nonempty fields are expected.
        structural_clean = bool(
            audit["non_system_view_count"] == 0
            and audit["generative_link_count"] == 0
            and audit["revision_history_data_cell_count"] == 0
        )
        if require_clean_result and not structural_clean:
            raise ToolOperationError(
                "Created CATDrawing still contains model-specific views, links, "
                "or revision-history data.",
                data={"audit": audit, "file_transaction": transaction},
            )
        _commit_template_file_transaction(transaction)
        source_sha_after = _sha256_file(source)
        close_result = {
            "attempted": False,
            "succeeded": True,
            "reason": "document left open for drafting",
        }
        if close_after_create:
            close_result = _close_document(document, discard_unsaved_changes=False)
            document = None
        return _success({
            "operation": "create_drawing_from_template",
            "template_path": source,
            "template_resolution": (
                "explicit" if str(template_path or "").strip() else
                "environment" if str(os.getenv(DEFAULT_TEMPLATE_ENV, "")).strip() else
                "built_in_default"
            ),
            "template_sha256_before": source_sha_before,
            "template_sha256_after": source_sha_after,
            "template_source_protected": source_sha_before == source_sha_after,
            "output_path": target,
            "output_sha256": _sha256_file(target),
            "sanitisation": sanitisation,
            "title_block_update": title_result,
            "audit": audit,
            "structural_clean_result": structural_clean,
            "document_saved_verified": saved_verified,
            "close_result": close_result,
            "document_left_open": not close_after_create,
            "file_transaction": transaction,
            "model_modified": True,
            "document_save_required": False,
        })
    except Exception:
        if document is not None:
            _close_document(document, discard_unsaved_changes=True)
        _rollback_template_file_transaction(transaction)
        raise


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------

def register_tools(mcp: Any, ctx: Any) -> list[str]:
    """Register drawing-template tools using the project-standard decorator format."""

    conn = ctx.conn
    names: list[str] = []

    implementations = {
        "get_title_block_properties": globals()["get_title_block_properties"],
        "set_title_block_properties": globals()["set_title_block_properties"],
        "load_drawing_frame": globals()["load_drawing_frame"],
        "detect_projection_standard": globals()["detect_projection_standard"],
        "audit_drawing_template": globals()["audit_drawing_template"],
        "create_clean_drawing_template": globals()["create_clean_drawing_template"],
        "create_drawing_from_template": globals()["create_drawing_from_template"],
    }

    def _call(name: str, **kwargs: Any) -> Dict[str, Any]:
        try:
            catia_app = conn.connect(visible=True)
            return implementations[name](catia_app, **kwargs)
        except ToolOperationError as exc:
            return _error(
                str(exc),
                data=exc.data,
                warnings=exc.warnings,
                status=exc.status,
            )
        except Exception as exc:
            return _error(_format_com_error(exc))

    @mcp.tool()
    def get_title_block_properties(
        sheet_index: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Read verified title-block fields from a CATDrawing sheet."""
        return _call(
            "get_title_block_properties",
            sheet_index=sheet_index,
        )

    names.append("get_title_block_properties")

    @mcp.tool()
    def set_title_block_properties(
        properties: Dict[str, str],
        sheet_index: Optional[int] = None,
        create_if_missing: bool = False,
        require_unique: bool = True,
        base_x_mm: Optional[float] = None,
        base_y_mm: Optional[float] = None,
        row_spacing_mm: float = 7.0,
        sheet_margin_mm: float = 5.0,
    ) -> Dict[str, Any]:
        """Update or create verified DrawingText title-block fields."""
        return _call(
            "set_title_block_properties",
            properties=properties,
            sheet_index=sheet_index,
            create_if_missing=create_if_missing,
            require_unique=require_unique,
            base_x_mm=base_x_mm,
            base_y_mm=base_y_mm,
            row_spacing_mm=row_spacing_mm,
            sheet_margin_mm=sheet_margin_mm,
        )

    names.append("set_title_block_properties")

    @mcp.tool()
    def load_drawing_frame(
        template_path: str = "",
        replace_existing: bool = True,
        sheet_index: Optional[int] = None,
        apply_sheet_settings: bool = False,
    ) -> Dict[str, Any]:
        """Copy a verified template frame into an already-open target sheet."""
        try:
            resolved = _resolve_template_resource(template_path)
        except Exception as exc:
            return _error(_format_com_error(exc))
        return _call(
            "load_drawing_frame",
            template_path=resolved,
            replace_existing=replace_existing,
            sheet_index=sheet_index,
            apply_sheet_settings=apply_sheet_settings,
        )

    names.append("load_drawing_frame")

    @mcp.tool()
    def detect_projection_standard(
        sheet_index: Optional[int] = None,
    ) -> Dict[str, Any]:
        """Detect first-angle or third-angle projection for a sheet."""
        return _call(
            "detect_projection_standard",
            sheet_index=sheet_index,
        )

    names.append("detect_projection_standard")

    @mcp.tool()
    def audit_drawing_template(
        template_path: str = "",
    ) -> Dict[str, Any]:
        """Audit a CATDrawing template without saving or modifying it."""
        return _call(
            "audit_drawing_template",
            template_path=template_path or None,
        )

    names.append("audit_drawing_template")

    @mcp.tool()
    def create_clean_drawing_template(
        output_path: str,
        source_template_path: str = "",
        overwrite: bool = False,
        remove_non_system_views: bool = True,
        clear_title_fields: bool = True,
        clear_revision_history: bool = True,
    ) -> Dict[str, Any]:
        """Create and verify a clean reusable CATDrawing template copy."""
        return _call(
            "create_clean_drawing_template",
            output_path=output_path,
            source_template_path=source_template_path or None,
            overwrite=overwrite,
            remove_non_system_views=remove_non_system_views,
            clear_title_fields=clear_title_fields,
            clear_revision_history=clear_revision_history,
        )

    names.append("create_clean_drawing_template")

    @mcp.tool()
    def create_drawing_from_template(
        output_path: str,
        template_path: str = "",
        overwrite: bool = False,
        title_block_properties: Optional[Dict[str, str]] = None,
        sanitise_template_copy: bool = True,
        clear_revision_history: bool = True,
        require_clean_result: bool = True,
        close_after_create: bool = False,
    ) -> Dict[str, Any]:
        """Copy the managed template, sanitise it and open a new CATDrawing."""
        return _call(
            "create_drawing_from_template",
            output_path=output_path,
            template_path=template_path or None,
            overwrite=overwrite,
            title_block_properties=title_block_properties,
            sanitise_template_copy=sanitise_template_copy,
            clear_revision_history=clear_revision_history,
            require_clean_result=require_clean_result,
            close_after_create=close_after_create,
        )

    names.append("create_drawing_from_template")

    return names
