"""
drafting.py
Version: drafting-safe-frame-2026-08-19-v9

CATIA V5 MCP generative drafting tools.

Key corrections:
- Front is created with DefineFrontView; Top and Right are true projected
  views created with DefineProjectionView and linked to Front.
- The sheet projection method is explicit (third-angle by default).
- View locations are calculated from the real paper width/height instead of
  using A3-only fixed coordinates.
- The localized background-view name is not required. CATIA's official
  collection order is used: main view at index 1, background view at index 2.
- Paper, orientation, scale, view type, parent/reference links, projection
  plane, view bounding box and generated geometry are read back and verified.
- GenerateDimensions reports the actual dimension-count delta. A successful
  API call that creates zero dimensions is success_with_warnings, not a false
  claim that dimensions were generated.
- Text/table creation, CATDrawing save and PDF/DWG/DXF export are verified.
- Invalid options are rejected instead of being silently ignored.
- Paper dimensions use a compatibility chain: direct COM return values,
  CATIA-side no-argument function calls, typed ByRef variants, then verified
  standard-paper dimensions derived from the read-back paper/orientation codes.
- The CatPaperSize numeric mapping follows the real CATIA enum values:
  Letter=0, Legal=1, A0=2, A1=3, A2=4, A3=5, A4=6, User=13.
- Standard-paper dimensions are checked against the requested paper size and
  orientation, so an enum mismatch cannot silently pass validation.
- Projected views receive the represented document, are aligned and positioned
  after their parent link is established, then enter a bounded ForceUpdate /
  viewer-refresh / readback loop until IsGenerative and non-empty geometry are
  both verified.
- PDF/DWG/DXF export first executes ExportData inside the CATIA process through
  SystemService.Evaluate, then uses direct COM compatibility fallbacks, waits
  for filesystem materialization and reports every attempt.
- Three-view layout is finalized from the actual DrawingView.Size bounding
  boxes, not merely from view-origin coordinates. Front/Top, Front/Right and
  Top/Right must satisfy a configurable minimum clear gap.
- Pairwise rectangle intersection, directional edge gaps and sheet-boundary
  clearance are returned and are part of creation success criteria.
- Projected views are explicitly unaligned before positioning. ReferenceView
  remains linked, but CATIA's positional alignment constraint is removed so
  Top can move vertically and Right can move horizontally without snapping
  back during ForceUpdate.
- View positioning uses CATIA-side SystemService.Evaluate first, direct x/y
  setters second, and xAxisData/yAxisData as a final compatibility fallback.
  Every strategy is verified by position readback.
- After local pairwise separation, the union bounding box of Front/Top/Right
  is compared with the paper's usable area. If the group fits but crosses a
  sheet margin, all three views are translated by the same dx/dy.
- Global group translation preserves pairwise gaps, projected-view reference
  links and generated geometry. A layout is rejected as insufficient space
  only when the union bounding box is genuinely larger than the usable area
  or verified group positioning fails.
- Failed drawing creation attempts close the newly created CATDrawing.
- Failure cleanup temporarily suppresses CATIA file-alert dialogs and restores the prior setting.
- Export verification protects unrelated pre-existing side-effect files such as .pdf.pdf.
- Standalone title-block placement is paper-size-aware when x/y are omitted and verifies sheet bounds.
"""

from __future__ import annotations

import math
import os
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

from catia_mcp.connection import CATIAError, normalize_path, safe_str


IMPLEMENTATION_VERSION = "drafting-safe-frame-2026-08-19-v9"
_CATVB_SCRIPT_LANGUAGE = 1

PAPER_SIZE_ENUM = {
    "LETTER": 0,
    "LEGAL": 1,
    "A0": 2,
    "A1": 3,
    "A2": 4,
    "A3": 5,
    "A4": 6,
    "USER": 13,
}

ORIENTATION_ENUM = {
    "PORTRAIT": 0,
    "LANDSCAPE": 1,
}

PROJECTION_METHODS = {
    "FIRST_ANGLE": "first_angle",
    "THIRD_ANGLE": "third_angle",
}

EXPORT_FORMATS = {
    "pdf": ".pdf",
    "dwg": ".dwg",
    "dxf": ".dxf",
}

TITLE_BLOCK_ROWS = 6
TITLE_BLOCK_COLUMNS = 2
TITLE_BLOCK_ROW_HEIGHT_MM = 8.0
TITLE_BLOCK_COLUMN_WIDTH_MM = 45.0
TITLE_BLOCK_WIDTH_MM = (
    TITLE_BLOCK_COLUMNS * TITLE_BLOCK_COLUMN_WIDTH_MM
)
TITLE_BLOCK_HEIGHT_MM = (
    TITLE_BLOCK_ROWS * TITLE_BLOCK_ROW_HEIGHT_MM
)
DEFAULT_TITLE_BLOCK_MARGIN_MM = 5.0
DEFAULT_TITLE_BLOCK_Y_MM = 35.0

# Additional keep-out band measured inward from the existing sheet-margin boundary.
# The template/background geometry is never inspected or modified by this rule.
INNER_FRAME_INSET_MM = 40.0

STANDARD_PAPER_PORTRAIT_MM = {
    "A0": (841.0, 1189.0),
    "A1": (594.0, 841.0),
    "A2": (420.0, 594.0),
    "A3": (297.0, 420.0),
    "A4": (210.0, 297.0),
    "LETTER": (215.9, 279.4),
    "LEGAL": (215.9, 355.6),
}

PAPER_SIZE_NAME_BY_CODE = {
    value: key
    for key, value in PAPER_SIZE_ENUM.items()
}

ORIENTATION_NAME_BY_CODE = {
    value: key
    for key, value in ORIENTATION_ENUM.items()
}

DRAWING_VIEW_TYPE_NAMES = {
    0: "background",
    1: "front",
    2: "left",
    3: "right",
    4: "top",
    5: "bottom",
    6: "rear",
    7: "auxiliary",
    8: "isometric",
    9: "section",
    10: "section_cut",
    11: "detail",
    12: "untyped",
    13: "main",
    14: "pure_sketch",
    15: "unfolded",
}


@dataclass(frozen=True)
class DrawingLayout:
    front_x: float
    front_y: float
    top_x: float
    top_y: float
    right_x: float
    right_y: float
    title_block_x: float
    title_block_y: float
    notes_x: float
    notes_y: float


class DraftingOperationError(RuntimeError):
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


# ---------------------------------------------------------------------------
# Standard result helpers
# ---------------------------------------------------------------------------

def _success(
    data: Any,
    warnings: Optional[list[str]] = None,
) -> dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": True,
        "status": (
            "success_with_warnings"
            if warning_list
            else "success"
        ),
        "data": data,
        "warnings": warning_list,
    }


def _error(
    message: str,
    *,
    data: Any = None,
    warnings: Optional[list[str]] = None,
    status: str = "error",
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": False,
        "status": status,
        "error": str(message),
        "warnings": list(warnings or []),
    }
    if data is not None:
        result["data"] = data
    return result


# ---------------------------------------------------------------------------
# Generic validation and COM helpers
# ---------------------------------------------------------------------------

def _format_com_error(exc: BaseException) -> str:
    details = getattr(exc, "excepinfo", None)
    if details and len(details) >= 3 and details[2]:
        return str(details[2])
    return str(exc)


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
        raise ValueError(
            f"{parameter_name} must be a finite number."
        ) from exc
    if not math.isfinite(number):
        raise ValueError(f"{parameter_name} must be finite.")
    return number


def _positive_float(value: Any, parameter_name: str) -> float:
    number = _finite_float(value, parameter_name)
    if number <= 0.0:
        raise ValueError(
            f"{parameter_name} must be greater than zero."
        )
    return number


def _normalise_choice(
    value: Any,
    parameter_name: str,
    allowed: dict[str, Any],
) -> str:
    key = _nonempty_text(value, parameter_name).upper()
    if key not in allowed:
        choices = ", ".join(sorted(allowed))
        raise ValueError(
            f"{parameter_name} must be one of: {choices}."
        )
    return key


def _safe_attribute(
    obj: Any,
    name: str,
    default: Any = None,
) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _safe_count(collection: Any) -> Optional[int]:
    try:
        return int(collection.Count)
    except Exception:
        return None


def _evaluate(
    application: Any,
    script: str,
    function_name: str,
    parameters: list[Any],
) -> Any:
    try:
        system_service = application.SystemService
    except Exception as exc:
        raise RuntimeError(
            f"Cannot access CATIA SystemService: {exc}"
        ) from exc

    try:
        return system_service.Evaluate(
            script,
            _CATVB_SCRIPT_LANGUAGE,
            function_name,
            parameters,
        )
    except Exception as exc:
        raise RuntimeError(
            f"SystemService.Evaluate failed for "
            f"{function_name}: {_format_com_error(exc)}"
        ) from exc


def _numeric_sequence(
    value: Any,
    expected_length: int,
) -> list[float]:
    try:
        sequence = list(value)
    except Exception as exc:
        raise RuntimeError(
            f"CATIA did not return an array: {exc}"
        ) from exc

    if len(sequence) != expected_length:
        raise RuntimeError(
            f"CATIA returned {len(sequence)} values; "
            f"{expected_length} were required."
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


def _document_name(document: Any) -> str:
    return safe_str(_safe_attribute(document, "Name", ""))


def _document_full_name(document: Any) -> str:
    return safe_str(_safe_attribute(document, "FullName", ""))


def _document_saved(document: Any) -> Optional[bool]:
    try:
        return bool(document.Saved)
    except Exception:
        return None


def _normalised_path_key(path: str) -> str:
    return os.path.normcase(os.path.abspath(path))


def _find_open_document_by_path(
    application: Any,
    path: str,
) -> Any:
    target = _normalised_path_key(path)
    documents = application.Documents

    for index in range(1, int(documents.Count) + 1):
        document = documents.Item(index)
        full_name = _document_full_name(document)
        if full_name and _normalised_path_key(full_name) == target:
            return document
    return None


def _close_document(
    document: Any,
    *,
    discard_unsaved_changes: bool = True,
) -> dict[str, Any]:
    """Close a document without leaving CATIA blocked on a save prompt.

    Failure cleanup may close an unsaved CATDrawing.  CATIA can display a
    modal confirmation dialog in that case, which blocks an MCP request.  The
    caller's DisplayFileAlerts value is therefore disabled only for the close
    operation and restored in a finally block.
    """

    details = {
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

    application = _safe_attribute(document, "Application")
    previous_file_alerts: Optional[bool] = None

    try:
        if discard_unsaved_changes and application is not None:
            try:
                previous_file_alerts = bool(
                    application.DisplayFileAlerts
                )
                application.DisplayFileAlerts = False
                details[
                    "file_alerts_temporarily_disabled"
                ] = True
            except Exception as exc:
                details["file_alerts_disable_error"] = (
                    _format_com_error(exc)
                )

        document.Close()
        details["succeeded"] = True
    except Exception as exc:
        details["succeeded"] = False
        details["error"] = _format_com_error(exc)
    finally:
        if previous_file_alerts is not None:
            try:
                application.DisplayFileAlerts = previous_file_alerts
            except Exception as exc:
                details["file_alerts_restore_error"] = (
                    _format_com_error(exc)
                )

    return details


def _require_active_drawing_document(conn: Any) -> Any:
    return conn.get_active_drawing_document()


def _represented_3d_object(document: Any) -> Any:
    try:
        return document.Part
    except Exception:
        pass
    try:
        return document.Product
    except Exception:
        pass
    return document


def _call_update(
    sheet: Any,
    drawing_document: Any,
) -> tuple[list[dict[str, Any]], list[str]]:
    attempts: list[dict[str, Any]] = []
    warnings: list[str] = []

    for target_name, target, method_names in (
        (
            "sheet",
            sheet,
            ("ForceUpdate", "Update"),
        ),
        (
            "drawing_document",
            drawing_document,
            ("Update",),
        ),
    ):
        succeeded = False
        errors: list[str] = []
        selected_method: Optional[str] = None

        for method_name in method_names:
            try:
                getattr(target, method_name)()
                selected_method = method_name
                succeeded = True
                break
            except Exception as exc:
                errors.append(
                    f"{method_name}: {_format_com_error(exc)}"
                )

        attempts.append(
            {
                "target": target_name,
                "succeeded": succeeded,
                "selected_method": selected_method,
                "errors": errors,
            }
        )
        if not succeeded:
            warnings.append(
                f"Could not update {target_name}: {errors}"
            )

    return attempts, warnings


# ---------------------------------------------------------------------------
# Path and output verification
# ---------------------------------------------------------------------------

def _normalise_model_path(model_path: Any) -> str:
    raw = _nonempty_text(model_path, "model_path")
    normalized = normalize_path(raw)
    path = Path(normalized)

    if not path.exists():
        raise CATIAError(
            f"Model file does not exist: {normalized}"
        )
    if not path.is_file():
        raise CATIAError(
            f"Model path is not a file: {normalized}"
        )
    if path.suffix.lower() not in {".catpart", ".catproduct"}:
        raise CATIAError(
            "model_path must point to a CATPart or CATProduct file."
        )
    return str(path.resolve())


def _prepare_output_path(
    value: Any,
    parameter_name: str,
    expected_suffix: str,
    overwrite: bool,
) -> str:
    raw = _nonempty_text(value, parameter_name)
    normalized = normalize_path(raw)
    path = Path(normalized)

    if path.suffix.lower() != expected_suffix.lower():
        raise ValueError(
            f"{parameter_name} must end with "
            f"'{expected_suffix}'."
        )

    path.parent.mkdir(parents=True, exist_ok=True)

    if path.exists() and not overwrite:
        raise FileExistsError(
            f"Output file already exists and overwrite=false: "
            f"{path}"
        )

    return str(path.resolve())


def _remove_existing_output(
    path: str,
    overwrite: bool,
) -> None:
    target = Path(path)
    if not target.exists():
        return
    if not overwrite:
        raise FileExistsError(
            f"Output file already exists: {target}"
        )
    target.unlink()


def _file_verification(path: str) -> dict[str, Any]:
    target = Path(path)
    exists = target.exists() and target.is_file()
    stat = target.stat() if exists else None
    size = stat.st_size if stat is not None else 0
    return {
        "path": str(target),
        "exists": exists,
        "size_bytes": size,
        "modified_time_ns": (
            stat.st_mtime_ns if stat is not None else None
        ),
        "verified": bool(exists and size > 0),
    }


def _wait_for_nonempty_file(
    candidates: list[Path],
    timeout_seconds: float = 12.0,
    poll_interval_seconds: float = 0.20,
    baseline: Optional[dict[str, dict[str, Any]]] = None,
) -> tuple[Optional[Path], list[dict[str, Any]]]:
    """Wait for a candidate created or changed by the current export.

    A pre-existing non-empty file is not evidence that ExportData succeeded.
    The baseline comparison also prevents an unrelated .pdf.pdf/.dwg.dwg/
    .dxf.dxf file from being accepted as this request's output.
    """

    unique_candidates: list[Path] = []
    seen: set[str] = set()

    for candidate in candidates:
        key = _normalised_path_key(str(candidate))
        if key in seen:
            continue
        seen.add(key)
        unique_candidates.append(candidate)

    deadline = time.monotonic() + timeout_seconds
    observations: list[dict[str, Any]] = []

    while True:
        observations = []
        for candidate in unique_candidates:
            verification = _file_verification(str(candidate))
            baseline_item = (baseline or {}).get(
                _normalised_path_key(str(candidate))
            )
            changed = bool(
                baseline_item is None
                or not baseline_item.get("exists", False)
                or verification["size_bytes"]
                != baseline_item.get("size_bytes")
                or verification["modified_time_ns"]
                != baseline_item.get("modified_time_ns")
            )
            verification["changed_since_start"] = changed
            observations.append(verification)

        for candidate, verification in zip(
            unique_candidates,
            observations,
        ):
            if (
                verification["verified"]
                and verification["changed_since_start"]
            ):
                return candidate, observations

        if time.monotonic() >= deadline:
            return None, observations

        time.sleep(poll_interval_seconds)


def _export_candidate_paths(
    requested_path: str,
    api_argument: str,
    format_key: str,
) -> list[Path]:
    """Return only deterministic paths CATIA may create for one strategy."""

    requested = Path(requested_path)
    suffix = EXPORT_FORMATS[format_key]
    api_path = Path(api_argument)

    candidates = [requested]
    if api_path.suffix.lower() == suffix:
        # Some V5 releases accept a complete output path; others append the
        # format and produce e.g. drawing.pdf.pdf.
        candidates.append(Path(str(api_path) + suffix))
    else:
        candidates.append(Path(str(api_path) + suffix))
        candidates.append(api_path.with_suffix(suffix))

    unique: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        key = _normalised_path_key(str(candidate))
        if key in seen:
            continue
        seen.add(key)
        unique.append(candidate)
    return unique

def _normalise_export_result_path(
    produced_path: Path,
    requested_path: Path,
) -> dict[str, Any]:
    if (
        _normalised_path_key(str(produced_path))
        == _normalised_path_key(str(requested_path))
    ):
        return {
            "performed": False,
            "source_path": str(produced_path),
            "target_path": str(requested_path),
        }

    if requested_path.exists():
        requested_path.unlink()

    produced_path.replace(requested_path)
    return {
        "performed": True,
        "source_path": str(produced_path),
        "target_path": str(requested_path),
    }


def _verified_save_as(
    drawing_document: Any,
    path: str,
    overwrite: bool,
) -> dict[str, Any]:
    normalized = _prepare_output_path(
        path,
        "path",
        ".CATDrawing",
        overwrite,
    )

    current_full_name = _document_full_name(drawing_document)
    same_current_path = bool(
        current_full_name
        and _normalised_path_key(current_full_name)
        == _normalised_path_key(normalized)
    )

    if same_current_path:
        drawing_document.Save()
        method = "Save"
    else:
        _remove_existing_output(normalized, overwrite)
        drawing_document.SaveAs(normalized)
        method = "SaveAs"

    verification = _file_verification(normalized)
    if not verification["verified"]:
        raise RuntimeError(
            "CATDrawing save call returned without error, but the "
            "output file could not be verified."
        )

    actual_full_name = _document_full_name(drawing_document)
    return {
        "saved": True,
        "method": method,
        "requested_path": normalized,
        "actual_document_full_name": actual_full_name,
        "file": verification,
        "save_verified": bool(
            actual_full_name
            and _normalised_path_key(actual_full_name)
            == _normalised_path_key(normalized)
            and verification["verified"]
        ),
    }



def _export_via_evaluate(
    application: Any,
    drawing_document: Any,
    output_path: str,
    format_name: str,
) -> None:
    script = (
        "Public Function MCP_ExportDrawingDocument("
        "documentObject, outputPath, formatName)\n"
        "    documentObject.ExportData "
        "CStr(outputPath), CStr(formatName)\n"
        "    MCP_ExportDrawingDocument = True\n"
        "End Function"
    )
    _evaluate(
        application,
        script,
        "MCP_ExportDrawingDocument",
        [
            drawing_document,
            output_path,
            format_name,
        ],
    )



def _verified_export(
    application: Any,
    drawing_document: Any,
    path: str,
    format_name: Any,
    overwrite: bool,
) -> dict[str, Any]:
    format_key = _nonempty_text(
        format_name,
        "format_name",
    ).lower()

    if format_key not in EXPORT_FORMATS:
        raise ValueError(
            "format_name must be one of: pdf, dwg, dxf."
        )

    normalized = _prepare_output_path(
        path,
        "path",
        EXPORT_FORMATS[format_key],
        overwrite,
    )
    requested = Path(normalized)

    # overwrite applies only to the path explicitly requested by the caller.
    # Never pre-delete inferred side-effect paths such as drawing.pdf.pdf.
    _remove_existing_output(normalized, overwrite)

    # Keep the v6 strategy order because it was verified in this CATIA
    # environment.  This revision changes only candidate safety/verification.
    strategy_arguments = [
        {
            "strategy": "catia_evaluate_full_path",
            "api_argument": str(requested),
            "format_argument": format_key,
            "call_backend": "SystemService.Evaluate",
        },
        {
            "strategy": "catia_evaluate_extensionless_base_path",
            "api_argument": str(requested.with_suffix("")),
            "format_argument": format_key,
            "call_backend": "SystemService.Evaluate",
        },
        {
            "strategy": "direct_full_path_with_extension",
            "api_argument": str(requested),
            "format_argument": format_key,
            "call_backend": "direct_com",
        },
        {
            "strategy": "direct_extensionless_base_path",
            "api_argument": str(requested.with_suffix("")),
            "format_argument": format_key,
            "call_backend": "direct_com",
        },
        {
            "strategy": "catia_evaluate_uppercase_format",
            "api_argument": str(requested),
            "format_argument": format_key.upper(),
            "call_backend": "SystemService.Evaluate",
        },
    ]

    all_candidates: list[Path] = []
    for strategy in strategy_arguments:
        all_candidates.extend(
            _export_candidate_paths(
                normalized,
                strategy["api_argument"],
                format_key,
            )
        )

    baseline = {
        _normalised_path_key(str(candidate)): _file_verification(
            str(candidate)
        )
        for candidate in all_candidates
    }

    attempts: list[dict[str, Any]] = []
    produced_path: Optional[Path] = None
    selected_strategy: Optional[str] = None
    normalization_action = {
        "performed": False,
        "source_path": None,
        "target_path": str(requested),
    }

    for strategy in strategy_arguments:
        strategy_candidates = _export_candidate_paths(
            normalized,
            strategy["api_argument"],
            format_key,
        )
        attempt: dict[str, Any] = {
            **strategy,
            "call_succeeded": False,
            "error": None,
            "produced_path": None,
            "candidate_observations": [],
            "protected_candidates": [],
        }

        # If this strategy may write to a non-requested candidate that already
        # existed before the call, skip it rather than risk overwriting or
        # deleting unrelated user data.
        protected_candidates = [
            str(candidate)
            for candidate in strategy_candidates
            if (
                _normalised_path_key(str(candidate))
                != _normalised_path_key(normalized)
                and baseline.get(
                    _normalised_path_key(str(candidate)), {}
                ).get("exists", False)
            )
        ]
        if protected_candidates:
            attempt["error"] = (
                "Strategy skipped because it could overwrite an "
                "unrelated pre-existing side-effect path."
            )
            attempt["protected_candidates"] = protected_candidates
            attempts.append(attempt)
            continue

        try:
            if strategy["call_backend"] == "SystemService.Evaluate":
                _export_via_evaluate(
                    application,
                    drawing_document,
                    strategy["api_argument"],
                    strategy["format_argument"],
                )
            else:
                drawing_document.ExportData(
                    strategy["api_argument"],
                    strategy["format_argument"],
                )
            attempt["call_succeeded"] = True
        except Exception as exc:
            attempt["error"] = _format_com_error(exc)
            attempts.append(attempt)
            continue

        found, observations = _wait_for_nonempty_file(
            strategy_candidates,
            timeout_seconds=(
                30.0 if format_key == "pdf" else 15.0
            ),
            baseline=baseline,
        )
        attempt["candidate_observations"] = observations

        if found is not None:
            produced_path = found
            attempt["produced_path"] = str(found)
            selected_strategy = strategy["strategy"]
            attempts.append(attempt)
            break

        attempt["error"] = (
            "ExportData returned without error, but no new or changed "
            "non-empty candidate file appeared within the verification "
            "timeout."
        )
        attempts.append(attempt)

    if produced_path is None:
        raise DraftingOperationError(
            "All ExportData strategies failed to produce a "
            "verifiable output file without risking unrelated files.",
            data={
                "requested_path": normalized,
                "format": format_key,
                "candidate_baseline": list(baseline.values()),
                "export_attempts": attempts,
                "export_verified": False,
                "model_modified": False,
                "document_save_required": False,
            },
        )

    normalization_action = _normalise_export_result_path(
        produced_path,
        requested,
    )
    verification = _file_verification(normalized)

    if not verification["verified"]:
        raise DraftingOperationError(
            "An export candidate was produced, but normalization to "
            "the requested path could not be verified.",
            data={
                "requested_path": normalized,
                "format": format_key,
                "produced_path": str(produced_path),
                "normalization_action": normalization_action,
                "candidate_baseline": list(baseline.values()),
                "export_attempts": attempts,
                "file": verification,
                "export_verified": False,
                "model_modified": False,
                "document_save_required": False,
            },
        )

    return {
        "exported": True,
        "format": format_key,
        "requested_path": normalized,
        "selected_strategy": selected_strategy,
        "api_filename_semantics": (
            "CATIA may append the file_type extension to an "
            "extensionless base path."
        ),
        "produced_path_before_normalization": str(
            produced_path
        ),
        "normalization_action": normalization_action,
        "candidate_baseline": list(baseline.values()),
        "export_attempts": attempts,
        "file": verification,
        "export_verified": True,
        "unrelated_preexisting_candidates_preserved": True,
        "model_modified": False,
        "document_save_required": False,
    }


# ---------------------------------------------------------------------------
# Sheet configuration and layout
# ---------------------------------------------------------------------------

def _set_scale(
    obj: Any,
    value: float,
) -> tuple[str, float]:
    errors: list[str] = []

    for property_name in ("Scale", "Scale2"):
        try:
            setattr(obj, property_name, float(value))
            actual = float(getattr(obj, property_name))
            if abs(actual - value) > 1e-9:
                raise RuntimeError(
                    f"readback={actual}, requested={value}"
                )
            return property_name, actual
        except Exception as exc:
            errors.append(
                f"{property_name}: {_format_com_error(exc)}"
            )

    raise RuntimeError(
        f"Could not set/read scale: {errors}"
    )


def _set_projection_method(
    application: Any,
    sheet: Any,
    method_key: str,
) -> dict[str, Any]:
    script = (
        "Public Function MCP_SetProjectionMethod("
        "sheetObject, methodName)\n"
        "    If LCase(CStr(methodName)) = "
        "\"first_angle\" Then\n"
        "        sheetObject.ProjectionMethod = catFirstAngle\n"
        "    ElseIf LCase(CStr(methodName)) = "
        "\"third_angle\" Then\n"
        "        sheetObject.ProjectionMethod = catThirdAngle\n"
        "    Else\n"
        "        Err.Raise 5, , \"Unsupported projection method\"\n"
        "    End If\n"
        "    MCP_SetProjectionMethod = "
        "CLng(sheetObject.ProjectionMethod)\n"
        "End Function"
    )
    actual_code = int(
        _evaluate(
            application,
            script,
            "MCP_SetProjectionMethod",
            [sheet, PROJECTION_METHODS[method_key]],
        )
    )
    return {
        "requested": PROJECTION_METHODS[method_key],
        "actual_code": actual_code,
        "verified": True,
        "method": "SystemService.Evaluate",
    }


def _standard_paper_dimensions(
    paper_key: str,
    orientation_key: str,
) -> dict[str, float]:
    if paper_key not in STANDARD_PAPER_PORTRAIT_MM:
        raise RuntimeError(
            "Standard paper-dimension fallback is unavailable for "
            f"paper size '{paper_key}'."
        )

    portrait_width, portrait_height = (
        STANDARD_PAPER_PORTRAIT_MM[paper_key]
    )
    if orientation_key == "LANDSCAPE":
        return {
            "width_mm": max(
                portrait_width,
                portrait_height,
            ),
            "height_mm": min(
                portrait_width,
                portrait_height,
            ),
        }
    return {
        "width_mm": min(
            portrait_width,
            portrait_height,
        ),
        "height_mm": max(
            portrait_width,
            portrait_height,
        ),
    }


def _validate_paper_dimensions(
    width: float,
    height: float,
) -> tuple[float, float]:
    width_value = float(width)
    height_value = float(height)

    if (
        not math.isfinite(width_value)
        or not math.isfinite(height_value)
        or width_value <= 0.0
        or height_value <= 0.0
    ):
        raise RuntimeError(
            "Invalid paper dimensions: "
            f"width={width_value}, height={height_value}."
        )
    return width_value, height_value


def _paper_dimensions_direct_return(
    sheet: Any,
) -> tuple[float, float]:
    width = sheet.GetPaperWidth()
    height = sheet.GetPaperHeight()
    return _validate_paper_dimensions(width, height)


def _paper_dimensions_evaluate_return(
    application: Any,
    sheet: Any,
) -> tuple[float, float]:
    script = (
        "Public Function MCP_GetPaperDimensionsReturn(sheetObject)\n"
        "    Dim paperWidth\n"
        "    Dim paperHeight\n"
        "    paperWidth = sheetObject.GetPaperWidth()\n"
        "    paperHeight = sheetObject.GetPaperHeight()\n"
        "    MCP_GetPaperDimensionsReturn = Array("
        "CDbl(paperWidth), CDbl(paperHeight))\n"
        "End Function"
    )
    width, height = _numeric_sequence(
        _evaluate(
            application,
            script,
            "MCP_GetPaperDimensionsReturn",
            [sheet],
        ),
        2,
    )
    return _validate_paper_dimensions(width, height)


def _paper_dimensions_evaluate_property_style(
    application: Any,
    sheet: Any,
) -> tuple[float, float]:
    script = (
        "Public Function MCP_GetPaperDimensionsProperty(sheetObject)\n"
        "    Dim paperWidth\n"
        "    Dim paperHeight\n"
        "    paperWidth = sheetObject.GetPaperWidth\n"
        "    paperHeight = sheetObject.GetPaperHeight\n"
        "    MCP_GetPaperDimensionsProperty = Array("
        "CDbl(paperWidth), CDbl(paperHeight))\n"
        "End Function"
    )
    width, height = _numeric_sequence(
        _evaluate(
            application,
            script,
            "MCP_GetPaperDimensionsProperty",
            [sheet],
        ),
        2,
    )
    return _validate_paper_dimensions(width, height)


def _paper_dimensions_typed_byref(
    sheet: Any,
) -> tuple[float, float]:
    try:
        import pythoncom  # type: ignore
        from win32com.client import VARIANT  # type: ignore
    except Exception as exc:
        raise RuntimeError(
            f"Typed ByRef VARIANT support is unavailable: {exc}"
        ) from exc

    width_value = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_R8,
        0.0,
    )
    height_value = VARIANT(
        pythoncom.VT_BYREF | pythoncom.VT_R8,
        0.0,
    )
    sheet.GetPaperWidth(width_value)
    sheet.GetPaperHeight(height_value)
    return _validate_paper_dimensions(
        width_value.value,
        height_value.value,
    )


def _get_paper_dimensions(
    application: Any,
    sheet: Any,
    paper_key: Optional[str] = None,
    orientation_key: Optional[str] = None,
) -> dict[str, Any]:
    attempts: list[dict[str, Any]] = []
    strategies = (
        (
            "DrawingSheet.GetPaperWidth_return_value",
            lambda: _paper_dimensions_direct_return(sheet),
        ),
        (
            "SystemService.Evaluate_no_argument_return",
            lambda: _paper_dimensions_evaluate_return(
                application,
                sheet,
            ),
        ),
        (
            "SystemService.Evaluate_property_style_return",
            lambda: _paper_dimensions_evaluate_property_style(
                application,
                sheet,
            ),
        ),
        (
            "DrawingSheet.GetPaperWidth_typed_BYREF_VARIANT",
            lambda: _paper_dimensions_typed_byref(sheet),
        ),
    )

    for method_name, strategy in strategies:
        try:
            width, height = strategy()
            attempts.append(
                {
                    "method": method_name,
                    "succeeded": True,
                    "error": None,
                }
            )
            return {
                "width_mm": width,
                "height_mm": height,
                "read_method": method_name,
                "read_attempts": attempts,
                "fallback_used": False,
                "verified": True,
            }
        except Exception as exc:
            attempts.append(
                {
                    "method": method_name,
                    "succeeded": False,
                    "error": _format_com_error(exc),
                }
            )

    actual_paper_code = int(sheet.PaperSize)
    actual_orientation_code = int(sheet.Orientation)
    resolved_paper_key = (
        paper_key
        or PAPER_SIZE_NAME_BY_CODE.get(actual_paper_code)
    )
    resolved_orientation_key = (
        orientation_key
        or ORIENTATION_NAME_BY_CODE.get(
            actual_orientation_code
        )
    )

    if (
        resolved_paper_key is None
        or resolved_orientation_key is None
    ):
        raise RuntimeError(
            "All paper-dimension API strategies failed and the "
            "paper/orientation codes could not be resolved. "
            f"Attempts: {attempts}"
        )

    fallback = _standard_paper_dimensions(
        resolved_paper_key,
        resolved_orientation_key,
    )
    attempts.append(
        {
            "method": (
                "standard_dimensions_from_verified_"
                "PaperSize_and_Orientation"
            ),
            "succeeded": True,
            "error": None,
        }
    )
    return {
        **fallback,
        "read_method": (
            "standard_dimensions_from_verified_"
            "PaperSize_and_Orientation"
        ),
        "read_attempts": attempts,
        "fallback_used": True,
        "fallback_basis": {
            "paper_key": resolved_paper_key,
            "paper_size_code": actual_paper_code,
            "orientation_key": resolved_orientation_key,
            "orientation_code": actual_orientation_code,
        },
        "verified": True,
    }



def _paper_dimension_verification(
    actual_width_mm: float,
    actual_height_mm: float,
    paper_key: str,
    orientation_key: str,
    tolerance_mm: float = 0.5,
) -> dict[str, Any]:
    if paper_key == "USER":
        return {
            "applicable": False,
            "verified": True,
            "reason": (
                "USER paper size has no fixed standard dimensions."
            ),
            "tolerance_mm": tolerance_mm,
        }

    expected = _standard_paper_dimensions(
        paper_key,
        orientation_key,
    )
    width_error = (
        float(actual_width_mm) - expected["width_mm"]
    )
    height_error = (
        float(actual_height_mm) - expected["height_mm"]
    )
    verified = bool(
        abs(width_error) <= tolerance_mm
        and abs(height_error) <= tolerance_mm
    )
    return {
        "applicable": True,
        "verified": verified,
        "expected_width_mm": expected["width_mm"],
        "expected_height_mm": expected["height_mm"],
        "actual_width_mm": float(actual_width_mm),
        "actual_height_mm": float(actual_height_mm),
        "width_error_mm": width_error,
        "height_error_mm": height_error,
        "tolerance_mm": tolerance_mm,
    }



def _configure_sheet(
    application: Any,
    sheet: Any,
    paper_size: Any,
    orientation: Any,
    scale: Any,
    projection_method: Any,
) -> dict[str, Any]:
    paper_key = _normalise_choice(
        paper_size,
        "paper_size",
        PAPER_SIZE_ENUM,
    )
    orientation_key = _normalise_choice(
        orientation,
        "orientation",
        ORIENTATION_ENUM,
    )
    projection_key = _normalise_choice(
        projection_method,
        "projection_method",
        PROJECTION_METHODS,
    )
    scale_value = _positive_float(scale, "scale")

    sheet.PaperSize = PAPER_SIZE_ENUM[paper_key]
    actual_paper_code = int(sheet.PaperSize)
    if actual_paper_code != PAPER_SIZE_ENUM[paper_key]:
        raise RuntimeError(
            "PaperSize readback does not match the requested value."
        )

    sheet.Orientation = ORIENTATION_ENUM[orientation_key]
    actual_orientation_code = int(sheet.Orientation)
    if actual_orientation_code != ORIENTATION_ENUM[
        orientation_key
    ]:
        raise RuntimeError(
            "Orientation readback does not match the requested value."
        )

    scale_method, actual_scale = _set_scale(
        sheet,
        scale_value,
    )
    projection_result = _set_projection_method(
        application,
        sheet,
        projection_key,
    )
    dimensions = _get_paper_dimensions(
        application,
        sheet,
        paper_key=paper_key,
        orientation_key=orientation_key,
    )
    dimension_verification = _paper_dimension_verification(
        dimensions["width_mm"],
        dimensions["height_mm"],
        paper_key,
        orientation_key,
    )
    if not dimension_verification["verified"]:
        raise RuntimeError(
            "Paper dimension readback does not match the requested "
            "standard paper size/orientation. "
            f"Details: {dimension_verification}"
        )

    return {
        "paper_size": {
            "requested": paper_key,
            "actual_code": actual_paper_code,
            "verified": True,
        },
        "orientation": {
            "requested": orientation_key.lower(),
            "actual_code": actual_orientation_code,
            "verified": True,
        },
        "scale": {
            "requested": scale_value,
            "actual": actual_scale,
            "set_method": scale_method,
            "verified": True,
        },
        "projection_method": projection_result,
        "paper_dimensions": dimensions,
        "paper_dimension_verification": (
            dimension_verification
        ),
        "configuration_verified": bool(
            dimension_verification["verified"]
        ),
    }


def _effective_safe_margin_mm(
    requested_sheet_margin_mm: float,
    inner_frame_inset_mm: float = INNER_FRAME_INSET_MM,
) -> float:
    requested = _positive_float(
        requested_sheet_margin_mm, "requested_sheet_margin_mm"
    )
    inset = _finite_float(inner_frame_inset_mm, "inner_frame_inset_mm")
    if inset < 0.0:
        raise ValueError("inner_frame_inset_mm cannot be negative.")
    return requested + inset


def _safe_area_descriptor(
    paper_width_mm: float,
    paper_height_mm: float,
    requested_sheet_margin_mm: float,
    inner_frame_inset_mm: float = INNER_FRAME_INSET_MM,
) -> dict[str, Any]:
    effective = _effective_safe_margin_mm(
        requested_sheet_margin_mm, inner_frame_inset_mm
    )
    xmin = effective
    ymin = effective
    xmax = float(paper_width_mm) - effective
    ymax = float(paper_height_mm) - effective
    if xmax <= xmin or ymax <= ymin:
        raise DraftingOperationError(
            "The requested sheet margin plus inner-frame inset leaves no usable drawing area.",
            data={
                "requested_sheet_margin_mm": float(requested_sheet_margin_mm),
                "inner_frame_inset_mm": float(inner_frame_inset_mm),
                "effective_safe_margin_mm": effective,
                "paper_width_mm": float(paper_width_mm),
                "paper_height_mm": float(paper_height_mm),
            },
        )
    return {
        "requested_sheet_margin_mm": float(requested_sheet_margin_mm),
        "inner_frame_inset_mm": float(inner_frame_inset_mm),
        "effective_safe_margin_mm": effective,
        "xmin_mm": xmin,
        "xmax_mm": xmax,
        "ymin_mm": ymin,
        "ymax_mm": ymax,
        "width_mm": xmax - xmin,
        "height_mm": ymax - ymin,
        "policy": "requested_sheet_margin_plus_fixed_inner_frame_inset",
    }


def _calculate_layout(
    paper_width: float,
    paper_height: float,
    projection_method: str,
    safe_margin_mm: float = 0.0,
) -> DrawingLayout:
    margin = max(0.0, float(safe_margin_mm))
    usable_width = float(paper_width) - 2.0 * margin
    usable_height = float(paper_height) - 2.0 * margin
    if usable_width <= 0.0 or usable_height <= 0.0:
        raise DraftingOperationError(
            "The effective safe margin leaves no usable area for initial layout."
        )

    def ux(fraction: float) -> float:
        return margin + usable_width * fraction

    def uy(fraction: float) -> float:
        return margin + usable_height * fraction

    # Keep the original relative layout proportions, but apply them to the safe area
    # instead of to the full paper. Final DrawingView.Size verification still decides
    # whether the actual generated geometry fits.
    front_x = ux(0.36)
    front_y = uy(0.46)

    if projection_method == "third_angle":
        top_x = front_x
        top_y = uy(0.76)
        right_x = ux(0.70)
        right_y = front_y
    else:
        top_x = front_x
        top_y = uy(0.18)
        right_x = ux(0.12)
        right_y = front_y

    return DrawingLayout(
        front_x=front_x,
        front_y=front_y,
        top_x=top_x,
        top_y=top_y,
        right_x=right_x,
        right_y=right_y,
        title_block_x=max(margin, paper_width - margin - TITLE_BLOCK_WIDTH_MM),
        title_block_y=max(margin, min(DEFAULT_TITLE_BLOCK_Y_MM, paper_height - margin - TITLE_BLOCK_HEIGHT_MM)),
        notes_x=margin + 2.0,
        notes_y=margin + 2.0,
    )


# ---------------------------------------------------------------------------
# Drawing view inspection and creation
# ---------------------------------------------------------------------------

def _background_view(sheet: Any) -> tuple[Any, dict[str, Any]]:
    views = sheet.Views
    count = int(views.Count)
    if count < 2:
        raise RuntimeError(
            "The active sheet does not contain CATIA's main and "
            "background system views."
        )

    view = views.Item(2)
    view_type = int(_safe_attribute(view, "ViewType", -1))
    return view, {
        "selection_method": (
            "DrawingSheet.Views.Item(2)_official_system_order"
        ),
        "index": 2,
        "name": safe_str(_safe_attribute(view, "Name", "")),
        "view_type_code": view_type,
        "view_type_name": DRAWING_VIEW_TYPE_NAMES.get(
            view_type,
            "unknown",
        ),
    }


def _view_scale(view: Any) -> tuple[Optional[float], Optional[str]]:
    errors: list[str] = []
    for property_name in ("Scale", "Scale2"):
        try:
            return float(getattr(view, property_name)), property_name
        except Exception as exc:
            errors.append(
                f"{property_name}: {_format_com_error(exc)}"
            )
    return None, "; ".join(errors)


def _view_size(
    application: Any,
    view: Any,
) -> dict[str, Any]:
    script = (
        "Public Function MCP_GetDrawingViewSize(viewObject)\n"
        "    Dim values(3)\n"
        "    viewObject.Size values\n"
        "    MCP_GetDrawingViewSize = Array("
        "CDbl(values(0)), CDbl(values(1)), "
        "CDbl(values(2)), CDbl(values(3)))\n"
        "End Function"
    )
    values = _numeric_sequence(
        _evaluate(
            application,
            script,
            "MCP_GetDrawingViewSize",
            [view],
        ),
        4,
    )
    xmin, xmax, ymin, ymax = values
    return {
        "xmin": xmin,
        "xmax": xmax,
        "ymin": ymin,
        "ymax": ymax,
        "width_mm": xmax - xmin,
        "height_mm": ymax - ymin,
        "read_method": "SystemService.Evaluate.DrawingView.Size",
    }


def _projection_plane(
    application: Any,
    behavior: Any,
) -> dict[str, Any]:
    script = (
        "Public Function MCP_GetProjectionPlane(behaviorObject)\n"
        "    Dim x1, y1, z1, x2, y2, z2\n"
        "    Dim nx, ny, nz\n"
        "    behaviorObject.GetProjectionPlane "
        "x1, y1, z1, x2, y2, z2\n"
        "    behaviorObject.GetProjectionPlaneNormal nx, ny, nz\n"
        "    MCP_GetProjectionPlane = Array("
        "CDbl(x1), CDbl(y1), CDbl(z1), "
        "CDbl(x2), CDbl(y2), CDbl(z2), "
        "CDbl(nx), CDbl(ny), CDbl(nz))\n"
        "End Function"
    )
    values = _numeric_sequence(
        _evaluate(
            application,
            script,
            "MCP_GetProjectionPlane",
            [behavior],
        ),
        9,
    )
    return {
        "vector_1": values[0:3],
        "vector_2": values[3:6],
        "normal": values[6:9],
        "read_method": (
            "SystemService.Evaluate."
            "GetProjectionPlane_and_Normal"
        ),
    }



def _view_is_generative(
    application: Any,
    view: Any,
) -> tuple[bool, str, list[dict[str, Any]]]:
    attempts: list[dict[str, Any]] = []

    try:
        value = bool(view.IsGenerative())
        attempts.append(
            {
                "method": "DrawingView.IsGenerative_return_value",
                "succeeded": True,
                "value": value,
                "error": None,
            }
        )
        if value:
            return value, attempts[-1]["method"], attempts
    except Exception as exc:
        attempts.append(
            {
                "method": "DrawingView.IsGenerative_return_value",
                "succeeded": False,
                "value": None,
                "error": _format_com_error(exc),
            }
        )

    script = (
        "Public Function MCP_IsDrawingViewGenerative(viewObject)\n"
        "    MCP_IsDrawingViewGenerative = "
        "CBool(viewObject.IsGenerative())\n"
        "End Function"
    )
    try:
        value = bool(
            _evaluate(
                application,
                script,
                "MCP_IsDrawingViewGenerative",
                [view],
            )
        )
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate."
                    "DrawingView.IsGenerative"
                ),
                "succeeded": True,
                "value": value,
                "error": None,
            }
        )
        return value, attempts[-1]["method"], attempts
    except Exception as exc:
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate."
                    "DrawingView.IsGenerative"
                ),
                "succeeded": False,
                "value": None,
                "error": _format_com_error(exc),
            }
        )

    return False, "all_methods_failed", attempts


def _refresh_drafting_view(
    application: Any,
    drawing_document: Any,
    sheet: Any,
    view: Any,
) -> list[dict[str, Any]]:
    attempts: list[dict[str, Any]] = []

    actions = (
        ("drawing_document.Activate", lambda: drawing_document.Activate()),
        ("sheet.Activate", lambda: sheet.Activate()),
        ("view.Activate", lambda: view.Activate()),
        (
            "view.GenerativeBehavior.ForceUpdate",
            lambda: view.GenerativeBehavior.ForceUpdate(),
        ),
        ("sheet.ForceUpdate", lambda: sheet.ForceUpdate()),
        ("drawing_document.Update", lambda: drawing_document.Update()),
        (
            "application.ActiveWindow.ActiveViewer.Update",
            lambda: application.ActiveWindow.ActiveViewer.Update(),
        ),
    )

    for name, action in actions:
        try:
            action()
            attempts.append(
                {
                    "method": name,
                    "succeeded": True,
                    "error": None,
                }
            )
        except Exception as exc:
            attempts.append(
                {
                    "method": name,
                    "succeeded": False,
                    "error": _format_com_error(exc),
                }
            )

    try:
        application.RefreshDisplay = True
        attempts.append(
            {
                "method": "application.RefreshDisplay=true",
                "succeeded": True,
                "error": None,
            }
        )
    except Exception as exc:
        attempts.append(
            {
                "method": "application.RefreshDisplay=true",
                "succeeded": False,
                "error": _format_com_error(exc),
            }
        )

    return attempts


def _view_generation_snapshot(
    application: Any,
    view: Any,
) -> dict[str, Any]:
    generative, generative_method, generative_attempts = (
        _view_is_generative(application, view)
    )

    size: Optional[dict[str, Any]] = None
    size_error: Optional[str] = None
    try:
        size = _view_size(application, view)
    except Exception as exc:
        size_error = str(exc)

    geometry_nonempty = bool(
        size is not None
        and abs(size["width_mm"]) > 1.01
        and abs(size["height_mm"]) > 1.01
    )

    return {
        "is_generative": generative,
        "is_generative_read_method": generative_method,
        "is_generative_read_attempts": generative_attempts,
        "bounding_box": size,
        "bounding_box_error": size_error,
        "geometry_nonempty": geometry_nonempty,
        "verified": bool(
            generative and geometry_nonempty
        ),
    }



def _set_drawing_view_position_via_evaluate(
    application: Any,
    view: Any,
    target_x: float,
    target_y: float,
) -> tuple[float, float]:
    script = (
        "Public Function MCP_SetDrawingViewPosition("
        "viewObject, targetX, targetY)\n"
        "    On Error Resume Next\n"
        "    viewObject.UnAlignedWithReferenceView\n"
        "    Err.Clear\n"
        "    On Error GoTo 0\n"
        "    viewObject.x = CDbl(targetX)\n"
        "    viewObject.y = CDbl(targetY)\n"
        "    MCP_SetDrawingViewPosition = Array("
        "CDbl(viewObject.x), CDbl(viewObject.y))\n"
        "End Function"
    )
    actual_x, actual_y = _numeric_sequence(
        _evaluate(
            application,
            script,
            "MCP_SetDrawingViewPosition",
            [view, float(target_x), float(target_y)],
        ),
        2,
    )
    return actual_x, actual_y


def _position_readback(
    view: Any,
) -> dict[str, float]:
    return {
        "x_mm": float(_safe_attribute(view, "x", 0.0)),
        "y_mm": float(_safe_attribute(view, "y", 0.0)),
    }


def _position_matches(
    actual_x: float,
    actual_y: float,
    target_x: float,
    target_y: float,
    tolerance_mm: float,
) -> bool:
    return bool(
        abs(float(actual_x) - float(target_x)) <= tolerance_mm
        and abs(float(actual_y) - float(target_y)) <= tolerance_mm
    )



def _position_projected_view(
    application: Any,
    view: Any,
    front_view: Any,
    projection_type: str,
    target_x: float,
    target_y: float,
    position_tolerance_mm: float = 0.01,
) -> dict[str, Any]:
    attempts: list[dict[str, Any]] = []

    try:
        view.ReferenceView = front_view
        attempts.append(
            {
                "method": "ReferenceView=front_view",
                "succeeded": True,
                "error": None,
            }
        )
    except Exception as exc:
        attempts.append(
            {
                "method": "ReferenceView=front_view",
                "succeeded": False,
                "error": _format_com_error(exc),
            }
        )

    if projection_type == "top":
        requested_x = float(front_view.x)
        requested_y = float(target_y)
        manual_alignment_axis = "x"
    else:
        requested_x = float(target_x)
        requested_y = float(front_view.y)
        manual_alignment_axis = "y"

    try:
        view.UnAlignedWithReferenceView()
        attempts.append(
            {
                "method": "UnAlignedWithReferenceView",
                "succeeded": True,
                "error": None,
            }
        )
    except Exception as exc:
        attempts.append(
            {
                "method": "UnAlignedWithReferenceView",
                "succeeded": False,
                "error": _format_com_error(exc),
            }
        )

    selected_strategy: Optional[str] = None
    position_verified = False

    # Strategy 1: execute the position write inside CATIA.
    try:
        actual_x, actual_y = (
            _set_drawing_view_position_via_evaluate(
                application,
                view,
                requested_x,
                requested_y,
            )
        )
        position_verified = _position_matches(
            actual_x,
            actual_y,
            requested_x,
            requested_y,
            position_tolerance_mm,
        )
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate."
                    "MCP_SetDrawingViewPosition"
                ),
                "succeeded": True,
                "position_verified": position_verified,
                "actual_x_mm": actual_x,
                "actual_y_mm": actual_y,
                "error": None,
            }
        )
        if position_verified:
            selected_strategy = (
                "SystemService.Evaluate."
                "MCP_SetDrawingViewPosition"
            )
    except Exception as exc:
        attempts.append(
            {
                "method": (
                    "SystemService.Evaluate."
                    "MCP_SetDrawingViewPosition"
                ),
                "succeeded": False,
                "position_verified": False,
                "error": _format_com_error(exc),
            }
        )

    # Strategy 2: direct x/y properties after unalignment.
    if not position_verified:
        try:
            view.UnAlignedWithReferenceView()
        except Exception:
            pass

        try:
            view.x = requested_x
            view.y = requested_y
            readback = _position_readback(view)
            position_verified = _position_matches(
                readback["x_mm"],
                readback["y_mm"],
                requested_x,
                requested_y,
                position_tolerance_mm,
            )
            attempts.append(
                {
                    "method": (
                        "DrawingView.x_y_after_unalignment"
                    ),
                    "succeeded": True,
                    "position_verified": position_verified,
                    "actual_x_mm": readback["x_mm"],
                    "actual_y_mm": readback["y_mm"],
                    "error": None,
                }
            )
            if position_verified:
                selected_strategy = (
                    "DrawingView.x_y_after_unalignment"
                )
        except Exception as exc:
            attempts.append(
                {
                    "method": (
                        "DrawingView.x_y_after_unalignment"
                    ),
                    "succeeded": False,
                    "position_verified": False,
                    "error": _format_com_error(exc),
                }
            )

    # Strategy 3: xAxisData/yAxisData compatibility fallback.
    if not position_verified:
        try:
            view.UnAlignedWithReferenceView()
        except Exception:
            pass

        try:
            view.xAxisData = requested_x
            view.yAxisData = requested_y
            actual_x = float(
                _safe_attribute(
                    view,
                    "xAxisData",
                    _safe_attribute(view, "x", 0.0),
                )
            )
            actual_y = float(
                _safe_attribute(
                    view,
                    "yAxisData",
                    _safe_attribute(view, "y", 0.0),
                )
            )
            position_verified = _position_matches(
                actual_x,
                actual_y,
                requested_x,
                requested_y,
                position_tolerance_mm,
            )
            attempts.append(
                {
                    "method": (
                        "DrawingView.xAxisData_yAxisData_"
                        "after_unalignment"
                    ),
                    "succeeded": True,
                    "position_verified": position_verified,
                    "actual_x_mm": actual_x,
                    "actual_y_mm": actual_y,
                    "error": None,
                }
            )
            if position_verified:
                selected_strategy = (
                    "DrawingView.xAxisData_yAxisData_"
                    "after_unalignment"
                )
        except Exception as exc:
            attempts.append(
                {
                    "method": (
                        "DrawingView.xAxisData_yAxisData_"
                        "after_unalignment"
                    ),
                    "succeeded": False,
                    "position_verified": False,
                    "error": _format_com_error(exc),
                }
            )

    final_readback = _position_readback(view)
    final_position_verified = _position_matches(
        final_readback["x_mm"],
        final_readback["y_mm"],
        requested_x,
        requested_y,
        position_tolerance_mm,
    )
    position_verified = bool(
        position_verified and final_position_verified
    )

    reference_view_name: Optional[str] = None
    try:
        reference_view_name = safe_str(view.ReferenceView.Name)
    except Exception:
        pass

    return {
        "projection_type": projection_type,
        "alignment_mode": (
            "reference_link_preserved_manual_axis_alignment"
        ),
        "manual_alignment_axis": manual_alignment_axis,
        "requested_x_mm": requested_x,
        "requested_y_mm": requested_y,
        "actual_x_mm": final_readback["x_mm"],
        "actual_y_mm": final_readback["y_mm"],
        "position_tolerance_mm": position_tolerance_mm,
        "position_verified": position_verified,
        "selected_strategy": selected_strategy,
        "reference_view_name": reference_view_name,
        "reference_link_preserved": bool(
            reference_view_name
            == safe_str(_safe_attribute(front_view, "Name", ""))
        ),
        "attempts": attempts,
    }


def _wait_for_view_generation(
    application: Any,
    drawing_document: Any,
    sheet: Any,
    view: Any,
    *,
    timeout_seconds: float = 20.0,
    poll_interval_seconds: float = 0.25,
    position_callback: Optional[Any] = None,
) -> dict[str, Any]:
    deadline = time.monotonic() + timeout_seconds
    polls: list[dict[str, Any]] = []
    selected_snapshot: Optional[dict[str, Any]] = None

    while True:
        position_result = (
            position_callback()
            if position_callback is not None
            else None
        )
        refresh_attempts = _refresh_drafting_view(
            application,
            drawing_document,
            sheet,
            view,
        )
        time.sleep(poll_interval_seconds)
        snapshot = _view_generation_snapshot(
            application,
            view,
        )
        polls.append(
            {
                "elapsed_seconds": (
                    timeout_seconds
                    - max(0.0, deadline - time.monotonic())
                ),
                "position_result": position_result,
                "refresh_attempts": refresh_attempts,
                "snapshot": snapshot,
            }
        )

        if snapshot["verified"]:
            selected_snapshot = snapshot
            break

        if time.monotonic() >= deadline:
            selected_snapshot = snapshot
            break

    return {
        "verified": bool(
            selected_snapshot
            and selected_snapshot["verified"]
        ),
        "timeout_seconds": timeout_seconds,
        "poll_interval_seconds": poll_interval_seconds,
        "poll_count": len(polls),
        "final_snapshot": selected_snapshot,
        "polls": polls,
    }




def _rectangle_from_view_size(
    size: dict[str, Any],
) -> dict[str, float]:
    return {
        "xmin": float(size["xmin"]),
        "xmax": float(size["xmax"]),
        "ymin": float(size["ymin"]),
        "ymax": float(size["ymax"]),
        "width_mm": float(size["width_mm"]),
        "height_mm": float(size["height_mm"]),
    }


def _rectangle_overlap(
    first: dict[str, float],
    second: dict[str, float],
    clearance_mm: float = 0.0,
    tolerance_mm: float = 0.01,
) -> dict[str, Any]:
    overlap_x = min(
        first["xmax"],
        second["xmax"],
    ) - max(
        first["xmin"],
        second["xmin"],
    )
    overlap_y = min(
        first["ymax"],
        second["ymax"],
    ) - max(
        first["ymin"],
        second["ymin"],
    )
    effective_clearance = max(
        0.0,
        float(clearance_mm) - float(tolerance_mm),
    )
    intersects = bool(
        overlap_x > -effective_clearance
        and overlap_y > -effective_clearance
    )
    geometric_overlap = bool(
        overlap_x > 0.0
        and overlap_y > 0.0
    )
    return {
        "intersects_with_clearance": intersects,
        "geometric_overlap": geometric_overlap,
        "overlap_x_mm": max(0.0, overlap_x),
        "overlap_y_mm": max(0.0, overlap_y),
        "signed_x_separation_mm": -overlap_x,
        "signed_y_separation_mm": -overlap_y,
        "required_clearance_mm": float(clearance_mm),
        "tolerance_mm": float(tolerance_mm),
        "effective_clearance_mm": effective_clearance,
    }


def _sheet_bounds_verification(
    rectangles: dict[str, dict[str, float]],
    paper_width_mm: float,
    paper_height_mm: float,
    margin_mm: float,
) -> dict[str, Any]:
    per_view: dict[str, Any] = {}
    all_within = True

    for name, rectangle in rectangles.items():
        left_clearance = rectangle["xmin"]
        right_clearance = (
            paper_width_mm - rectangle["xmax"]
        )
        bottom_clearance = rectangle["ymin"]
        top_clearance = (
            paper_height_mm - rectangle["ymax"]
        )
        within = bool(
            left_clearance >= margin_mm
            and right_clearance >= margin_mm
            and bottom_clearance >= margin_mm
            and top_clearance >= margin_mm
        )
        all_within = all_within and within
        per_view[name] = {
            "within_sheet_margin": within,
            "left_clearance_mm": left_clearance,
            "right_clearance_mm": right_clearance,
            "bottom_clearance_mm": bottom_clearance,
            "top_clearance_mm": top_clearance,
        }

    return {
        "margin_mm": float(margin_mm),
        "paper_width_mm": float(paper_width_mm),
        "paper_height_mm": float(paper_height_mm),
        "per_view": per_view,
        "all_views_within_sheet_margin": all_within,
    }



def _group_bounds(
    rectangles: dict[str, dict[str, float]],
) -> dict[str, float]:
    if not rectangles:
        raise ValueError(
            "rectangles cannot be empty when calculating group bounds."
        )

    xmin = min(item["xmin"] for item in rectangles.values())
    xmax = max(item["xmax"] for item in rectangles.values())
    ymin = min(item["ymin"] for item in rectangles.values())
    ymax = max(item["ymax"] for item in rectangles.values())

    return {
        "xmin": float(xmin),
        "xmax": float(xmax),
        "ymin": float(ymin),
        "ymax": float(ymax),
        "width_mm": float(xmax - xmin),
        "height_mm": float(ymax - ymin),
        "center_x_mm": float((xmin + xmax) / 2.0),
        "center_y_mm": float((ymin + ymax) / 2.0),
    }


def _group_fit_verification(
    group: dict[str, float],
    paper_width_mm: float,
    paper_height_mm: float,
    sheet_margin_mm: float,
    tolerance_mm: float = 0.01,
) -> dict[str, Any]:
    usable_width = (
        float(paper_width_mm)
        - 2.0 * float(sheet_margin_mm)
    )
    usable_height = (
        float(paper_height_mm)
        - 2.0 * float(sheet_margin_mm)
    )
    width_excess = group["width_mm"] - usable_width
    height_excess = group["height_mm"] - usable_height
    fits_width = bool(width_excess <= tolerance_mm)
    fits_height = bool(height_excess <= tolerance_mm)

    return {
        "group_width_mm": group["width_mm"],
        "group_height_mm": group["height_mm"],
        "usable_width_mm": usable_width,
        "usable_height_mm": usable_height,
        "width_excess_mm": width_excess,
        "height_excess_mm": height_excess,
        "fits_width": fits_width,
        "fits_height": fits_height,
        "group_fits_usable_area": bool(
            fits_width and fits_height
        ),
        "tolerance_mm": float(tolerance_mm),
    }


def _minimum_group_translation_to_margins(
    group: dict[str, float],
    paper_width_mm: float,
    paper_height_mm: float,
    sheet_margin_mm: float,
    tolerance_mm: float = 0.01,
) -> dict[str, Any]:
    fit = _group_fit_verification(
        group,
        paper_width_mm,
        paper_height_mm,
        sheet_margin_mm,
        tolerance_mm,
    )
    if not fit["group_fits_usable_area"]:
        return {
            "possible": False,
            "required_dx_mm": 0.0,
            "required_dy_mm": 0.0,
            "translation_required": False,
            "policy": "minimum_translation_to_sheet_margins",
            "fit_verification": fit,
            "reason": (
                "The three-view union bounding box is larger than "
                "the usable sheet area."
            ),
        }

    left_limit = float(sheet_margin_mm)
    right_limit = float(paper_width_mm) - float(sheet_margin_mm)
    bottom_limit = float(sheet_margin_mm)
    top_limit = float(paper_height_mm) - float(sheet_margin_mm)

    if group["xmin"] < left_limit - tolerance_mm:
        dx = left_limit - group["xmin"]
    elif group["xmax"] > right_limit + tolerance_mm:
        dx = right_limit - group["xmax"]
    else:
        dx = 0.0

    if group["ymin"] < bottom_limit - tolerance_mm:
        dy = bottom_limit - group["ymin"]
    elif group["ymax"] > top_limit + tolerance_mm:
        dy = top_limit - group["ymax"]
    else:
        dy = 0.0

    return {
        "possible": True,
        "required_dx_mm": float(dx),
        "required_dy_mm": float(dy),
        "translation_required": bool(
            abs(dx) > tolerance_mm
            or abs(dy) > tolerance_mm
        ),
        "policy": "minimum_translation_to_sheet_margins",
        "fit_verification": fit,
        "target_group_bounds": {
            "xmin": group["xmin"] + dx,
            "xmax": group["xmax"] + dx,
            "ymin": group["ymin"] + dy,
            "ymax": group["ymax"] + dy,
        },
    }


def _position_independent_view(
    application: Any,
    view: Any,
    target_x: float,
    target_y: float,
    position_tolerance_mm: float = 0.01,
) -> dict[str, Any]:
    attempts: list[dict[str, Any]] = []
    requested_x = float(target_x)
    requested_y = float(target_y)
    selected_strategy: Optional[str] = None
    position_verified = False

    try:
        actual_x, actual_y = _set_drawing_view_position_via_evaluate(
            application,
            view,
            requested_x,
            requested_y,
        )
        position_verified = _position_matches(
            actual_x,
            actual_y,
            requested_x,
            requested_y,
            position_tolerance_mm,
        )
        attempts.append({
            "method": "SystemService.Evaluate.MCP_SetDrawingViewPosition",
            "succeeded": True,
            "position_verified": position_verified,
            "actual_x_mm": actual_x,
            "actual_y_mm": actual_y,
            "error": None,
        })
        if position_verified:
            selected_strategy = "SystemService.Evaluate.MCP_SetDrawingViewPosition"
    except Exception as exc:
        attempts.append({
            "method": "SystemService.Evaluate.MCP_SetDrawingViewPosition",
            "succeeded": False,
            "position_verified": False,
            "error": _format_com_error(exc),
        })

    if not position_verified:
        try:
            view.x = requested_x
            view.y = requested_y
            readback = _position_readback(view)
            position_verified = _position_matches(
                readback["x_mm"],
                readback["y_mm"],
                requested_x,
                requested_y,
                position_tolerance_mm,
            )
            attempts.append({
                "method": "DrawingView.x_y",
                "succeeded": True,
                "position_verified": position_verified,
                "actual_x_mm": readback["x_mm"],
                "actual_y_mm": readback["y_mm"],
                "error": None,
            })
            if position_verified:
                selected_strategy = "DrawingView.x_y"
        except Exception as exc:
            attempts.append({
                "method": "DrawingView.x_y",
                "succeeded": False,
                "position_verified": False,
                "error": _format_com_error(exc),
            })

    if not position_verified:
        try:
            view.xAxisData = requested_x
            view.yAxisData = requested_y
            actual_x = float(_safe_attribute(
                view,
                "xAxisData",
                _safe_attribute(view, "x", 0.0),
            ))
            actual_y = float(_safe_attribute(
                view,
                "yAxisData",
                _safe_attribute(view, "y", 0.0),
            ))
            position_verified = _position_matches(
                actual_x,
                actual_y,
                requested_x,
                requested_y,
                position_tolerance_mm,
            )
            attempts.append({
                "method": "DrawingView.xAxisData_yAxisData",
                "succeeded": True,
                "position_verified": position_verified,
                "actual_x_mm": actual_x,
                "actual_y_mm": actual_y,
                "error": None,
            })
            if position_verified:
                selected_strategy = "DrawingView.xAxisData_yAxisData"
        except Exception as exc:
            attempts.append({
                "method": "DrawingView.xAxisData_yAxisData",
                "succeeded": False,
                "position_verified": False,
                "error": _format_com_error(exc),
            })

    final_readback = _position_readback(view)
    final_verified = _position_matches(
        final_readback["x_mm"],
        final_readback["y_mm"],
        requested_x,
        requested_y,
        position_tolerance_mm,
    )
    return {
        "view_name": safe_str(_safe_attribute(view, "Name", "")),
        "requested_x_mm": requested_x,
        "requested_y_mm": requested_y,
        "actual_x_mm": final_readback["x_mm"],
        "actual_y_mm": final_readback["y_mm"],
        "position_tolerance_mm": position_tolerance_mm,
        "position_verified": bool(position_verified and final_verified),
        "selected_strategy": selected_strategy,
        "attempts": attempts,
    }


def _relative_view_offsets(
    front_view: Any,
    top_view: Any,
    right_view: Any,
) -> dict[str, dict[str, float]]:
    front = _position_readback(front_view)
    top = _position_readback(top_view)
    right = _position_readback(right_view)
    return {
        "top_minus_front": {
            "dx_mm": top["x_mm"] - front["x_mm"],
            "dy_mm": top["y_mm"] - front["y_mm"],
        },
        "right_minus_front": {
            "dx_mm": right["x_mm"] - front["x_mm"],
            "dy_mm": right["y_mm"] - front["y_mm"],
        },
    }


def _relative_offsets_match(
    before: dict[str, dict[str, float]],
    after: dict[str, dict[str, float]],
    tolerance_mm: float = 0.01,
) -> dict[str, Any]:
    errors: dict[str, dict[str, float]] = {}
    verified = True
    for relation in ("top_minus_front", "right_minus_front"):
        errors[relation] = {}
        for axis in ("dx_mm", "dy_mm"):
            error = after[relation][axis] - before[relation][axis]
            errors[relation][axis] = error
            if abs(error) > tolerance_mm:
                verified = False
    return {
        "verified": verified,
        "tolerance_mm": float(tolerance_mm),
        "errors": errors,
    }


def _translate_three_view_group(
    application: Any,
    drawing_document: Any,
    sheet: Any,
    front_view: Any,
    top_view: Any,
    right_view: Any,
    dx_mm: float,
    dy_mm: float,
    *,
    position_tolerance_mm: float = 0.01,
) -> dict[str, Any]:
    dx = float(dx_mm)
    dy = float(dy_mm)
    before_positions = {
        "front": _position_readback(front_view),
        "top": _position_readback(top_view),
        "right": _position_readback(right_view),
    }
    before_offsets = _relative_view_offsets(front_view, top_view, right_view)
    targets = {
        name: {
            "x_mm": position["x_mm"] + dx,
            "y_mm": position["y_mm"] + dy,
        }
        for name, position in before_positions.items()
    }

    front_result = _position_independent_view(
        application,
        front_view,
        targets["front"]["x_mm"],
        targets["front"]["y_mm"],
        position_tolerance_mm,
    )
    top_result = _position_projected_view(
        application,
        top_view,
        front_view,
        "top",
        targets["top"]["x_mm"],
        targets["top"]["y_mm"],
        position_tolerance_mm,
    )
    right_result = _position_projected_view(
        application,
        right_view,
        front_view,
        "right",
        targets["right"]["x_mm"],
        targets["right"]["y_mm"],
        position_tolerance_mm,
    )

    refresh_attempts: list[dict[str, Any]] = []
    for item in (front_view, top_view, right_view):
        refresh_attempts.extend(_refresh_drafting_view(
            application,
            drawing_document,
            sheet,
            item,
        ))
    time.sleep(0.20)

    after_positions = {
        "front": _position_readback(front_view),
        "top": _position_readback(top_view),
        "right": _position_readback(right_view),
    }
    after_offsets = _relative_view_offsets(front_view, top_view, right_view)
    relative_offsets = _relative_offsets_match(
        before_offsets,
        after_offsets,
        position_tolerance_mm,
    )
    position_results = {
        "front": front_result,
        "top": top_result,
        "right": right_result,
    }
    all_positions_verified = all(
        item["position_verified"]
        for item in position_results.values()
    )
    reference_links_preserved = bool(
        top_result["reference_link_preserved"]
        and right_result["reference_link_preserved"]
    )

    return {
        "attempted": True,
        "translation_dx_mm": dx,
        "translation_dy_mm": dy,
        "position_tolerance_mm": position_tolerance_mm,
        "before_positions": before_positions,
        "targets": targets,
        "position_results": position_results,
        "after_positions": after_positions,
        "before_relative_offsets": before_offsets,
        "after_relative_offsets": after_offsets,
        "relative_offsets_verification": relative_offsets,
        "all_positions_verified": all_positions_verified,
        "reference_links_preserved": reference_links_preserved,
        "refresh_attempts": refresh_attempts,
        "verified": bool(
            all_positions_verified
            and reference_links_preserved
            and relative_offsets["verified"]
        ),
    }


def _layout_snapshot(
    application: Any,
    front_view: Any,
    top_view: Any,
    right_view: Any,
    projection_method: str,
    minimum_gap_mm: float,
    paper_width_mm: float,
    paper_height_mm: float,
    sheet_margin_mm: float,
) -> dict[str, Any]:
    views = {
        "front": front_view,
        "top": top_view,
        "right": right_view,
    }
    rectangles: dict[str, dict[str, float]] = {}
    errors: dict[str, Optional[str]] = {}

    for name, view in views.items():
        try:
            rectangles[name] = _rectangle_from_view_size(
                _view_size(application, view)
            )
            errors[name] = None
        except Exception as exc:
            errors[name] = str(exc)

    if len(rectangles) != 3:
        return {
            "verified": False,
            "rectangles": rectangles,
            "read_errors": errors,
            "minimum_gap_mm": float(minimum_gap_mm),
            "error": (
                "Could not read all three view bounding boxes."
            ),
        }

    front = rectangles["front"]
    top = rectangles["top"]
    right = rectangles["right"]

    layout_tolerance_mm = 0.01
    if projection_method == "third_angle":
        front_top_gap = top["ymin"] - front["ymax"]
        front_right_gap = right["xmin"] - front["xmax"]
        top_direction_verified = bool(
            front_top_gap
            >= minimum_gap_mm - layout_tolerance_mm
        )
        right_direction_verified = bool(
            front_right_gap
            >= minimum_gap_mm - layout_tolerance_mm
        )
    else:
        front_top_gap = front["ymin"] - top["ymax"]
        front_right_gap = front["xmin"] - right["xmax"]
        top_direction_verified = bool(
            front_top_gap
            >= minimum_gap_mm - layout_tolerance_mm
        )
        right_direction_verified = bool(
            front_right_gap
            >= minimum_gap_mm - layout_tolerance_mm
        )

    pairwise = {
        "front_top": _rectangle_overlap(
            front,
            top,
            minimum_gap_mm,
        ),
        "front_right": _rectangle_overlap(
            front,
            right,
            minimum_gap_mm,
        ),
        "top_right": _rectangle_overlap(
            top,
            right,
            minimum_gap_mm,
        ),
    }
    pairwise_clear = all(
        not result["intersects_with_clearance"]
        for result in pairwise.values()
    )

    sheet_bounds = _sheet_bounds_verification(
        rectangles,
        paper_width_mm,
        paper_height_mm,
        sheet_margin_mm,
    )
    group_bounds = _group_bounds(rectangles)
    group_fit = _group_fit_verification(
        group_bounds,
        paper_width_mm,
        paper_height_mm,
        sheet_margin_mm,
    )
    group_translation_plan = _minimum_group_translation_to_margins(
        group_bounds,
        paper_width_mm,
        paper_height_mm,
        sheet_margin_mm,
    )

    verified = bool(
        top_direction_verified
        and right_direction_verified
        and pairwise_clear
        and sheet_bounds["all_views_within_sheet_margin"]
    )

    return {
        "verified": verified,
        "projection_method": projection_method,
        "minimum_gap_mm": float(minimum_gap_mm),
        "layout_tolerance_mm": layout_tolerance_mm,
        "sheet_margin_mm": float(sheet_margin_mm),
        "rectangles": rectangles,
        "read_errors": errors,
        "front_top_directional_gap_mm": front_top_gap,
        "front_right_directional_gap_mm": front_right_gap,
        "top_direction_verified": top_direction_verified,
        "right_direction_verified": right_direction_verified,
        "pairwise": pairwise,
        "pairwise_clear_with_minimum_gap": pairwise_clear,
        "sheet_bounds": sheet_bounds,
        "group_bounds": group_bounds,
        "group_fit": group_fit,
        "group_translation_plan": group_translation_plan,
    }


def _set_view_xy(
    view: Any,
    *,
    x: Optional[float] = None,
    y: Optional[float] = None,
) -> list[dict[str, Any]]:
    attempts: list[dict[str, Any]] = []
    for attribute, value in (("x", x), ("y", y)):
        if value is None:
            continue
        try:
            setattr(view, attribute, float(value))
            attempts.append(
                {
                    "method": f"view.{attribute}={float(value)}",
                    "succeeded": True,
                    "error": None,
                }
            )
        except Exception as exc:
            attempts.append(
                {
                    "method": f"view.{attribute}={float(value)}",
                    "succeeded": False,
                    "error": _format_com_error(exc),
                }
            )
    return attempts


def _resolve_three_view_overlap(
    application: Any,
    drawing_document: Any,
    sheet: Any,
    front_view: Any,
    top_view: Any,
    right_view: Any,
    projection_method: str,
    minimum_gap_mm: float,
    paper_width_mm: float,
    paper_height_mm: float,
    sheet_margin_mm: float,
    *,
    maximum_iterations: int = 8,
) -> dict[str, Any]:
    iterations: list[dict[str, Any]] = []

    for iteration_index in range(1, maximum_iterations + 1):
        refresh_attempts: list[dict[str, Any]] = []
        for view in (front_view, top_view, right_view):
            refresh_attempts.extend(
                _refresh_drafting_view(
                    application,
                    drawing_document,
                    sheet,
                    view,
                )
            )

        time.sleep(0.20)
        before = _layout_snapshot(
            application,
            front_view,
            top_view,
            right_view,
            projection_method,
            minimum_gap_mm,
            paper_width_mm,
            paper_height_mm,
            sheet_margin_mm,
        )

        if before["verified"]:
            iterations.append(
                {
                    "iteration": iteration_index,
                    "layout_before": before,
                    "adjustments": [],
                    "refresh_attempts": refresh_attempts,
                    "layout_after": before,
                }
            )
            return {
                "verified": True,
                "minimum_gap_mm": float(minimum_gap_mm),
                "sheet_margin_mm": float(sheet_margin_mm),
                "maximum_iterations": maximum_iterations,
                "iterations": iterations,
                "global_group_fitting_enabled": True,
                "global_group_fitting_policy": (
                    "minimum_translation_to_sheet_margins"
                ),
                "final_layout": before,
            }

        rectangles = before.get("rectangles", {})
        if len(rectangles) != 3:
            iterations.append(
                {
                    "iteration": iteration_index,
                    "layout_before": before,
                    "adjustments": [],
                    "refresh_attempts": refresh_attempts,
                    "layout_after": before,
                }
            )
            break

        front = rectangles["front"]
        top = rectangles["top"]
        right = rectangles["right"]
        adjustments: list[dict[str, Any]] = []

        top_target_y = float(_safe_attribute(top_view, "y", 0.0))
        right_target_x = float(_safe_attribute(right_view, "x", 0.0))

        if projection_method == "third_angle":
            required_top_delta = max(
                0.0,
                front["ymax"] + minimum_gap_mm - top["ymin"],
            )
            required_right_delta = max(
                0.0,
                front["xmax"] + minimum_gap_mm - right["xmin"],
            )
            top_target_y += required_top_delta
            right_target_x += required_right_delta
        else:
            required_top_delta = min(
                0.0,
                front["ymin"] - minimum_gap_mm - top["ymax"],
            )
            required_right_delta = min(
                0.0,
                front["xmin"] - minimum_gap_mm - right["xmax"],
            )
            top_target_y += required_top_delta
            right_target_x += required_right_delta

        # Resolve the diagonal Top/Right pair as well. Moving the Right
        # view farther horizontally preserves its standard projection
        # alignment with Front.
        projected_top = dict(top)
        projected_right = dict(right)
        top_shift = top_target_y - float(
            _safe_attribute(top_view, "y", 0.0)
        )
        right_shift = right_target_x - float(
            _safe_attribute(right_view, "x", 0.0)
        )
        projected_top["ymin"] += top_shift
        projected_top["ymax"] += top_shift
        projected_right["xmin"] += right_shift
        projected_right["xmax"] += right_shift

        top_right_check = _rectangle_overlap(
            projected_top,
            projected_right,
            minimum_gap_mm,
        )
        if top_right_check["intersects_with_clearance"]:
            if projection_method == "third_angle":
                additional = max(
                    0.0,
                    projected_top["xmax"]
                    + minimum_gap_mm
                    - projected_right["xmin"],
                )
            else:
                additional = min(
                    0.0,
                    projected_top["xmin"]
                    - minimum_gap_mm
                    - projected_right["xmax"],
                )
            right_target_x += additional

        current_top_y = float(
            _safe_attribute(top_view, "y", 0.0)
        )
        current_right_x = float(
            _safe_attribute(right_view, "x", 0.0)
        )

        if abs(top_target_y - current_top_y) > 1e-9:
            attempts = _position_projected_view(
                application,
                top_view,
                front_view,
                "top",
                float(_safe_attribute(front_view, "x", 0.0)),
                top_target_y,
            )
            adjustments.append(
                {
                    "view": "top",
                    "axis": "y",
                    "before_mm": current_top_y,
                    "target_mm": top_target_y,
                    "delta_mm": top_target_y - current_top_y,
                    "position_result": attempts,
                    "position_verified": attempts[
                        "position_verified"
                    ],
                }
            )

        if abs(right_target_x - current_right_x) > 1e-9:
            attempts = _position_projected_view(
                application,
                right_view,
                front_view,
                "right",
                right_target_x,
                float(_safe_attribute(front_view, "y", 0.0)),
            )
            adjustments.append(
                {
                    "view": "right",
                    "axis": "x",
                    "before_mm": current_right_x,
                    "target_mm": right_target_x,
                    "delta_mm": right_target_x - current_right_x,
                    "position_result": attempts,
                    "position_verified": attempts[
                        "position_verified"
                    ],
                }
            )

        for view in (front_view, top_view, right_view):
            refresh_attempts.extend(
                _refresh_drafting_view(
                    application,
                    drawing_document,
                    sheet,
                    view,
                )
            )
        time.sleep(0.20)

        after_local_adjustment = _layout_snapshot(
            application,
            front_view,
            top_view,
            right_view,
            projection_method,
            minimum_gap_mm,
            paper_width_mm,
            paper_height_mm,
            sheet_margin_mm,
        )

        group_reposition = {
            "attempted": False,
            "verified": None,
            "reason": (
                "Pairwise layout is not ready for global sheet fitting."
            ),
        }
        after = after_local_adjustment
        local_pairwise_ready = bool(
            after_local_adjustment.get("top_direction_verified", False)
            and after_local_adjustment.get("right_direction_verified", False)
            and after_local_adjustment.get(
                "pairwise_clear_with_minimum_gap", False
            )
        )
        group_plan = after_local_adjustment.get(
            "group_translation_plan", {}
        )

        if (
            local_pairwise_ready
            and not after_local_adjustment["sheet_bounds"][
                "all_views_within_sheet_margin"
            ]
        ):
            if not group_plan.get("possible", False):
                group_reposition = {
                    "attempted": False,
                    "verified": False,
                    "reason": group_plan.get(
                        "reason",
                        "The three-view group cannot fit in the usable sheet area.",
                    ),
                    "translation_plan": group_plan,
                }
            elif group_plan.get("translation_required", False):
                group_reposition = _translate_three_view_group(
                    application,
                    drawing_document,
                    sheet,
                    front_view,
                    top_view,
                    right_view,
                    group_plan["required_dx_mm"],
                    group_plan["required_dy_mm"],
                )
                group_reposition["translation_plan"] = group_plan
                after = _layout_snapshot(
                    application,
                    front_view,
                    top_view,
                    right_view,
                    projection_method,
                    minimum_gap_mm,
                    paper_width_mm,
                    paper_height_mm,
                    sheet_margin_mm,
                )
                group_reposition["layout_after_translation"] = after
                group_reposition["verified"] = bool(
                    group_reposition["verified"]
                    and after["sheet_bounds"]["all_views_within_sheet_margin"]
                    and after["pairwise_clear_with_minimum_gap"]
                    and after["top_direction_verified"]
                    and after["right_direction_verified"]
                )
            else:
                group_reposition = {
                    "attempted": False,
                    "verified": True,
                    "reason": (
                        "The group already satisfies the sheet margins; "
                        "no global translation is required."
                    ),
                    "translation_plan": group_plan,
                }

        position_writes_verified = all(
            item.get("position_verified", True)
            for item in adjustments
        )
        if group_reposition.get("attempted", False):
            position_writes_verified = bool(
                position_writes_verified
                and group_reposition.get("verified", False)
            )

        iterations.append(
            {
                "iteration": iteration_index,
                "layout_before": before,
                "adjustments": adjustments,
                "position_writes_verified": position_writes_verified,
                "refresh_attempts": refresh_attempts,
                "layout_after_local_adjustment": after_local_adjustment,
                "group_reposition": group_reposition,
                "layout_after": after,
            }
        )

        if after["verified"]:
            return {
                "verified": True,
                "minimum_gap_mm": float(minimum_gap_mm),
                "sheet_margin_mm": float(sheet_margin_mm),
                "maximum_iterations": maximum_iterations,
                "iterations": iterations,
                "global_group_fitting_enabled": True,
                "global_group_fitting_policy": (
                    "minimum_translation_to_sheet_margins"
                ),
                "final_layout": after,
            }

        effective_change_attempted = bool(
            adjustments or group_reposition.get("attempted", False)
        )
        if not effective_change_attempted:
            break

    final_layout = _layout_snapshot(
        application,
        front_view,
        top_view,
        right_view,
        projection_method,
        minimum_gap_mm,
        paper_width_mm,
        paper_height_mm,
        sheet_margin_mm,
    )
    return {
        "verified": False,
        "minimum_gap_mm": float(minimum_gap_mm),
        "sheet_margin_mm": float(sheet_margin_mm),
        "maximum_iterations": maximum_iterations,
        "iterations": iterations,
        "global_group_fitting_enabled": True,
        "global_group_fitting_policy": (
            "minimum_translation_to_sheet_margins"
        ),
        "insufficient_space_confirmed": bool(
            not final_layout.get("group_fit", {}).get(
                "group_fits_usable_area", False
            )
        ),
        "final_layout": final_layout,
        "error": (
            "Three-view layout could not satisfy the requested "
            "pairwise clearance and sheet-boundary constraints."
        ),
    }



def _view_summary(
    application: Any,
    view: Any,
    *,
    expected_type: Optional[str] = None,
) -> dict[str, Any]:
    name = safe_str(_safe_attribute(view, "Name", ""))
    view_type_code = int(
        _safe_attribute(view, "ViewType", -1)
    )
    view_type_name = DRAWING_VIEW_TYPE_NAMES.get(
        view_type_code,
        "unknown",
    )

    (
        is_generative,
        is_generative_read_method,
        is_generative_read_attempts,
    ) = _view_is_generative(application, view)

    scale_value, scale_method_or_error = _view_scale(view)

    reference_view_name: Optional[str] = None
    try:
        reference_view_name = safe_str(view.ReferenceView.Name)
    except Exception:
        pass

    behavior = None
    parent_view_name: Optional[str] = None
    represented_document_name: Optional[str] = None
    projection: Optional[dict[str, Any]] = None
    projection_error: Optional[str] = None

    try:
        behavior = view.GenerativeBehavior
        try:
            parent_view_name = safe_str(
                behavior.ParentView.Name
            )
        except Exception:
            pass
        try:
            represented_document_name = safe_str(
                behavior.Document.Name
            )
        except Exception:
            pass
        try:
            projection = _projection_plane(
                application,
                behavior,
            )
        except Exception as exc:
            projection_error = str(exc)
    except Exception:
        pass

    size: Optional[dict[str, Any]] = None
    size_error: Optional[str] = None
    try:
        size = _view_size(application, view)
    except Exception as exc:
        size_error = str(exc)

    geometric_elements_count = _safe_count(
        _safe_attribute(view, "GeometricElements")
    )
    dimensions_count = _safe_count(
        _safe_attribute(view, "Dimensions")
    )
    texts_count = _safe_count(
        _safe_attribute(view, "Texts")
    )
    tables_count = _safe_count(
        _safe_attribute(view, "Tables")
    )

    geometry_nonempty = bool(
        (
            geometric_elements_count is not None
            and geometric_elements_count > 0
        )
        or (
            size is not None
            and (
                abs(size["width_mm"]) > 1e-9
                or abs(size["height_mm"]) > 1e-9
            )
        )
    )
    type_verified = bool(
        expected_type is None
        or view_type_name == expected_type
    )

    return {
        "name": name,
        "expected_type": expected_type,
        "view_type_code": view_type_code,
        "view_type_name": view_type_name,
        "type_verified": type_verified,
        "is_generative": is_generative,
        "is_generative_read_method": (
            is_generative_read_method
        ),
        "is_generative_read_attempts": (
            is_generative_read_attempts
        ),
        "x_mm": float(_safe_attribute(view, "x", 0.0)),
        "y_mm": float(_safe_attribute(view, "y", 0.0)),
        "scale": scale_value,
        "scale_read_method_or_error": scale_method_or_error,
        "reference_view_name": reference_view_name,
        "parent_view_name": parent_view_name,
        "represented_document_name": (
            represented_document_name
        ),
        "projection_plane": projection,
        "projection_plane_error": projection_error,
        "bounding_box": size,
        "bounding_box_error": size_error,
        "geometric_elements_count": geometric_elements_count,
        "dimensions_count": dimensions_count,
        "texts_count": texts_count,
        "tables_count": tables_count,
        "geometry_nonempty": geometry_nonempty,
        "view_verified": bool(
            type_verified
            and is_generative
            and geometry_nonempty
        ),
    }


def _add_front_view(
    sheet: Any,
    model_document: Any,
    represented_object: Any,
    layout: DrawingLayout,
    scale: float,
) -> Any:
    view = sheet.Views.Add("Front view")
    view.x = float(layout.front_x)
    view.y = float(layout.front_y)
    _set_scale(view, scale)

    behavior = view.GenerativeBehavior
    try:
        behavior.Document = represented_object
    except Exception:
        behavior.Document = model_document

    behavior.DefineFrontView(
        1.0, 0.0, 0.0,
        0.0, 0.0, 1.0,
    )
    try:
        behavior.ForceUpdate()
    except Exception:
        behavior.Update()
    return view


def _define_projection_view(
    application: Any,
    child_behavior: Any,
    parent_behavior: Any,
    projection_type: str,
) -> None:
    script = (
        "Public Function MCP_DefineProjectionView("
        "childBehavior, parentBehavior, viewTypeName)\n"
        "    Select Case LCase(CStr(viewTypeName))\n"
        "    Case \"top\"\n"
        "        childBehavior.DefineProjectionView "
        "parentBehavior, catTopView\n"
        "    Case \"right\"\n"
        "        childBehavior.DefineProjectionView "
        "parentBehavior, catRightView\n"
        "    Case Else\n"
        "        Err.Raise 5, , \"Unsupported projection view type\"\n"
        "    End Select\n"
        "    MCP_DefineProjectionView = True\n"
        "End Function"
    )
    _evaluate(
        application,
        script,
        "MCP_DefineProjectionView",
        [
            child_behavior,
            parent_behavior,
            projection_type,
        ],
    )


def _add_projected_view(
    application: Any,
    sheet: Any,
    front_view: Any,
    model_document: Any,
    represented_object: Any,
    name: str,
    projection_type: str,
    x: float,
    y: float,
    scale: float,
) -> Any:
    view = sheet.Views.Add(name)
    behavior = view.GenerativeBehavior
    parent_behavior = front_view.GenerativeBehavior

    try:
        behavior.Document = represented_object
    except Exception:
        try:
            behavior.Document = model_document
        except Exception:
            pass

    try:
        view.ReferenceView = front_view
    except Exception:
        pass

    _define_projection_view(
        application,
        behavior,
        parent_behavior,
        projection_type,
    )

    _set_scale(view, scale)
    _position_projected_view(
        application,
        view,
        front_view,
        projection_type,
        x,
        y,
    )

    try:
        behavior.ForceUpdate()
    except Exception:
        behavior.Update()
    return view


def _verify_three_views(
    view_summaries: list[dict[str, Any]],
) -> dict[str, Any]:
    by_type = {
        item["view_type_name"]: item
        for item in view_summaries
    }
    required = ("front", "top", "right")
    missing = [item for item in required if item not in by_type]

    front_name = (
        by_type["front"]["name"]
        if "front" in by_type
        else None
    )
    projected_links = {}

    for projection_type in ("top", "right"):
        summary = by_type.get(projection_type)
        if summary is None:
            projected_links[projection_type] = False
            continue
        projected_links[projection_type] = bool(
            summary["parent_view_name"] == front_name
            or summary["reference_view_name"] == front_name
        )

    all_views_verified = bool(
        not missing
        and all(
            by_type[item]["view_verified"]
            for item in required
        )
        and all(projected_links.values())
    )

    return {
        "required_view_types": list(required),
        "missing_view_types": missing,
        "projected_parent_links": projected_links,
        "all_views_verified": all_views_verified,
    }


# ---------------------------------------------------------------------------
# Dimensions, text and title-block helpers
# ---------------------------------------------------------------------------

def _dimension_count(sheet: Any) -> dict[str, Any]:
    views = sheet.Views
    total = 0
    per_view: list[dict[str, Any]] = []

    for index in range(1, int(views.Count) + 1):
        view = views.Item(index)
        count = _safe_count(_safe_attribute(view, "Dimensions"))
        count_value = count if count is not None else 0
        total += count_value
        per_view.append(
            {
                "index": index,
                "name": safe_str(
                    _safe_attribute(view, "Name", "")
                ),
                "count": count,
            }
        )

    return {
        "total": total,
        "per_view": per_view,
    }


def _generate_dimensions_internal(
    sheet: Any,
    drawing_document: Any,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    before = _dimension_count(sheet)

    sheet.GenerateDimensions()
    update_attempts, update_warnings = _call_update(
        sheet,
        drawing_document,
    )
    warnings.extend(update_warnings)

    after = _dimension_count(sheet)
    generated_count = after["total"] - before["total"]
    generated = generated_count > 0

    if not generated:
        warnings.append(
            "DrawingSheet.GenerateDimensions completed, but no new "
            "dimensions were created. CATIA only generates dimensions "
            "from eligible 3D constraints."
        )

    return {
        "generation_call_succeeded": True,
        "call_method": "DrawingSheet.GenerateDimensions",
        "dimensions_before": before,
        "dimensions_after": after,
        "generated_count": generated_count,
        "generated": generated,
        "generation_verified": generated,
        "update_attempts": update_attempts,
        "model_modified": generated,
        "document_save_required": generated,
    }, warnings


def _font_size_readback(
    application: Any,
    drawing_text: Any,
) -> Optional[float]:
    script = (
        "Public Function MCP_GetDrawingTextFontSize(textObject)\n"
        "    Dim fontSize\n"
        "    textObject.GetFontSize 0, 0, fontSize\n"
        "    MCP_GetDrawingTextFontSize = CDbl(fontSize)\n"
        "End Function"
    )
    try:
        return float(
            _evaluate(
                application,
                script,
                "MCP_GetDrawingTextFontSize",
                [drawing_text],
            )
        )
    except Exception:
        return None


def _add_text_internal(
    application: Any,
    drawing_document: Any,
    sheet: Any,
    text: str,
    x: float,
    y: float,
    font_size: float,
    use_background_view: bool,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    text_value = _nonempty_text(text, "text")
    x_value = _finite_float(x, "x")
    y_value = _finite_float(y, "y")
    font_value = _positive_float(font_size, "font_size")

    if use_background_view:
        view, view_selection = _background_view(sheet)
    else:
        view = sheet.Views.ActiveView
        view_selection = {
            "selection_method": "DrawingSheet.Views.ActiveView",
            "name": safe_str(_safe_attribute(view, "Name", "")),
            "view_type_code": int(
                _safe_attribute(view, "ViewType", -1)
            ),
            "view_type_name": DRAWING_VIEW_TYPE_NAMES.get(
                int(_safe_attribute(view, "ViewType", -1)),
                "unknown",
            ),
        }

    texts = view.Texts
    before_count = int(texts.Count)
    drawing_text = texts.Add(
        text_value,
        x_value,
        y_value,
    )

    font_set = False
    font_error: Optional[str] = None
    try:
        drawing_text.SetFontSize(0, 0, font_value)
        font_set = True
    except Exception as exc:
        font_error = _format_com_error(exc)
        warnings.append(
            f"Text was added, but SetFontSize failed: {font_error}"
        )

    update_attempts, update_warnings = _call_update(
        sheet,
        drawing_document,
    )
    warnings.extend(update_warnings)

    after_count = int(texts.Count)
    actual_text = safe_str(
        _safe_attribute(drawing_text, "Text", "")
    )
    font_readback = _font_size_readback(
        application,
        drawing_text,
    )
    added_verified = bool(
        after_count == before_count + 1
        and actual_text == text_value
    )

    if not added_verified:
        raise RuntimeError(
            "Drawing text creation could not be verified."
        )

    return {
        "added": True,
        "added_verified": True,
        "text": text_value,
        "actual_text": actual_text,
        "x_mm": x_value,
        "y_mm": y_value,
        "font_size_requested_mm": font_value,
        "font_size_set": font_set,
        "font_size_error": font_error,
        "font_size_readback_mm": font_readback,
        "texts_count_before": before_count,
        "texts_count_after": after_count,
        "target_view": view_selection,
        "update_attempts": update_attempts,
        "model_modified": True,
        "document_save_required": True,
    }, warnings


def _title_block_bounds(
    x_mm: float,
    y_mm: float,
    paper_width_mm: float,
    paper_height_mm: float,
    margin_mm: float,
) -> dict[str, Any]:
    xmin = float(x_mm)
    xmax = xmin + TITLE_BLOCK_WIDTH_MM
    ymin = float(y_mm)
    ymax = ymin + TITLE_BLOCK_HEIGHT_MM
    within = bool(
        xmin >= margin_mm
        and ymin >= margin_mm
        and xmax <= paper_width_mm - margin_mm
        and ymax <= paper_height_mm - margin_mm
    )
    return {
        "xmin_mm": xmin,
        "xmax_mm": xmax,
        "ymin_mm": ymin,
        "ymax_mm": ymax,
        "table_width_mm": TITLE_BLOCK_WIDTH_MM,
        "table_height_mm": TITLE_BLOCK_HEIGHT_MM,
        "paper_width_mm": float(paper_width_mm),
        "paper_height_mm": float(paper_height_mm),
        "sheet_margin_mm": float(margin_mm),
        "left_clearance_mm": xmin,
        "right_clearance_mm": paper_width_mm - xmax,
        "bottom_clearance_mm": ymin,
        "top_clearance_mm": paper_height_mm - ymax,
        "within_sheet_margin": within,
    }


def _resolve_title_block_position(
    application: Any,
    sheet: Any,
    x: Optional[float],
    y: Optional[float],
    sheet_margin_mm: float,
) -> dict[str, Any]:
    margin = _positive_float(
        sheet_margin_mm,
        "sheet_margin_mm",
    )
    dimensions = _get_paper_dimensions(application, sheet)
    paper_width = float(dimensions["width_mm"])
    paper_height = float(dimensions["height_mm"])

    usable_width = paper_width - 2.0 * margin
    usable_height = paper_height - 2.0 * margin
    if (
        TITLE_BLOCK_WIDTH_MM > usable_width
        or TITLE_BLOCK_HEIGHT_MM > usable_height
    ):
        raise ValueError(
            "The title block cannot fit inside the current sheet "
            "with the requested margin."
        )

    auto_x = x is None
    auto_y = y is None
    resolved_x = (
        paper_width - margin - TITLE_BLOCK_WIDTH_MM
        if auto_x
        else _finite_float(x, "x")
    )
    resolved_y = (
        min(
            DEFAULT_TITLE_BLOCK_Y_MM,
            paper_height - margin - TITLE_BLOCK_HEIGHT_MM,
        )
        if auto_y
        else _finite_float(y, "y")
    )
    resolved_y = max(margin, resolved_y)

    bounds = _title_block_bounds(
        resolved_x,
        resolved_y,
        paper_width,
        paper_height,
        margin,
    )
    if not bounds["within_sheet_margin"]:
        coordinate_kind = (
            "resolved default coordinates"
            if auto_x or auto_y
            else "explicit coordinates"
        )
        raise ValueError(
            f"Title block {coordinate_kind} place the table outside "
            f"the sheet margin: {bounds}"
        )

    return {
        "x_mm": resolved_x,
        "y_mm": resolved_y,
        "x_auto_positioned": auto_x,
        "y_auto_positioned": auto_y,
        "auto_positioned": bool(auto_x or auto_y),
        "paper_dimensions": dimensions,
        "bounds": bounds,
        "bounds_verified": True,
    }


def _add_title_block_internal(
    application: Any,
    drawing_document: Any,
    sheet: Any,
    title: str,
    part_number: str,
    material: str,
    general_tolerance: str,
    drawn_by: str,
    x: Optional[float],
    y: Optional[float],
    sheet_margin_mm: float = DEFAULT_TITLE_BLOCK_MARGIN_MM,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    position = _resolve_title_block_position(
        application,
        sheet,
        x,
        y,
        sheet_margin_mm,
    )
    view, view_selection = _background_view(sheet)
    tables = view.Tables
    before_count = int(tables.Count)

    table = tables.Add(
        position["x_mm"],
        position["y_mm"],
        TITLE_BLOCK_ROWS,
        TITLE_BLOCK_COLUMNS,
        TITLE_BLOCK_ROW_HEIGHT_MM,
        TITLE_BLOCK_COLUMN_WIDTH_MM,
    )

    rows = [
        ("TITLE", str(title)),
        ("PART NO.", str(part_number)),
        ("MATERIAL", str(material)),
        ("GENERAL TOL.", str(general_tolerance)),
        ("DRAWN BY", str(drawn_by)),
        (
            "SCALE",
            safe_str(
                _safe_attribute(
                    sheet,
                    "Scale",
                    _safe_attribute(sheet, "Scale2", ""),
                )
            ),
        ),
    ]

    for row_index, (label, value) in enumerate(
        rows,
        start=1,
    ):
        table.SetCellString(row_index, 1, label)
        table.SetCellString(row_index, 2, value)

    update_attempts, update_warnings = _call_update(
        sheet,
        drawing_document,
    )
    warnings.extend(update_warnings)

    after_count = int(tables.Count)
    cell_readback: list[dict[str, Any]] = []
    readback_verified = True

    for row_index, (expected_label, expected_value) in enumerate(
        rows,
        start=1,
    ):
        try:
            actual_label = safe_str(
                table.GetCellString(row_index, 1)
            )
            actual_value = safe_str(
                table.GetCellString(row_index, 2)
            )
        except Exception:
            actual_label = ""
            actual_value = ""
            readback_verified = False

        if (
            actual_label != expected_label
            or actual_value != expected_value
        ):
            readback_verified = False

        cell_readback.append(
            {
                "row": row_index,
                "expected": [
                    expected_label,
                    expected_value,
                ],
                "actual": [
                    actual_label,
                    actual_value,
                ],
            }
        )

    added_verified = bool(
        after_count == before_count + 1
        and readback_verified
    )
    if not added_verified:
        raise RuntimeError(
            "Title-block table creation or cell readback failed."
        )

    return {
        "added": True,
        "added_verified": True,
        "rows": rows,
        "cell_readback": cell_readback,
        "tables_count_before": before_count,
        "tables_count_after": after_count,
        "target_view": view_selection,
        "x_mm": position["x_mm"],
        "y_mm": position["y_mm"],
        "position": position,
        "bounds_verified": position["bounds_verified"],
        "update_attempts": update_attempts,
        "model_modified": True,
        "document_save_required": True,
    }, warnings


def _engineering_note_lines(
    dimension_notes: list[str],
    tolerance_notes: list[str],
    gdt_notes: list[str],
) -> list[str]:
    lines: list[str] = []

    if dimension_notes:
        lines.append("DIMENSION NOTES:")
        lines.extend(
            f"D{index}. {note}"
            for index, note in enumerate(
                dimension_notes,
                start=1,
            )
        )

    if tolerance_notes:
        if lines:
            lines.append("")
        lines.append("TOLERANCE NOTES:")
        lines.extend(
            f"T{index}. {note}"
            for index, note in enumerate(
                tolerance_notes,
                start=1,
            )
        )

    if gdt_notes:
        if lines:
            lines.append("")
        lines.append("GD&T NOTES:")
        lines.extend(
            f"G{index}. {note}"
            for index, note in enumerate(
                gdt_notes,
                start=1,
            )
        )

    if not lines:
        lines = [
            "GENERAL NOTES:",
            (
                "1. Dimensions and GD&T require engineering "
                "review before manufacturing release."
            ),
        ]

    return lines


# ---------------------------------------------------------------------------
# Drawing and source summaries
# ---------------------------------------------------------------------------

def _drawing_summary(
    conn: Any,
    application: Any,
    drawing_document: Any,
) -> dict[str, Any]:
    info = conn.describe_document(drawing_document)

    try:
        sheets = drawing_document.Sheets
        sheet = sheets.ActiveSheet
        paper_dimensions = _get_paper_dimensions(
            application,
            sheet,
        )
        scale, scale_method = _view_scale(sheet)
        info.update(
            {
                "sheets_count": int(sheets.Count),
                "active_sheet": safe_str(sheet.Name),
                "views_count": int(sheet.Views.Count),
                "paper_size_code": int(sheet.PaperSize),
                "orientation_code": int(sheet.Orientation),
                "projection_method_code": int(
                    sheet.ProjectionMethod
                ),
                "scale": scale,
                "scale_read_method_or_error": scale_method,
                "paper_dimensions": paper_dimensions,
                "saved": _document_saved(drawing_document),
                "full_name": _document_full_name(
                    drawing_document
                ),
            }
        )
    except Exception as exc:
        info["summary_error"] = _format_com_error(exc)

    return info


# ---------------------------------------------------------------------------
# Final safe-area annotation verification
# ---------------------------------------------------------------------------

def _verify_final_safe_area_annotations(
    application: Any,
    sheet: Any,
    requested_sheet_margin_mm: float,
    *,
    attempt_repair: bool = True,
) -> dict[str, Any]:
    """Verify view geometry plus view-local dimensions/text against one sheet safe area.

    The implementation lives in smart_annotation_layout so exact dimension
    GetBoundaryBox handling and local-to-sheet coordinate conversion have one owner.
    System/background template views are excluded, so the template itself is never
    moved, rewritten or treated as an annotation that must be repaired.
    """
    try:
        from catia_mcp.tools.smart_annotation_layout import (
            verify_annotations_within_sheet_safe_area,
        )
    except Exception as exc:
        raise DraftingOperationError(
            "Safe-area annotation verification is unavailable: "
            f"{_format_com_error(exc)}"
        ) from exc

    result = verify_annotations_within_sheet_safe_area(
        application,
        sheet_index=None,
        requested_sheet_margin_mm=float(requested_sheet_margin_mm),
        inner_frame_inset_mm=INNER_FRAME_INSET_MM,
        include_dimensions=True,
        include_texts=True,
        include_view_geometry=True,
        attempt_repair=bool(attempt_repair),
        strict_readback=True,
        exclude_system_views=True,
    )
    if not result.get("ok", False):
        raise DraftingOperationError(
            "Views or annotations cross the configured inner-frame safety area.",
            data={"final_safe_area_verification": result},
            warnings=list(result.get("warnings", [])),
        )
    return result


# ---------------------------------------------------------------------------
# Core three-view creation
# ---------------------------------------------------------------------------

def _create_3view_drawing_from_model_doc(
    conn: Any,
    application: Any,
    model_document: Any,
    paper_size: str,
    orientation: str,
    projection_method: str,
    scale: float,
    drawing_title: str,
    part_number: str,
    material: str,
    general_tolerance: str,
    generate_dimensions: bool,
    dimension_notes: list[str],
    tolerance_notes: list[str],
    gdt_notes: list[str],
    output_path: str,
    export_pdf_path: str,
    overwrite: bool,
    minimum_view_gap_mm: float,
    sheet_margin_mm: float,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    drawing_document = None
    drawing_cleanup = None

    minimum_gap_value = _positive_float(
        minimum_view_gap_mm,
        "minimum_view_gap_mm",
    )
    sheet_margin_value = _positive_float(
        sheet_margin_mm,
        "sheet_margin_mm",
    )

    model_info = conn.describe_document(model_document)
    if model_info.get("kind") not in ("CATPart", "CATProduct"):
        raise CATIAError(
            "Active/source document must be CATPart or CATProduct."
        )

    if output_path:
        _prepare_output_path(
            output_path,
            "output_path",
            ".CATDrawing",
            overwrite,
        )
    if export_pdf_path:
        _prepare_output_path(
            export_pdf_path,
            "export_pdf_path",
            ".pdf",
            overwrite,
        )

    try:
        represented_object = _represented_3d_object(
            model_document
        )
        drawing_document = application.Documents.Add(
            "Drawing"
        )
        sheet = drawing_document.Sheets.ActiveSheet

        sheet_configuration = _configure_sheet(
            application,
            sheet,
            paper_size,
            orientation,
            scale,
            projection_method,
        )
        projection_name = sheet_configuration[
            "projection_method"
        ]["requested"]
        paper_dimensions = sheet_configuration[
            "paper_dimensions"
        ]
        safe_area = _safe_area_descriptor(
            paper_dimensions["width_mm"],
            paper_dimensions["height_mm"],
            sheet_margin_value,
        )
        effective_safe_margin = safe_area["effective_safe_margin_mm"]
        layout = _calculate_layout(
            paper_dimensions["width_mm"],
            paper_dimensions["height_mm"],
            projection_name,
            effective_safe_margin,
        )

        views_before = int(sheet.Views.Count)

        front_view = _add_front_view(
            sheet,
            model_document,
            represented_object,
            layout,
            float(scale),
        )
        top_view = _add_projected_view(
            application,
            sheet,
            front_view,
            model_document,
            represented_object,
            "Top view",
            "top",
            layout.top_x,
            layout.top_y,
            float(scale),
        )
        right_view = _add_projected_view(
            application,
            sheet,
            front_view,
            model_document,
            represented_object,
            "Right view",
            "right",
            layout.right_x,
            layout.right_y,
            float(scale),
        )

        update_attempts, update_warnings = _call_update(
            sheet,
            drawing_document,
        )
        warnings.extend(update_warnings)

        view_generation_waits = {
            "front": _wait_for_view_generation(
                application,
                drawing_document,
                sheet,
                front_view,
            ),
            "top": _wait_for_view_generation(
                application,
                drawing_document,
                sheet,
                top_view,
                position_callback=lambda: (
                    _position_projected_view(
                        application,
                        top_view,
                        front_view,
                        "top",
                        layout.top_x,
                        layout.top_y,
                    )
                ),
            ),
            "right": _wait_for_view_generation(
                application,
                drawing_document,
                sheet,
                right_view,
                position_callback=lambda: (
                    _position_projected_view(
                        application,
                        right_view,
                        front_view,
                        "right",
                        layout.right_x,
                        layout.right_y,
                    )
                ),
            ),
        }

        layout_resolution = _resolve_three_view_overlap(
            application,
            drawing_document,
            sheet,
            front_view,
            top_view,
            right_view,
            projection_name,
            minimum_gap_value,
            paper_dimensions["width_mm"],
            paper_dimensions["height_mm"],
            effective_safe_margin,
        )

        views_after = int(sheet.Views.Count)
        view_summaries = [
            _view_summary(
                application,
                front_view,
                expected_type="front",
            ),
            _view_summary(
                application,
                top_view,
                expected_type="top",
            ),
            _view_summary(
                application,
                right_view,
                expected_type="right",
            ),
        ]
        view_verification = _verify_three_views(
            view_summaries
        )
        generation_waits_verified = all(
            item["verified"]
            for item in view_generation_waits.values()
        )
        creation_verified = bool(
            views_after == views_before + 3
            and view_verification["all_views_verified"]
            and generation_waits_verified
            and layout_resolution["verified"]
        )

        if not creation_verified:
            raise DraftingOperationError(
                "Three-view drawing creation could not be "
                "fully verified.",
                data={
                    "views_count_before": views_before,
                    "views_count_after": views_after,
                    "views": view_summaries,
                    "view_verification": view_verification,
                    "view_generation_waits": (
                        view_generation_waits
                    ),
                    "generation_waits_verified": (
                        generation_waits_verified
                    ),
                    "layout_resolution": layout_resolution,
                    "layout_collision_free": (
                        layout_resolution["verified"]
                    ),
                },
            )

        dimension_generation = {
            "requested": bool(generate_dimensions),
            "generation_call_succeeded": False,
            "generated": False,
            "generated_count": 0,
        }
        if generate_dimensions:
            (
                dimension_generation,
                dimension_warnings,
            ) = _generate_dimensions_internal(
                sheet,
                drawing_document,
            )
            dimension_generation["requested"] = True
            warnings.extend(dimension_warnings)

        title_value = (
            drawing_title
            or str(model_info.get("name", ""))
        )
        part_number_value = (
            part_number
            or str(model_info.get("part_number", ""))
        )

        title_block, title_warnings = (
            _add_title_block_internal(
                application,
                drawing_document,
                sheet,
                title_value,
                part_number_value,
                material,
                general_tolerance,
                "CATIA MCP",
                layout.title_block_x,
                layout.title_block_y,
                effective_safe_margin,
            )
        )
        warnings.extend(title_warnings)

        notes_result = None
        if dimension_notes or tolerance_notes or gdt_notes:
            lines = _engineering_note_lines(
                dimension_notes,
                tolerance_notes,
                gdt_notes,
            )
            notes_result, note_warnings = _add_text_internal(
                application,
                drawing_document,
                sheet,
                "\n".join(lines),
                layout.notes_x,
                layout.notes_y,
                3.2,
                True,
            )
            warnings.extend(note_warnings)

        final_update_attempts, final_update_warnings = (
            _call_update(sheet, drawing_document)
        )
        warnings.extend(final_update_warnings)

        final_safe_area_verification = _verify_final_safe_area_annotations(
            application,
            sheet,
            sheet_margin_value,
            attempt_repair=True,
        )
        warnings.extend(final_safe_area_verification.get("warnings", []))

        background_annotation_verification = None
        if notes_result is not None:
            from catia_mcp.tools.smart_annotation_layout import (
                verify_annotations_within_sheet_safe_area,
            )
            background_annotation_verification = (
                verify_annotations_within_sheet_safe_area(
                    application,
                    sheet_index=None,
                    requested_sheet_margin_mm=sheet_margin_value,
                    inner_frame_inset_mm=INNER_FRAME_INSET_MM,
                    include_dimensions=False,
                    include_texts=True,
                    include_view_geometry=False,
                    attempt_repair=True,
                    strict_readback=True,
                    exclude_system_views=False,
                )
            )
            warnings.extend(
                background_annotation_verification.get("warnings", [])
            )
            if not background_annotation_verification.get("ok", False):
                raise DraftingOperationError(
                    "Engineering tolerance/GD&T note text crosses the configured "
                    "inner-frame safety area.",
                    data={
                        "final_safe_area_verification": final_safe_area_verification,
                        "background_annotation_verification": background_annotation_verification,
                    },
                )

        save_result = None
        if output_path:
            save_result = _verified_save_as(
                drawing_document,
                output_path,
                overwrite,
            )

        export_result = None
        if export_pdf_path:
            export_result = _verified_export(
                application,
                drawing_document,
                export_pdf_path,
                "pdf",
                overwrite,
            )

        return {
            "created": True,
            "creation_verified": True,
            "model_document": model_info,
            "drawing": _drawing_summary(
                conn,
                application,
                drawing_document,
            ),
            "sheet_configuration": sheet_configuration,
            "layout": {
                "front": [
                    layout.front_x,
                    layout.front_y,
                ],
                "top": [
                    layout.top_x,
                    layout.top_y,
                ],
                "right": [
                    layout.right_x,
                    layout.right_y,
                ],
                "title_block": [
                    layout.title_block_x,
                    layout.title_block_y,
                ],
                "notes": [
                    layout.notes_x,
                    layout.notes_y,
                ],
                "layout_basis": (
                    "effective inner-frame safe area derived from real paper size"
                ),
            },
            "views_count_before": views_before,
            "views_count_after": views_after,
            "views_added_count": (
                views_after - views_before
            ),
            "views": view_summaries,
            "view_verification": view_verification,
            "view_generation_waits": view_generation_waits,
            "generation_waits_verified": (
                generation_waits_verified
            ),
            "layout_resolution": layout_resolution,
            "layout_collision_free": (
                layout_resolution["verified"]
            ),
            "projected_view_positioning": {
                "alignment_mode": (
                    "reference_link_preserved_"
                    "manual_axis_alignment"
                ),
                "constraint_removed_with": (
                    "UnAlignedWithReferenceView"
                ),
                "position_write_primary_method": (
                    "SystemService.Evaluate."
                    "MCP_SetDrawingViewPosition"
                ),
                "global_group_fitting": {
                    "enabled": True,
                    "policy": (
                        "minimum_translation_to_sheet_margins"
                    ),
                    "views_translated_together": [
                        "front",
                        "top",
                        "right",
                    ],
                },
            },
            "minimum_view_gap_mm": minimum_gap_value,
            "requested_sheet_margin_mm": sheet_margin_value,
            "inner_frame_inset_mm": INNER_FRAME_INSET_MM,
            "effective_safe_margin_mm": effective_safe_margin,
            "safe_area": safe_area,
            "final_safe_area_verification": final_safe_area_verification,
            "background_annotation_verification": background_annotation_verification,
            "dimension_generation": dimension_generation,
            "title_block": title_block,
            "engineering_notes": notes_result,
            "update_attempts": update_attempts,
            "final_update_attempts": (
                final_update_attempts
            ),
            "save_result": save_result,
            "export_pdf_result": export_result,
            "model_modified": True,
            "document_save_required": save_result is None,
            "engineering_warning": (
                "Auto-generated dimensions and editable-text "
                "GD&T notes are drafting aids. Final manufacturing "
                "drawings require engineering verification."
            ),
        }, warnings
    except DraftingOperationError as exc:
        drawing_cleanup = _close_document(
            drawing_document
        )
        data = dict(exc.data or {})
        data["drawing_cleanup"] = drawing_cleanup
        raise DraftingOperationError(
            str(exc),
            data=data,
            warnings=[*warnings, *exc.warnings],
            status=(
                "partial_success"
                if drawing_cleanup.get("succeeded") is False
                else exc.status
            ),
        ) from exc
    except Exception as exc:
        drawing_cleanup = _close_document(
            drawing_document
        )
        raise DraftingOperationError(
            _format_com_error(exc),
            data={
                "drawing_created": drawing_document is not None,
                "drawing_cleanup": drawing_cleanup,
            },
            warnings=warnings,
            status=(
                "partial_success"
                if drawing_cleanup.get("succeeded") is False
                else "error"
            ),
        ) from exc



# ---------------------------------------------------------------------------
# Existing-template three-view integration (v8)
# ---------------------------------------------------------------------------

def _projection_method_from_existing_sheet(
    sheet: Any,
) -> dict[str, Any]:
    """Read the existing sheet projection without mutating template settings."""
    code = int(sheet.ProjectionMethod)
    mapping = {
        0: "first_angle",
        1: "third_angle",
    }
    method = mapping.get(code)
    if method is None:
        raise DraftingOperationError(
            "The active template sheet projection method could not be "
            f"classified (ProjectionMethod={code}).",
            data={
                "projection_method_code": code,
                "model_modified": False,
                "document_save_required": False,
            },
        )
    return {
        "actual_code": code,
        "method": method,
        "verified": True,
        "read_method": "DrawingSheet.ProjectionMethod",
        "mutated": False,
    }


def _delete_drawing_view(
    drawing_document: Any,
    sheet: Any,
    view: Any,
) -> dict[str, Any]:
    """Delete one added non-system DrawingView and verify the count delta."""
    before_count = int(sheet.Views.Count)
    selection = drawing_document.Selection
    attempts: list[dict[str, Any]] = []

    try:
        selection.Clear()
        selection.Add(view)
        selection.Delete()
        selection.Clear()
        after_count = int(sheet.Views.Count)
        verified = after_count == before_count - 1
        attempts.append({
            "method": "DrawingDocument.Selection.Delete",
            "succeeded": True,
            "views_count_before": before_count,
            "views_count_after": after_count,
            "verified": verified,
            "error": None,
        })
        if verified:
            return {
                "attempted": True,
                "succeeded": True,
                "selected_method": attempts[-1]["method"],
                "attempts": attempts,
            }
    except Exception as exc:
        try:
            selection.Clear()
        except Exception:
            pass
        attempts.append({
            "method": "DrawingDocument.Selection.Delete",
            "succeeded": False,
            "verified": False,
            "error": _format_com_error(exc),
        })

    name = safe_str(_safe_attribute(view, "Name", "")).strip()
    for key in (name or None, before_count):
        if key is None:
            continue
        try:
            count_before = int(sheet.Views.Count)
            sheet.Views.Remove(key)
            count_after = int(sheet.Views.Count)
            verified = count_after == count_before - 1
            attempts.append({
                "method": f"DrawingViews.Remove({key!r})",
                "succeeded": True,
                "views_count_before": count_before,
                "views_count_after": count_after,
                "verified": verified,
                "error": None,
            })
            if verified:
                return {
                    "attempted": True,
                    "succeeded": True,
                    "selected_method": attempts[-1]["method"],
                    "attempts": attempts,
                }
        except Exception as exc:
            attempts.append({
                "method": f"DrawingViews.Remove({key!r})",
                "succeeded": False,
                "verified": False,
                "error": _format_com_error(exc),
            })

    return {
        "attempted": True,
        "succeeded": False,
        "selected_method": None,
        "attempts": attempts,
    }


def _rollback_added_drawing_views(
    drawing_document: Any,
    sheet: Any,
    baseline_view_count: int,
    created_views: list[Any],
) -> dict[str, Any]:
    results: list[dict[str, Any]] = []
    for view in reversed(created_views):
        results.append(
            _delete_drawing_view(
                drawing_document,
                sheet,
                view,
            )
        )

    update_attempts, update_warnings = _call_update(
        sheet,
        drawing_document,
    )
    final_count = int(sheet.Views.Count)
    verified = bool(
        final_count == int(baseline_view_count)
        and all(item["succeeded"] for item in results)
    )
    return {
        "attempted": bool(created_views),
        "created_view_count": len(created_views),
        "results": results,
        "views_count_expected": int(baseline_view_count),
        "views_count_actual": final_count,
        "update_attempts": update_attempts,
        "warnings": update_warnings,
        "rollback_verified": verified,
    }


def _resolve_existing_drawing_document(
    conn: Any,
    application: Any,
    drawing_path: str,
) -> tuple[Any, dict[str, Any]]:
    raw = str(drawing_path or "").strip()
    if not raw:
        document = _require_active_drawing_document(conn)
        return document, {
            "selection_method": "active_CATDrawing",
            "path": _document_full_name(document),
            "opened_by_tool": False,
        }

    normalized = str(normalize_path(raw))
    if not normalized.lower().endswith(".catdrawing"):
        raise ValueError("drawing_path must point to a .CATDrawing file.")
    if not os.path.isfile(normalized):
        raise FileNotFoundError(
            f"Target template CATDrawing was not found: {normalized}"
        )

    document = _find_open_document_by_path(
        application,
        normalized,
    )
    opened_by_tool = document is None
    if document is None:
        document = application.Documents.Open(normalized)

    try:
        _ = int(document.Sheets.Count)
    except Exception as exc:
        if opened_by_tool:
            _close_document(document)
        raise CATIAError(
            f"Target document is not a CATDrawing: {normalized}"
        ) from exc

    return document, {
        "selection_method": (
            "opened_by_path" if opened_by_tool
            else "already_open_by_path"
        ),
        "path": normalized,
        "opened_by_tool": opened_by_tool,
    }


def _add_3view_to_existing_drawing_doc(
    conn: Any,
    application: Any,
    drawing_document: Any,
    model_document: Any,
    scale: float,
    generate_dimensions: bool,
    minimum_view_gap_mm: float,
    sheet_margin_mm: float,
    require_empty_model_views: bool,
    save_after_add: bool,
) -> tuple[dict[str, Any], list[str]]:
    """Add Front/Top/Right to an existing template CATDrawing.

    This operation deliberately does not create, replace or sanitise the
    CATDrawing file. On failure it rolls back only the newly added views.
    """
    warnings: list[str] = []
    created_views: list[Any] = []

    scale_value = _positive_float(scale, "scale")
    minimum_gap_value = _positive_float(
        minimum_view_gap_mm,
        "minimum_view_gap_mm",
    )
    sheet_margin_value = _positive_float(
        sheet_margin_mm,
        "sheet_margin_mm",
    )
    if not isinstance(require_empty_model_views, bool):
        raise ValueError(
            "require_empty_model_views must be a boolean."
        )
    if not isinstance(save_after_add, bool):
        raise ValueError("save_after_add must be a boolean.")

    model_info = conn.describe_document(model_document)
    if model_info.get("kind") not in ("CATPart", "CATProduct"):
        raise CATIAError(
            "model_path must resolve to a CATPart or CATProduct."
        )

    try:
        _ = int(drawing_document.Sheets.Count)
    except Exception as exc:
        raise CATIAError(
            "Target document must be an existing CATDrawing."
        ) from exc

    drawing_document.Activate()
    sheet = drawing_document.Sheets.ActiveSheet
    views_before = int(sheet.Views.Count)
    existing_non_system_views = max(0, views_before - 2)
    if require_empty_model_views and existing_non_system_views != 0:
        raise DraftingOperationError(
            "The template drawing already contains non-system views. "
            "No new view was created.",
            data={
                "views_count_before": views_before,
                "existing_non_system_view_count": (
                    existing_non_system_views
                ),
                "require_empty_model_views": True,
                "model_modified": False,
                "document_save_required": False,
            },
        )

    projection = _projection_method_from_existing_sheet(sheet)
    paper_dimensions = _get_paper_dimensions(
        application,
        sheet,
    )
    safe_area = _safe_area_descriptor(
        paper_dimensions["width_mm"],
        paper_dimensions["height_mm"],
        sheet_margin_value,
    )
    effective_safe_margin = safe_area["effective_safe_margin_mm"]
    layout = _calculate_layout(
        paper_dimensions["width_mm"],
        paper_dimensions["height_mm"],
        projection["method"],
        effective_safe_margin,
    )
    represented_object = _represented_3d_object(
        model_document
    )

    try:
        front_view = _add_front_view(
            sheet,
            model_document,
            represented_object,
            layout,
            scale_value,
        )
        created_views.append(front_view)

        top_view = _add_projected_view(
            application,
            sheet,
            front_view,
            model_document,
            represented_object,
            "Top view",
            "top",
            layout.top_x,
            layout.top_y,
            scale_value,
        )
        created_views.append(top_view)

        right_view = _add_projected_view(
            application,
            sheet,
            front_view,
            model_document,
            represented_object,
            "Right view",
            "right",
            layout.right_x,
            layout.right_y,
            scale_value,
        )
        created_views.append(right_view)

        update_attempts, update_warnings = _call_update(
            sheet,
            drawing_document,
        )
        warnings.extend(update_warnings)

        view_generation_waits = {
            "front": _wait_for_view_generation(
                application,
                drawing_document,
                sheet,
                front_view,
            ),
            "top": _wait_for_view_generation(
                application,
                drawing_document,
                sheet,
                top_view,
                position_callback=lambda: (
                    _position_projected_view(
                        application,
                        top_view,
                        front_view,
                        "top",
                        layout.top_x,
                        layout.top_y,
                    )
                ),
            ),
            "right": _wait_for_view_generation(
                application,
                drawing_document,
                sheet,
                right_view,
                position_callback=lambda: (
                    _position_projected_view(
                        application,
                        right_view,
                        front_view,
                        "right",
                        layout.right_x,
                        layout.right_y,
                    )
                ),
            ),
        }

        layout_resolution = _resolve_three_view_overlap(
            application,
            drawing_document,
            sheet,
            front_view,
            top_view,
            right_view,
            projection["method"],
            minimum_gap_value,
            paper_dimensions["width_mm"],
            paper_dimensions["height_mm"],
            effective_safe_margin,
        )

        views_after = int(sheet.Views.Count)
        view_summaries = [
            _view_summary(
                application,
                front_view,
                expected_type="front",
            ),
            _view_summary(
                application,
                top_view,
                expected_type="top",
            ),
            _view_summary(
                application,
                right_view,
                expected_type="right",
            ),
        ]
        view_verification = _verify_three_views(
            view_summaries
        )
        generation_waits_verified = all(
            item["verified"]
            for item in view_generation_waits.values()
        )
        creation_verified = bool(
            views_after == views_before + 3
            and view_verification["all_views_verified"]
            and generation_waits_verified
            and layout_resolution["verified"]
        )
        if not creation_verified:
            raise DraftingOperationError(
                "Three views were added to the template drawing, but "
                "their generation or layout could not be fully verified.",
                data={
                    "views_count_before": views_before,
                    "views_count_after": views_after,
                    "views": view_summaries,
                    "view_verification": view_verification,
                    "view_generation_waits": (
                        view_generation_waits
                    ),
                    "layout_resolution": layout_resolution,
                },
            )

        dimension_generation = {
            "requested": bool(generate_dimensions),
            "generation_call_succeeded": False,
            "generated": False,
            "generated_count": 0,
        }
        if generate_dimensions:
            (
                dimension_generation,
                dimension_warnings,
            ) = _generate_dimensions_internal(
                sheet,
                drawing_document,
            )
            dimension_generation["requested"] = True
            warnings.extend(dimension_warnings)

        final_update_attempts, final_update_warnings = _call_update(
            sheet,
            drawing_document,
        )
        warnings.extend(final_update_warnings)

        final_safe_area_verification = _verify_final_safe_area_annotations(
            application,
            sheet,
            sheet_margin_value,
            attempt_repair=True,
        )
        warnings.extend(final_safe_area_verification.get("warnings", []))

        save_result = {
            "requested": bool(save_after_add),
            "attempted": False,
            "succeeded": None,
            "document_saved": _document_saved(
                drawing_document
            ),
            "error": None,
        }
        if save_after_add:
            save_result["attempted"] = True
            try:
                drawing_document.Save()
                saved = _document_saved(
                    drawing_document
                ) is True
                save_result["succeeded"] = saved
                save_result["document_saved"] = saved
                if not saved:
                    raise RuntimeError(
                        "DrawingDocument.Save returned without Saved=true."
                    )
            except Exception as exc:
                save_result["succeeded"] = False
                save_result["error"] = _format_com_error(exc)
                raise DraftingOperationError(
                    "Views were created, but the existing template "
                    "CATDrawing could not be saved.",
                    data={
                        "save_result": save_result,
                        "views_count_before": views_before,
                        "views_count_after": int(
                            sheet.Views.Count
                        ),
                    },
                ) from exc

        return {
            "operation": (
                "catia_add_3view_to_existing_drawing"
            ),
            "created": True,
            "creation_verified": True,
            "integration_mode": (
                "existing_template_drawing_in_place"
            ),
            "drawing_recreated": False,
            "template_sanitisation_called": False,
            "sheet_settings_mutated": False,
            "model_document": model_info,
            "drawing": _drawing_summary(
                conn,
                application,
                drawing_document,
            ),
            "sheet_projection": projection,
            "paper_dimensions": paper_dimensions,
            "scale": scale_value,
            "requested_sheet_margin_mm": sheet_margin_value,
            "inner_frame_inset_mm": INNER_FRAME_INSET_MM,
            "effective_safe_margin_mm": effective_safe_margin,
            "safe_area": safe_area,
            "final_safe_area_verification": final_safe_area_verification,
            "layout": {
                "front": [layout.front_x, layout.front_y],
                "top": [layout.top_x, layout.top_y],
                "right": [layout.right_x, layout.right_y],
                "layout_basis": (
                    "existing template paper size/projection method "
                    "plus the effective inner-frame safe area"
                ),
            },
            "views_count_before": views_before,
            "views_count_after": int(sheet.Views.Count),
            "views_added_count": 3,
            "views": view_summaries,
            "view_verification": view_verification,
            "view_generation_waits": (
                view_generation_waits
            ),
            "layout_resolution": layout_resolution,
            "dimension_generation": dimension_generation,
            "update_attempts": update_attempts,
            "final_update_attempts": (
                final_update_attempts
            ),
            "save_result": save_result,
            "rollback": {
                "attempted": False,
                "rollback_verified": True,
            },
            "model_modified": True,
            "document_save_required": (
                not bool(save_after_add)
            ),
        }, warnings

    except Exception as exc:
        rollback = _rollback_added_drawing_views(
            drawing_document,
            sheet,
            views_before,
            created_views,
        )
        warnings.extend(rollback.get("warnings", []))
        data = dict(
            getattr(exc, "data", None) or {}
        )
        data.update({
            "operation": (
                "catia_add_3view_to_existing_drawing"
            ),
            "integration_mode": (
                "existing_template_drawing_in_place"
            ),
            "views_count_before": views_before,
            "created_view_count_before_failure": len(
                created_views
            ),
            "rollback": rollback,
            "drawing_file_replaced": False,
            "template_sanitisation_called": False,
            "model_modified": not rollback[
                "rollback_verified"
            ],
            "document_save_required": not rollback[
                "rollback_verified"
            ],
        })
        status = (
            "error"
            if rollback["rollback_verified"]
            else "partial_success"
        )
        if isinstance(exc, DraftingOperationError):
            raise DraftingOperationError(
                str(exc),
                data=data,
                warnings=[*warnings, *exc.warnings],
                status=status,
            ) from exc
        raise DraftingOperationError(
            _format_com_error(exc),
            data=data,
            warnings=warnings,
            status=status,
        ) from exc



# ---------------------------------------------------------------------------
# MCP registration
# ---------------------------------------------------------------------------

def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = ctx.conn
    names: list[str] = []

    @mcp.tool()
    def catia_create_3view_drawing_from_file(
        model_path: str,
        paper_size: str = "A3",
        orientation: str = "landscape",
        projection_method: str = "third_angle",
        scale: float = 1.0,
        drawing_title: str = "",
        part_number: str = "",
        material: str = "",
        general_tolerance: str = (
            "ISO 2768-mK unless otherwise specified"
        ),
        generate_dimensions: bool = False,
        dimension_notes: list[str] | None = None,
        tolerance_notes: list[str] | None = None,
        gdt_notes: list[str] | None = None,
        output_path: str = "",
        export_pdf_path: str = "",
        overwrite: bool = False,
        minimum_view_gap_mm: float = 15.0,
        sheet_margin_mm: float = 5.0,
    ) -> dict[str, Any]:
        """Create a verified Front/Top/Right generative drawing.

        Top and Right are true projection views linked to Front.
        projection_method may be first_angle or third_angle.
        """

        warnings: list[str] = []
        model_document = None
        source_opened_by_tool = False
        normalized_model_path: Optional[str] = None

        try:
            application = conn.connect(visible=True)
            normalized_model_path = _normalise_model_path(
                model_path
            )
            model_document = _find_open_document_by_path(
                application,
                normalized_model_path,
            )

            if model_document is None:
                model_document = application.Documents.Open(
                    normalized_model_path
                )
                source_opened_by_tool = True

            data, create_warnings = (
                _create_3view_drawing_from_model_doc(
                    conn=conn,
                    application=application,
                    model_document=model_document,
                    paper_size=paper_size,
                    orientation=orientation,
                    projection_method=projection_method,
                    scale=scale,
                    drawing_title=drawing_title,
                    part_number=part_number,
                    material=material,
                    general_tolerance=general_tolerance,
                    generate_dimensions=generate_dimensions,
                    dimension_notes=dimension_notes or [],
                    tolerance_notes=tolerance_notes or [],
                    gdt_notes=gdt_notes or [],
                    output_path=output_path,
                    export_pdf_path=export_pdf_path,
                    overwrite=bool(overwrite),
                    minimum_view_gap_mm=minimum_view_gap_mm,
                    sheet_margin_mm=sheet_margin_mm,
                )
            )
            warnings.extend(create_warnings)
            data["source_document_lifecycle"] = {
                "path": normalized_model_path,
                "opened_by_tool": source_opened_by_tool,
                "left_open_intentionally": True,
                "reason": (
                    "The generative drawing keeps a loaded source "
                    "document for immediate updates."
                ),
            }
            return _success(data, warnings)
        except DraftingOperationError as exc:
            source_cleanup = None
            if source_opened_by_tool:
                source_cleanup = _close_document(
                    model_document
                )
            data = dict(exc.data or {})
            data["source_document_lifecycle"] = {
                "path": normalized_model_path,
                "opened_by_tool": source_opened_by_tool,
                "cleanup_on_failure": source_cleanup,
            }
            return _error(
                str(exc),
                data=data,
                warnings=[*warnings, *exc.warnings],
                status=exc.status,
            )
        except Exception as exc:
            source_cleanup = None
            if source_opened_by_tool:
                source_cleanup = _close_document(
                    model_document
                )
            return _error(
                _format_com_error(exc),
                data={
                    "source_document_lifecycle": {
                        "path": normalized_model_path,
                        "opened_by_tool": (
                            source_opened_by_tool
                        ),
                        "cleanup_on_failure": (
                            source_cleanup
                        ),
                    },
                },
                warnings=warnings,
            )

    names.append("catia_create_3view_drawing_from_file")

    @mcp.tool()
    def catia_create_3view_drawing_from_active_model(
        paper_size: str = "A3",
        orientation: str = "landscape",
        projection_method: str = "third_angle",
        scale: float = 1.0,
        drawing_title: str = "",
        part_number: str = "",
        material: str = "",
        general_tolerance: str = (
            "ISO 2768-mK unless otherwise specified"
        ),
        generate_dimensions: bool = False,
        dimension_notes: list[str] | None = None,
        tolerance_notes: list[str] | None = None,
        gdt_notes: list[str] | None = None,
        output_path: str = "",
        export_pdf_path: str = "",
        overwrite: bool = False,
        minimum_view_gap_mm: float = 15.0,
        sheet_margin_mm: float = 5.0,
    ) -> dict[str, Any]:
        """Create a verified projected three-view drawing from the active model."""

        try:
            application = conn.connect(visible=True)
            model_document = application.ActiveDocument
            data, warnings = _create_3view_drawing_from_model_doc(
                conn=conn,
                application=application,
                model_document=model_document,
                paper_size=paper_size,
                orientation=orientation,
                projection_method=projection_method,
                scale=scale,
                drawing_title=drawing_title,
                part_number=part_number,
                material=material,
                general_tolerance=general_tolerance,
                generate_dimensions=generate_dimensions,
                dimension_notes=dimension_notes or [],
                tolerance_notes=tolerance_notes or [],
                gdt_notes=gdt_notes or [],
                output_path=output_path,
                export_pdf_path=export_pdf_path,
                overwrite=bool(overwrite),
                minimum_view_gap_mm=minimum_view_gap_mm,
                sheet_margin_mm=sheet_margin_mm,
            )
            data["source_document_lifecycle"] = {
                "opened_by_tool": False,
                "left_open_intentionally": True,
                "reason": "The caller's active model is preserved.",
            }
            return _success(data, warnings)
        except DraftingOperationError as exc:
            return _error(
                str(exc),
                data=exc.data,
                warnings=exc.warnings,
                status=exc.status,
            )
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("catia_create_3view_drawing_from_active_model")

    @mcp.tool()
    def catia_add_3view_to_existing_drawing(
        model_path: str,
        drawing_path: str = "",
        scale: float = 0.5,
        generate_dimensions: bool = False,
        minimum_view_gap_mm: float = 20.0,
        sheet_margin_mm: float = 15.0,
        require_empty_model_views: bool = True,
        save_after_add: bool = True,
    ) -> dict[str, Any]:
        """Add verified Front/Top/Right views to an existing template drawing.

        Unlike catia_create_3view_drawing_from_file, this tool never creates a
        blank Drawing document and never reapplies or sanitises the template.
        It reads the target sheet's existing paper size and projection method.
        On failure, only views created by this call are rolled back.
        """
        model_document = None
        drawing_document = None
        model_opened_by_tool = False
        drawing_lifecycle: dict[str, Any] = {}
        normalized_model_path: Optional[str] = None

        try:
            application = conn.connect(visible=True)
            drawing_document, drawing_lifecycle = (
                _resolve_existing_drawing_document(
                    conn,
                    application,
                    drawing_path,
                )
            )

            normalized_model_path = _normalise_model_path(
                model_path
            )
            model_document = _find_open_document_by_path(
                application,
                normalized_model_path,
            )
            if model_document is None:
                model_document = application.Documents.Open(
                    normalized_model_path
                )
                model_opened_by_tool = True

            drawing_document.Activate()
            data, warnings = (
                _add_3view_to_existing_drawing_doc(
                    conn=conn,
                    application=application,
                    drawing_document=drawing_document,
                    model_document=model_document,
                    scale=scale,
                    generate_dimensions=bool(
                        generate_dimensions
                    ),
                    minimum_view_gap_mm=(
                        minimum_view_gap_mm
                    ),
                    sheet_margin_mm=sheet_margin_mm,
                    require_empty_model_views=bool(
                        require_empty_model_views
                    ),
                    save_after_add=bool(save_after_add),
                )
            )
            data["drawing_document_lifecycle"] = {
                **drawing_lifecycle,
                "left_open_intentionally": True,
                "reason": (
                    "The template drawing remains open for "
                    "annotation and export."
                ),
            }
            data["source_document_lifecycle"] = {
                "path": normalized_model_path,
                "opened_by_tool": model_opened_by_tool,
                "left_open_intentionally": True,
                "reason": (
                    "Generative views keep the source model "
                    "loaded for update."
                ),
            }
            return _success(data, warnings)

        except DraftingOperationError as exc:
            source_cleanup = None
            if model_opened_by_tool:
                source_cleanup = _close_document(
                    model_document
                )
            data = dict(exc.data or {})
            data["drawing_document_lifecycle"] = (
                drawing_lifecycle
            )
            data["source_document_lifecycle"] = {
                "path": normalized_model_path,
                "opened_by_tool": model_opened_by_tool,
                "cleanup_on_failure": source_cleanup,
            }
            return _error(
                str(exc),
                data=data,
                warnings=exc.warnings,
                status=exc.status,
            )
        except Exception as exc:
            source_cleanup = None
            if model_opened_by_tool:
                source_cleanup = _close_document(
                    model_document
                )
            return _error(
                _format_com_error(exc),
                data={
                    "drawing_document_lifecycle": (
                        drawing_lifecycle
                    ),
                    "source_document_lifecycle": {
                        "path": normalized_model_path,
                        "opened_by_tool": (
                            model_opened_by_tool
                        ),
                        "cleanup_on_failure": source_cleanup,
                    },
                },
            )

    names.append("catia_add_3view_to_existing_drawing")

    @mcp.tool()
    def catia_generate_drawing_dimensions() -> dict[str, Any]:
        """Generate and verify dimensions from eligible 3D constraints."""

        try:
            application = conn.connect(visible=True)
            drawing_document = _require_active_drawing_document(
                conn
            )
            sheet = drawing_document.Sheets.ActiveSheet
            data, warnings = _generate_dimensions_internal(
                sheet,
                drawing_document,
            )
            data["drawing"] = _drawing_summary(
                conn,
                application,
                drawing_document,
            )
            return _success(data, warnings)
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("catia_generate_drawing_dimensions")

    @mcp.tool()
    def catia_add_drawing_text(
        text: str,
        x: float = 20.0,
        y: float = 20.0,
        font_size: float = 3.5,
        use_background_view: bool = False,
    ) -> dict[str, Any]:
        """Add editable drawing text and verify the object-count delta."""

        try:
            application = conn.connect(visible=True)
            drawing_document = _require_active_drawing_document(
                conn
            )
            sheet = drawing_document.Sheets.ActiveSheet
            data, warnings = _add_text_internal(
                application,
                drawing_document,
                sheet,
                text,
                x,
                y,
                font_size,
                bool(use_background_view),
            )
            return _success(data, warnings)
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("catia_add_drawing_text")

    @mcp.tool()
    def catia_add_title_block(
        title: str = "",
        part_number: str = "",
        material: str = "",
        general_tolerance: str = "",
        drawn_by: str = "CATIA MCP",
        x: Optional[float] = None,
        y: Optional[float] = None,
        sheet_margin_mm: float = DEFAULT_TITLE_BLOCK_MARGIN_MM,
    ) -> dict[str, Any]:
        """Add a verified paper-aware title block in Background View.

        When x or y is omitted, that coordinate is resolved from the active
        sheet dimensions. Explicit coordinates are preserved but rejected if
        the complete 90 x 48 mm table would cross the requested sheet margin.
        """

        try:
            application = conn.connect(visible=True)
            drawing_document = _require_active_drawing_document(
                conn
            )
            sheet = drawing_document.Sheets.ActiveSheet
            data, warnings = _add_title_block_internal(
                application,
                drawing_document,
                sheet,
                title,
                part_number,
                material,
                general_tolerance,
                drawn_by,
                x,
                y,
                sheet_margin_mm,
            )
            return _success(data, warnings)
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("catia_add_title_block")

    @mcp.tool()
    def catia_add_engineering_notes(
        dimension_notes: list[str] | None = None,
        tolerance_notes: list[str] | None = None,
        gdt_notes: list[str] | None = None,
        x: float = 20.0,
        y: float = 35.0,
        font_size: float = 3.2,
    ) -> dict[str, Any]:
        """Add verified multiline engineering notes to Background View."""

        try:
            application = conn.connect(visible=True)
            drawing_document = _require_active_drawing_document(
                conn
            )
            sheet = drawing_document.Sheets.ActiveSheet
            lines = _engineering_note_lines(
                dimension_notes or [],
                tolerance_notes or [],
                gdt_notes or [],
            )
            data, warnings = _add_text_internal(
                application,
                drawing_document,
                sheet,
                "\n".join(lines),
                x,
                y,
                font_size,
                True,
            )
            data["note_lines"] = lines
            data["annotation_kind"] = (
                "editable_multiline_drawing_text"
            )
            return _success(data, warnings)
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("catia_add_engineering_notes")

    @mcp.tool()
    def catia_add_gdt_note(
        feature: str,
        control: str,
        tolerance: str,
        datum: str = "",
        x: float = 20.0,
        y: float = 70.0,
    ) -> dict[str, Any]:
        """Add a human-readable GD&T note as editable text.

        This is not a semantic CATIA feature-control-frame object.
        """

        try:
            feature_value = _nonempty_text(feature, "feature")
            control_value = _nonempty_text(control, "control")
            tolerance_value = _nonempty_text(
                tolerance,
                "tolerance",
            )
            text = (
                f"GD&T: {feature_value} | "
                f"{control_value} | {tolerance_value}"
            )
            if str(datum).strip():
                text += f" | Datum: {str(datum).strip()}"

            application = conn.connect(visible=True)
            drawing_document = _require_active_drawing_document(
                conn
            )
            sheet = drawing_document.Sheets.ActiveSheet
            data, warnings = _add_text_internal(
                application,
                drawing_document,
                sheet,
                text,
                x,
                y,
                3.5,
                True,
            )
            warnings.append(
                "This tool creates editable plain text, not a "
                "semantic CATIA GD&T feature-control frame."
            )
            data["annotation_kind"] = (
                "plain_text_gdt_note"
            )
            data["semantic_gdt_object_created"] = False
            return _success(data, warnings)
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("catia_add_gdt_note")

    @mcp.tool()
    def catia_save_active_drawing_as(
        path: str,
        overwrite: bool = False,
    ) -> dict[str, Any]:
        """Save the active CATDrawing and verify the output file."""

        try:
            application = conn.connect(visible=True)
            drawing_document = _require_active_drawing_document(
                conn
            )
            data = _verified_save_as(
                drawing_document,
                path,
                bool(overwrite),
            )
            data["drawing"] = _drawing_summary(
                conn,
                application,
                drawing_document,
            )
            return _success(data)
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("catia_save_active_drawing_as")

    @mcp.tool()
    def catia_export_active_drawing(
        path: str,
        format_name: str = "pdf",
        overwrite: bool = False,
    ) -> dict[str, Any]:
        """Export the active CATDrawing to PDF, DWG or DXF and verify it."""

        try:
            drawing_document = _require_active_drawing_document(
                conn
            )
            application = conn.connect(visible=True)
            data = _verified_export(
                application,
                drawing_document,
                path,
                format_name,
                bool(overwrite),
            )
            return _success(data)
        except DraftingOperationError as exc:
            return _error(
                str(exc),
                data=exc.data,
                warnings=exc.warnings,
                status=exc.status,
            )
        except Exception as exc:
            return _error(_format_com_error(exc))

    names.append("catia_export_active_drawing")

    return names


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
