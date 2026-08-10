"""Robust CATIA V5 export, screen capture and active-viewer MCP tools.

The module treats CATIA file generation as a transaction:

* the target extension and source-document type are checked before CATIA writes;
* an existing target is moved to a temporary sibling backup when overwrite=True;
* CATIA output must become non-empty and stable before it is accepted;
* supported file signatures are verified;
* failures remove partial output and restore the previous target;
* source-document state and viewer operations are reported explicitly;
* long IGES extensions are exported through a verified temporary ``.igs`` path
  and atomically finalized to the requested ``.iges`` target.

CATIA V5 ``Document.ExportData`` accepts a full output path and a format token.
The token must be supported by the active document type, installed licenses and
the current CATIA configuration.  This module intentionally limits automatic
exports to the Part/Product/Drawing combinations covered by its validation
matrix.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import hashlib
import math
import os
from pathlib import Path
import struct
import time
from typing import Any, Callable, Optional
import uuid

from catia_mcp.connection import CATIAError, normalize_path


IMPLEMENTATION_VERSION = "export-fixed-2026-08-05-v3"
_CATVB_SCRIPT_LANGUAGE = 1
_DEFAULT_TIMEOUT_SECONDS = 30.0
_STABLE_POLL_INTERVAL_SECONDS = 0.10
_STABLE_POLL_COUNT = 4
_VECTOR_TOLERANCE = 1.0e-5


@dataclass(frozen=True)
class ExportFormat:
    canonical_name: str
    catia_token: str
    extensions: tuple[str, ...]
    allowed_document_kinds: frozenset[str]


_PART = "CATPart"
_PRODUCT = "CATProduct"
_DRAWING = "CATDrawing"

_EXPORT_FORMATS = (
    ExportFormat("step", "stp", (".stp", ".step"), frozenset({_PART, _PRODUCT})),
    ExportFormat(
        "iges", "igs", (".igs", ".iges"),
        frozenset({_PART, _PRODUCT, _DRAWING}),
    ),
    ExportFormat("stl", "stl", (".stl",), frozenset({_PART})),
    ExportFormat(
        "3dxml", "3dxml", (".3dxml",), frozenset({_PART, _PRODUCT}),
    ),
    ExportFormat("vrml", "wrl", (".wrl", ".vrml"), frozenset({_PART, _PRODUCT})),
    ExportFormat("cgr", "cgr", (".cgr",), frozenset({_PART, _PRODUCT})),
    ExportFormat("pdf", "pdf", (".pdf",), frozenset({_DRAWING})),
    ExportFormat("dwg", "dwg", (".dwg",), frozenset({_DRAWING})),
    ExportFormat("dxf", "dxf", (".dxf",), frozenset({_DRAWING})),
)

_FORMAT_ALIASES: dict[str, ExportFormat] = {}
_FORMAT_BY_EXTENSION: dict[str, ExportFormat] = {}
for _format in _EXPORT_FORMATS:
    _FORMAT_ALIASES[_format.canonical_name] = _format
    _FORMAT_ALIASES[_format.catia_token] = _format
    for _extension in _format.extensions:
        _FORMAT_ALIASES[_extension.lstrip(".")] = _format
        _FORMAT_BY_EXTENSION[_extension] = _format

# CatCaptureFormat values from the CATIA Viewer Automation interface.
CAPTURE_FORMATS: dict[str, int] = {
    ".cgm": 0,   # catCaptureFormatCGM
    ".emf": 1,   # catCaptureFormatEMF
    ".tif": 2,   # catCaptureFormatTIFF
    ".tiff": 2,
    ".bmp": 4,   # catCaptureFormatBMP
    ".jpg": 5,   # catCaptureFormatJPEG
    ".jpeg": 5,
}


class ExportOperationError(RuntimeError):
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


@dataclass
class _TargetTransaction:
    path: str
    overwrite: bool
    existed_before: bool = False
    original_size_bytes: Optional[int] = None
    original_sha256: Optional[str] = None
    backup_path: Optional[str] = None
    prepared: bool = False
    committed: bool = False
    rollback_attempted: bool = False
    rollback_succeeded: Optional[bool] = None
    rollback_error: Optional[str] = None
    cleanup_warnings: list[str] = field(default_factory=list)
    restored_size_bytes: Optional[int] = None
    restored_sha256: Optional[str] = None
    target_exists_after_rollback: Optional[bool] = None
    backup_exists_after_completion: Optional[bool] = None

    def prepare(self) -> None:
        if not isinstance(self.overwrite, bool):
            raise ValueError("overwrite must be a boolean.")
        target = Path(self.path)
        target.parent.mkdir(parents=True, exist_ok=True)
        self.existed_before = target.exists()
        if self.existed_before and not target.is_file():
            raise IsADirectoryError(f"Target is not a regular file: {self.path}")
        if self.existed_before and not self.overwrite:
            raise FileExistsError(
                f"Target already exists: {self.path}. "
                "Set overwrite=true to replace it."
            )
        if self.existed_before:
            self.original_size_bytes = target.stat().st_size
            self.original_sha256 = _sha256_file(self.path)
            backup_name = (
                f".{target.name}.mcp_export_backup_{uuid.uuid4().hex}"
            )
            backup = target.with_name(backup_name)
            os.replace(self.path, str(backup))
            self.backup_path = str(backup)
        self.prepared = True

    def commit(self) -> None:
        if self.backup_path and os.path.isfile(self.backup_path):
            try:
                os.unlink(self.backup_path)
            except OSError as exc:
                self.cleanup_warnings.append(
                    f"Could not delete overwrite backup '{self.backup_path}': {exc}"
                )
        self.backup_exists_after_completion = bool(
            self.backup_path and os.path.exists(self.backup_path)
        )
        self.committed = True

    def rollback(self) -> None:
        self.rollback_attempted = True
        errors: list[str] = []
        try:
            if os.path.isfile(self.path):
                os.unlink(self.path)
        except OSError as exc:
            errors.append(f"partial target cleanup failed: {exc}")
        if self.backup_path and os.path.isfile(self.backup_path):
            try:
                os.replace(self.backup_path, self.path)
            except OSError as exc:
                errors.append(f"original target restore failed: {exc}")

        self.target_exists_after_rollback = os.path.exists(self.path)
        if os.path.isfile(self.path):
            try:
                self.restored_size_bytes = os.path.getsize(self.path)
            except OSError:
                self.restored_size_bytes = None
            self.restored_sha256 = _safe_sha256(self.path)
        else:
            self.restored_size_bytes = None
            self.restored_sha256 = None

        if self.existed_before:
            restored = (
                os.path.isfile(self.path)
                and self.original_sha256 is not None
                and self.restored_sha256 == self.original_sha256
            )
        else:
            restored = not os.path.exists(self.path)
        if errors:
            restored = False
            self.rollback_error = "; ".join(errors)
        self.rollback_succeeded = restored
        self.backup_exists_after_completion = bool(
            self.backup_path and os.path.exists(self.backup_path)
        )

    def as_dict(self) -> dict[str, Any]:
        return {
            "path": self.path,
            "overwrite_requested": self.overwrite,
            "existed_before": self.existed_before,
            "original_size_bytes": self.original_size_bytes,
            "original_sha256": self.original_sha256,
            "backup_created": bool(self.backup_path),
            "backup_path": self.backup_path,
            "prepared": self.prepared,
            "committed": self.committed,
            "rollback_attempted": self.rollback_attempted,
            "rollback_succeeded": self.rollback_succeeded,
            "rollback_error": self.rollback_error,
            "restored_size_bytes": self.restored_size_bytes,
            "restored_sha256": self.restored_sha256,
            "target_exists_after_rollback": self.target_exists_after_rollback,
            "backup_exists_after_completion": self.backup_exists_after_completion,
            "cleanup_warnings": list(self.cleanup_warnings),
        }


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
    if isinstance(exc, ExportOperationError):
        if data is None:
            data = exc.data
        warning_list = list(exc.warnings)
        warning_list.extend(warnings or [])
        resolved_status = status or exc.status
    else:
        warning_list = list(warnings or [])
        resolved_status = status or "error"
    result: dict[str, Any] = {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": False,
        "status": resolved_status,
        "operation": operation,
        "error": str(exc),
        "error_type": type(exc).__name__,
        "warnings": warning_list,
    }
    if data is not None:
        result["data"] = data
    if operation in {"catia_screenshot", "catia_set_view", "catia_fit_all"}:
        result["capability"] = "CATIA V5 ActiveWindow.ActiveViewer"
    return result


def _format_com_error(exc: BaseException) -> str:
    details = getattr(exc, "excepinfo", None)
    if details and len(details) >= 3 and details[2]:
        return str(details[2])
    hresult = getattr(exc, "hresult", None)
    if hresult is not None:
        return f"{exc} (HRESULT 0x{int(hresult) & 0xFFFFFFFF:08X})"
    return str(exc)


def _safe_attr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _safe_bool(value: Any) -> Optional[bool]:
    if value is None:
        return None
    try:
        return bool(value)
    except Exception:
        return None


def _nonempty_text(value: Any, parameter_name: str) -> str:
    text = str(value).strip()
    if not text:
        raise ValueError(f"{parameter_name} cannot be empty.")
    return text


def _normalised_file_path(value: Any) -> str:
    raw_path = _nonempty_text(value, "file_path")
    normalised = str(normalize_path(raw_path)).strip()
    if not normalised:
        raise ValueError("file_path could not be normalised.")
    return normalised


def _positive_timeout(value: Any) -> float:
    if isinstance(value, bool):
        raise ValueError("timeout_seconds must be a finite number.")
    try:
        timeout = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError("timeout_seconds must be a finite number.") from exc
    if not math.isfinite(timeout) or timeout < 1.0 or timeout > 300.0:
        raise ValueError("timeout_seconds must be in the range 1..300.")
    return timeout


def _sha256_file(file_path: str) -> str:
    digest = hashlib.sha256()
    with open(file_path, "rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_sha256(file_path: str) -> Optional[str]:
    try:
        return _sha256_file(file_path)
    except OSError:
        return None


def _same_path(left: str, right: str) -> bool:
    try:
        return os.path.normcase(os.path.abspath(left)) == os.path.normcase(
            os.path.abspath(right)
        )
    except Exception:
        return False


def _application(conn: Any) -> tuple[Any, dict[str, Any]]:
    attempts: list[dict[str, Any]] = []
    try:
        application = conn.connect(visible=True)
        _ = application.Documents
        attempts.append(
            {
                "method": "conn.connect(visible=True)",
                "succeeded": True,
                "error": None,
            }
        )
        return application, {
            "resolved": True,
            "method": "conn.connect(visible=True)",
            "attempts": attempts,
        }
    except Exception as exc:
        attempts.append(
            {
                "method": "conn.connect(visible=True)",
                "succeeded": False,
                "error": _format_com_error(exc),
            }
        )

    for attribute_name in ("app", "application", "catia"):
        try:
            candidate = getattr(conn, attribute_name)
            application = candidate() if callable(candidate) else candidate
            if application is None:
                raise RuntimeError("candidate is None")
            _ = application.Documents
            attempts.append(
                {
                    "method": f"conn.{attribute_name}",
                    "succeeded": True,
                    "error": None,
                }
            )
            return application, {
                "resolved": True,
                "method": f"conn.{attribute_name}",
                "attempts": attempts,
                "compatibility_fallback_used": True,
            }
        except Exception as exc:
            attempts.append(
                {
                    "method": f"conn.{attribute_name}",
                    "succeeded": False,
                    "error": _format_com_error(exc),
                }
            )
    raise ExportOperationError(
        "Cannot resolve the live CATIA Application.",
        data={
            "failure_stage": "A_application_resolution",
            "application_resolution": {
                "resolved": False,
                "method": None,
                "attempts": attempts,
            },
        },
    )


def _active_document(application: Any) -> Any:
    try:
        document = application.ActiveDocument
    except Exception as exc:
        raise CATIAError("CATIA has no accessible active document.") from exc
    if document is None:
        raise CATIAError("No active CATIA document is available.")
    return document


def _document_kind(document: Any) -> str:
    name = str(_safe_attr(document, "Name", "")).strip()
    suffix = Path(name).suffix.lower()
    if suffix == ".catpart":
        return _PART
    if suffix == ".catproduct":
        return _PRODUCT
    if suffix == ".catdrawing":
        return _DRAWING

    # Unsaved and custom-named documents may not expose the native extension.
    try:
        _ = document.Sheets
        return _DRAWING
    except Exception:
        pass
    try:
        _ = document.Part
        return _PART
    except Exception:
        pass
    try:
        _ = document.Product
        return _PRODUCT
    except Exception:
        pass
    return "Unknown"


def _document_info(document: Any) -> dict[str, Any]:
    name = str(_safe_attr(document, "Name", ""))
    full_name = str(_safe_attr(document, "FullName", ""))
    return {
        "name": name,
        "full_name": full_name,
        "kind": _document_kind(document),
        "saved": _safe_bool(_safe_attr(document, "Saved", None)),
        "read_only": _safe_bool(_safe_attr(document, "ReadOnly", None)),
    }


def _active_viewer(
    application: Any,
    *,
    require_3d: bool = False,
) -> tuple[Any, dict[str, Any]]:
    try:
        window = application.ActiveWindow
        viewer = window.ActiveViewer
    except Exception as exc:
        raise CATIAError(
            "No active CATIA viewer is available. Open and activate a document window."
        ) from exc
    try:
        viewer.Activate()
        activated = True
    except Exception:
        activated = False
    viewer_type = type(viewer).__name__
    has_viewpoint_3d = False
    try:
        _ = viewer.Viewpoint3D
        has_viewpoint_3d = True
    except Exception:
        pass
    if require_3d and not has_viewpoint_3d:
        raise CATIAError(
            "The active viewer is not a 3D viewer and has no Viewpoint3D."
        )
    return viewer, {
        "viewer_type": viewer_type,
        "viewer_width_px": _safe_attr(viewer, "Width"),
        "viewer_height_px": _safe_attr(viewer, "Height"),
        "viewer_activated": activated,
        "viewpoint_3d_available": has_viewpoint_3d,
    }


def _resolve_export_format(file_path: str, format_name: str) -> ExportFormat:
    suffix = Path(file_path).suffix.lower()
    requested = str(format_name).strip().lower().lstrip(".")
    if requested:
        export_format = _FORMAT_ALIASES.get(requested)
        if export_format is None:
            supported = ", ".join(item.canonical_name for item in _EXPORT_FORMATS)
            raise ValueError(
                f"Unsupported export format '{format_name}'. Supported: {supported}."
            )
        if suffix not in export_format.extensions:
            expected = ", ".join(export_format.extensions)
            raise ValueError(
                f"file_path extension '{suffix or '<none>'}' does not match "
                f"format '{export_format.canonical_name}' (expected {expected})."
            )
        return export_format
    export_format = _FORMAT_BY_EXTENSION.get(suffix)
    if export_format is None:
        supported_extensions = ", ".join(sorted(_FORMAT_BY_EXTENSION))
        raise ValueError(
            f"Cannot infer export format from extension '{suffix or '<none>'}'. "
            f"Supported extensions: {supported_extensions}."
        )
    return export_format


def _plan_export_path(
    requested_path: str,
    export_format: ExportFormat,
) -> tuple[str, dict[str, Any]]:
    """Return the CATIA write path and an explicit finalization plan.

    CATIA V5 R26 has been observed to materialise an ``.igs`` sibling when
    ``ExportData`` is asked for an ``.iges`` path.  Waiting on the long
    extension therefore times out even though CATIA produced a valid file.
    Export through a unique, same-directory ``.igs`` staging path and rename
    it atomically after validation.
    """
    suffix = Path(requested_path).suffix.lower()
    if export_format.canonical_name == "iges" and suffix == ".iges":
        requested = Path(requested_path)
        staging_name = (
            f"{requested.stem}.mcp_iges_stage_{uuid.uuid4().hex}.igs"
        )
        actual_path = str(requested.with_name(staging_name))
        return actual_path, {
            "strategy": "temporary_igs_then_atomic_rename",
            "requested_path": requested_path,
            "actual_export_path": actual_path,
            "requested_extension": suffix,
            "actual_export_extension": ".igs",
            "temporary_path_used": True,
            "reason": (
                "CATIA may force the IGES output extension to .igs; the "
                "temporary path prevents a false wait timeout for .iges."
            ),
        }
    return requested_path, {
        "strategy": "direct_to_requested_path",
        "requested_path": requested_path,
        "actual_export_path": requested_path,
        "requested_extension": suffix,
        "actual_export_extension": suffix,
        "temporary_path_used": False,
        "reason": None,
    }


def _cleanup_staging_path(file_path: str, requested_path: str) -> dict[str, Any]:
    attempted = not _same_path(file_path, requested_path)
    error = None
    if attempted and os.path.exists(file_path):
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            else:
                error = "staging path exists but is not a regular file"
        except OSError as exc:
            error = str(exc)
    return {
        "attempted": attempted,
        "path": file_path if attempted else None,
        "exists_after_cleanup": os.path.exists(file_path) if attempted else False,
        "succeeded": (not attempted) or (not os.path.exists(file_path) and error is None),
        "error": error,
    }


def _finalize_staged_export(
    actual_export_path: str,
    requested_path: str,
) -> dict[str, Any]:
    if _same_path(actual_export_path, requested_path):
        return {
            "required": False,
            "attempted": False,
            "succeeded": True,
            "method": "direct_output",
            "error": None,
        }
    try:
        os.replace(actual_export_path, requested_path)
    except OSError as exc:
        raise ExportOperationError(
            f"Validated staged export could not be finalized: {exc}",
            data={
                "failure_stage": "F_output_finalization",
                "actual_export_path": actual_export_path,
                "requested_path": requested_path,
                "finalization_error": str(exc),
            },
        ) from exc
    return {
        "required": True,
        "attempted": True,
        "succeeded": True,
        "method": "os.replace_atomic_same_directory",
        "error": None,
    }


def _check_export_compatibility(
    document_kind: str,
    export_format: ExportFormat,
) -> dict[str, Any]:
    supported = document_kind in export_format.allowed_document_kinds
    evidence = {
        "document_kind": document_kind,
        "format": export_format.canonical_name,
        "supported_by_module_matrix": supported,
        "allowed_document_kinds": sorted(export_format.allowed_document_kinds),
    }
    if not supported:
        raise ExportOperationError(
            f"Format '{export_format.canonical_name}' is not enabled for "
            f"active document type '{document_kind}'.",
            data={
                "failure_stage": "B_document_format_compatibility",
                "compatibility": evidence,
            },
        )
    return evidence


def _wait_for_stable_nonempty_file(
    file_path: str,
    timeout_seconds: float,
) -> dict[str, Any]:
    deadline = time.monotonic() + timeout_seconds
    previous_signature: Optional[tuple[int, int]] = None
    stable_polls = 0
    polls = 0
    last_size = -1
    last_mtime_ns = -1
    while time.monotonic() < deadline:
        polls += 1
        try:
            stat = os.stat(file_path)
            last_size = int(stat.st_size)
            last_mtime_ns = int(stat.st_mtime_ns)
            signature = (last_size, last_mtime_ns)
        except OSError:
            signature = None
            last_size = -1
            last_mtime_ns = -1
        if signature is not None and last_size > 0:
            if signature == previous_signature:
                stable_polls += 1
            else:
                stable_polls = 1
            if stable_polls >= _STABLE_POLL_COUNT:
                return {
                    "materialised": True,
                    "stable": True,
                    "size_bytes": last_size,
                    "mtime_ns": last_mtime_ns,
                    "poll_count": polls,
                    "stable_poll_count": stable_polls,
                    "timeout_seconds": timeout_seconds,
                }
        else:
            stable_polls = 0
        previous_signature = signature
        time.sleep(_STABLE_POLL_INTERVAL_SECONDS)
    if last_size == 0:
        raise CATIAError(f"CATIA created an empty file: {file_path}")
    raise CATIAError(
        "CATIA did not materialise a stable non-empty file within "
        f"{timeout_seconds:.1f}s: {file_path}"
    )


def _read_prefix(file_path: str, length: int = 4096) -> bytes:
    with open(file_path, "rb") as stream:
        return stream.read(length)


def _verify_iges(file_path: str) -> bool:
    with open(file_path, "rb") as stream:
        sample = stream.read(80 * 20)
    if len(sample) < 80:
        return False
    for offset in range(0, len(sample) - 79, 80):
        record = sample[offset:offset + 80]
        if len(record) == 80 and record[72:73] in {b"S", b"G", b"D", b"P", b"T"}:
            return True
    # Text-mode line endings can make physical records 81/82 bytes.
    text = sample.decode("ascii", errors="ignore")
    return any(
        len(line) >= 73 and line[72] in "SGDPT"
        for line in text.splitlines()[:20]
    )


def _verify_stl(file_path: str, prefix: bytes, size_bytes: int) -> bool:
    stripped = prefix.lstrip()
    if stripped[:5].lower() == b"solid" and b"facet" in prefix.lower():
        return True
    if size_bytes < 84:
        return False
    try:
        with open(file_path, "rb") as stream:
            stream.seek(80)
            triangle_count = struct.unpack("<I", stream.read(4))[0]
        return size_bytes == 84 + 50 * triangle_count
    except (OSError, struct.error):
        return False


def _file_signature_check(
    file_path: str,
    canonical_format: str,
) -> dict[str, Any]:
    size = os.path.getsize(file_path)
    prefix = _read_prefix(file_path)
    method = ""
    verified: Optional[bool]
    details = ""
    if canonical_format == "step":
        method = "ISO-10303-21 header"
        verified = prefix.lstrip().startswith(b"ISO-10303-21")
    elif canonical_format == "pdf":
        method = "PDF magic"
        verified = prefix.startswith(b"%PDF-")
    elif canonical_format == "3dxml":
        method = "ZIP/XML header"
        stripped = prefix.lstrip(b"\xef\xbb\xbf \t\r\n")
        verified = prefix.startswith(b"PK\x03\x04") or stripped.startswith(
            (b"<?xml", b"<Model_3dxml", b"<model_3dxml")
        )
    elif canonical_format == "vrml":
        method = "VRML header"
        verified = prefix.lstrip().startswith(b"#VRML")
    elif canonical_format == "dxf":
        method = "DXF header"
        upper = prefix.upper()
        verified = (
            upper.startswith(b"AUTOCAD BINARY DXF")
            or b"SECTION" in upper
        )
    elif canonical_format == "dwg":
        method = "DWG AC10 signature"
        verified = prefix.startswith(b"AC10")
    elif canonical_format == "stl":
        method = "ASCII/Binary STL structure"
        verified = _verify_stl(file_path, prefix, size)
    elif canonical_format == "iges":
        method = "IGES section-letter record"
        verified = _verify_iges(file_path)
    elif canonical_format == "bmp":
        method = "BMP magic"
        verified = prefix.startswith(b"BM")
    elif canonical_format == "jpeg":
        method = "JPEG SOI"
        verified = prefix.startswith(b"\xff\xd8\xff")
    elif canonical_format == "tiff":
        method = "TIFF endian/magic"
        verified = prefix.startswith((b"II*\x00", b"MM\x00*"))
    elif canonical_format == "emf":
        method = "EMF header signature"
        verified = len(prefix) >= 44 and prefix[40:44] == b" EMF"
    elif canonical_format in {"cgr", "cgm"}:
        method = "no portable lightweight signature validator"
        verified = None
        details = "Non-empty stable file verified; semantic validation is external."
    else:
        method = "unknown"
        verified = None
        details = "No signature validator is registered."

    return {
        "available": verified is not None,
        "verified": verified,
        "method": method,
        "details": details,
        "prefix_hex": prefix[:16].hex(),
    }


def _validate_generated_file(
    file_path: str,
    canonical_format: str,
    timeout_seconds: float,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    stability = _wait_for_stable_nonempty_file(file_path, timeout_seconds)
    signature = _file_signature_check(file_path, canonical_format)
    if signature["available"] and not signature["verified"]:
        raise ExportOperationError(
            f"Generated file does not match the expected "
            f"{canonical_format.upper()} signature.",
            data={
                "failure_stage": "E_output_signature_validation",
                "file_stability": stability,
                "signature_check": signature,
            },
        )
    if not signature["available"]:
        warnings.append(
            f"No lightweight signature validator is available for "
            f"{canonical_format}; only stable non-empty output was verified."
        )
    file_info = {
        **stability,
        "sha256": _sha256_file(file_path),
        "signature_check": signature,
    }
    return file_info, warnings


def _evaluate(
    application: Any,
    script: str,
    function_name: str,
    parameters: list[Any],
) -> Any:
    try:
        return application.SystemService.Evaluate(
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


_VIEWPOINT_SET_SCRIPT = """
Public Function MCP_SetViewpointDirections(vp, sx, sy, sz, ux, uy, uz)
    Dim sight(2)
    Dim up(2)
    sight(0) = CDbl(sx)
    sight(1) = CDbl(sy)
    sight(2) = CDbl(sz)
    up(0) = CDbl(ux)
    up(1) = CDbl(uy)
    up(2) = CDbl(uz)
    vp.PutSightDirection sight
    vp.PutUpDirection up
    MCP_SetViewpointDirections = True
End Function
"""

_VIEWPOINT_GET_SCRIPT = """
Public Function MCP_GetViewpointDirections(vp)
    Dim sight(2)
    Dim up(2)
    vp.GetSightDirection sight
    vp.GetUpDirection up
    MCP_GetViewpointDirections = Array( _
        CDbl(sight(0)), CDbl(sight(1)), CDbl(sight(2)), _
        CDbl(up(0)), CDbl(up(1)), CDbl(up(2)) _
    )
End Function
"""


def _finite_vector(
    values: tuple[float, float, float],
    name: str,
) -> tuple[float, float, float]:
    result: list[float] = []
    for item in values:
        try:
            number = float(item)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{name} contains a non-numeric value.") from exc
        if not math.isfinite(number):
            raise ValueError(f"{name} contains a non-finite value.")
        result.append(number)
    magnitude = math.sqrt(sum(number * number for number in result))
    if magnitude <= 1.0e-12:
        raise ValueError(f"{name} cannot be the zero vector.")
    return result[0], result[1], result[2]


def _normalise_vector(
    values: tuple[float, float, float],
) -> tuple[float, float, float]:
    magnitude = math.sqrt(sum(value * value for value in values))
    return tuple(value / magnitude for value in values)  # type: ignore[return-value]


def _vector_error(
    expected: tuple[float, float, float],
    actual: tuple[float, float, float],
) -> float:
    expected_n = _normalise_vector(expected)
    actual_n = _normalise_vector(actual)
    return math.sqrt(
        sum((left - right) ** 2 for left, right in zip(expected_n, actual_n))
    )


def _set_viewpoint_directions(
    application: Any,
    viewpoint: Any,
    sight: tuple[float, float, float],
    up: tuple[float, float, float],
) -> dict[str, Any]:
    attempts: list[dict[str, Any]] = []
    strategy = None
    try:
        _evaluate(
            application,
            _VIEWPOINT_SET_SCRIPT,
            "MCP_SetViewpointDirections",
            [viewpoint, *sight, *up],
        )
        attempts.append(
            {
                "method": "SystemService.Evaluate",
                "succeeded": True,
                "error": None,
            }
        )
        strategy = "SystemService.Evaluate"
    except Exception as exc:
        attempts.append(
            {
                "method": "SystemService.Evaluate",
                "succeeded": False,
                "error": _format_com_error(exc),
            }
        )
        try:
            viewpoint.PutSightDirection(sight)
            viewpoint.PutUpDirection(up)
            attempts.append(
                {
                    "method": "direct_COM_tuple",
                    "succeeded": True,
                    "error": None,
                }
            )
            strategy = "direct_COM_tuple"
        except Exception as direct_exc:
            attempts.append(
                {
                    "method": "direct_COM_tuple",
                    "succeeded": False,
                    "error": _format_com_error(direct_exc),
                }
            )
            raise ExportOperationError(
                "CATIA rejected the requested sight/up directions.",
                data={
                    "failure_stage": "C_viewpoint_write",
                    "set_attempts": attempts,
                },
            ) from direct_exc

    read_attempts: list[dict[str, Any]] = []
    values: Optional[list[float]] = None
    try:
        raw = _evaluate(
            application,
            _VIEWPOINT_GET_SCRIPT,
            "MCP_GetViewpointDirections",
            [viewpoint],
        )
        values = [float(item) for item in list(raw)]
        if len(values) != 6 or not all(math.isfinite(item) for item in values):
            raise RuntimeError("CATIA returned an invalid six-value direction array.")
        read_attempts.append(
            {
                "method": "SystemService.Evaluate",
                "succeeded": True,
                "error": None,
            }
        )
    except Exception as exc:
        read_attempts.append(
            {
                "method": "SystemService.Evaluate",
                "succeeded": False,
                "error": _format_com_error(exc),
            }
        )
        try:
            read_sight = list(viewpoint.GetSightDirection())
            read_up = list(viewpoint.GetUpDirection())
            values = [float(item) for item in read_sight + read_up]
            if len(values) != 6 or not all(math.isfinite(item) for item in values):
                raise RuntimeError("Direct COM returned invalid direction arrays.")
            read_attempts.append(
                {
                    "method": "direct_COM_return",
                    "succeeded": True,
                    "error": None,
                }
            )
        except Exception as direct_exc:
            read_attempts.append(
                {
                    "method": "direct_COM_return",
                    "succeeded": False,
                    "error": _format_com_error(direct_exc),
                }
            )

    if values is None:
        raise ExportOperationError(
            "Viewpoint directions were written but could not be verified.",
            data={
                "failure_stage": "D_viewpoint_readback",
                "write_strategy": strategy,
                "set_attempts": attempts,
                "read_attempts": read_attempts,
            },
        )

    actual_sight = (values[0], values[1], values[2])
    actual_up = (values[3], values[4], values[5])
    sight_error = _vector_error(sight, actual_sight)
    up_error = _vector_error(up, actual_up)
    verified = sight_error <= _VECTOR_TOLERANCE and up_error <= _VECTOR_TOLERANCE
    if not verified:
        raise ExportOperationError(
            "Viewpoint readback did not match the requested directions.",
            data={
                "failure_stage": "D_viewpoint_readback",
                "write_strategy": strategy,
                "set_attempts": attempts,
                "read_attempts": read_attempts,
                "requested_sight": list(sight),
                "requested_up": list(up),
                "actual_sight": list(actual_sight),
                "actual_up": list(actual_up),
                "sight_normalised_error": sight_error,
                "up_normalised_error": up_error,
            },
        )
    return {
        "write_strategy": strategy,
        "set_attempts": attempts,
        "read_attempts": read_attempts,
        "requested_sight": list(sight),
        "requested_up": list(up),
        "actual_sight": list(actual_sight),
        "actual_up": list(actual_up),
        "sight_normalised_error": sight_error,
        "up_normalised_error": up_error,
        "readback_verified": True,
    }


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    """Register export and active-viewer tools."""
    conn = ctx.conn
    names: list[str] = []

    @mcp.tool()
    def catia_export(
        file_path: str,
        format_name: str = "",
        overwrite: bool = False,
        timeout_seconds: float = _DEFAULT_TIMEOUT_SECONDS,
    ) -> dict[str, Any]:
        """Export the active document with transactional overwrite protection.

        Supported combinations:

        * CATPart: STEP, IGES, STL, 3DXML, VRML, CGR
        * CATProduct: STEP, IGES, 3DXML, VRML, CGR
        * CATDrawing: PDF, DWG, DXF, 2D IGES

        The requested format must agree with the filename extension.  A long
        ``.iges`` request is exported to a unique temporary ``.igs`` sibling,
        validated and atomically renamed to the requested target.
        """
        normalised = ""
        actual_export_path = ""
        transaction: Optional[_TargetTransaction] = None
        warnings: list[str] = []
        operation_started = time.time_ns()
        export_path_plan: Optional[dict[str, Any]] = None
        staging_cleanup: Optional[dict[str, Any]] = None
        finalization: Optional[dict[str, Any]] = None
        try:
            normalised = _normalised_file_path(file_path)
            timeout = _positive_timeout(timeout_seconds)
            export_format = _resolve_export_format(normalised, format_name)
            actual_export_path, export_path_plan = _plan_export_path(
                normalised, export_format
            )

            conn.ensure_connected()
            application, application_resolution = _application(conn)
            document = _active_document(application)
            source = _document_info(document)
            compatibility = _check_export_compatibility(
                source["kind"], export_format
            )
            if source["full_name"] and _same_path(source["full_name"], normalised):
                raise ExportOperationError(
                    "Export target must not overwrite the active CATIA source document.",
                    data={
                        "failure_stage": "B_source_target_collision",
                        "source_document": source,
                        "target_path": normalised,
                        "export_path_plan": export_path_plan,
                    },
                )

            if not _same_path(actual_export_path, normalised) and os.path.exists(
                actual_export_path
            ):
                raise ExportOperationError(
                    "The generated IGES staging path unexpectedly already exists.",
                    data={
                        "failure_stage": "B_staging_path_collision",
                        "export_path_plan": export_path_plan,
                    },
                )

            transaction = _TargetTransaction(normalised, overwrite)
            transaction.prepare()
            try:
                document.ExportData(
                    actual_export_path, export_format.catia_token
                )
            except Exception as exc:
                converter_context = {
                    "document_kind": source["kind"],
                    "format": export_format.canonical_name,
                    "catia_format_token": export_format.catia_token,
                    "requested_path": normalised,
                    "actual_export_path": actual_export_path,
                }
                if source["kind"] == _DRAWING and export_format.canonical_name == "iges":
                    converter_context.update({
                        "converter_classification": (
                            "CATDrawing_2D_IGES_converter_failure"
                        ),
                        "message": (
                            "The module permits CATDrawing 2D IGES, but the "
                            "active CATIA configuration rejected ExportData."
                        ),
                        "suggested_action": (
                            "Verify the installed 2D IGES translator/license and "
                            "retry with a minimal drawing."
                        ),
                    })
                raise ExportOperationError(
                    f"Document.ExportData failed: {_format_com_error(exc)}",
                    data={
                        "failure_stage": "C_ExportData_call",
                        "source_document": source,
                        "compatibility": compatibility,
                        "export_path_plan": export_path_plan,
                        "converter_context": converter_context,
                    },
                ) from exc

            try:
                file_info, validation_warnings = _validate_generated_file(
                    actual_export_path,
                    export_format.canonical_name,
                    timeout,
                )
            except ExportOperationError:
                raise
            except Exception as exc:
                raise ExportOperationError(
                    f"Generated output validation failed: {_format_com_error(exc)}",
                    data={
                        "failure_stage": "D_output_materialization_or_stability",
                        "source_document": source,
                        "compatibility": compatibility,
                        "export_path_plan": export_path_plan,
                        "requested_path": normalised,
                        "actual_export_path": actual_export_path,
                        "actual_path_exists": os.path.exists(actual_export_path),
                        "actual_path_size_bytes": (
                            os.path.getsize(actual_export_path)
                            if os.path.isfile(actual_export_path) else None
                        ),
                    },
                ) from exc
            warnings.extend(validation_warnings)
            validated_staging_sha256 = file_info["sha256"]
            finalization = _finalize_staged_export(
                actual_export_path, normalised
            )

            if not os.path.isfile(normalised):
                raise ExportOperationError(
                    "The validated export was not present at the requested final path.",
                    data={
                        "failure_stage": "F_output_finalization",
                        "export_path_plan": export_path_plan,
                        "finalization": finalization,
                    },
                )
            final_sha256 = _sha256_file(normalised)
            if final_sha256 != validated_staging_sha256:
                raise ExportOperationError(
                    "The final output SHA256 differs from the validated export.",
                    data={
                        "failure_stage": "F_output_finalization",
                        "validated_sha256": validated_staging_sha256,
                        "final_sha256": final_sha256,
                        "export_path_plan": export_path_plan,
                        "finalization": finalization,
                    },
                )
            file_info.update({
                "validated_path": actual_export_path,
                "final_path": normalised,
                "final_exists": True,
                "final_size_bytes": os.path.getsize(normalised),
                "final_sha256": final_sha256,
                "staging_path_exists_after_finalization": (
                    False if _same_path(actual_export_path, normalised)
                    else os.path.exists(actual_export_path)
                ),
            })

            source_after = _document_info(document)
            transaction.commit()
            warnings.extend(transaction.cleanup_warnings)
            return _success(
                {
                    "operation": "catia_export",
                    "exported": True,
                    "path": normalised,
                    "requested_path": normalised,
                    "actual_export_path": actual_export_path,
                    "format": export_format.canonical_name,
                    "catia_format_token": export_format.catia_token,
                    "export_path_plan": export_path_plan,
                    "finalization": finalization,
                    "source_document_before": source,
                    "source_document_after": source_after,
                    "source_document_modified": (
                        source.get("saved") is True
                        and source_after.get("saved") is False
                    ),
                    "compatibility": compatibility,
                    "application_resolution": application_resolution,
                    "output_file": file_info,
                    "target_transaction": transaction.as_dict(),
                    "operation_started_ns": operation_started,
                    "model_modified": False,
                    "document_save_required": False,
                },
                warnings,
            )
        except Exception as exc:
            if (
                actual_export_path
                and normalised
                and not _same_path(actual_export_path, normalised)
            ):
                staging_cleanup = _cleanup_staging_path(
                    actual_export_path, normalised
                )
                if staging_cleanup.get("error"):
                    warnings.append(
                        "Staging export cleanup failed: "
                        f"{staging_cleanup['error']}"
                    )
            if transaction is not None and transaction.prepared and not transaction.committed:
                transaction.rollback()
            error_data: dict[str, Any] = {}
            if isinstance(exc, ExportOperationError) and isinstance(exc.data, dict):
                error_data.update(exc.data)
            if normalised:
                error_data["target_path"] = normalised
                error_data["requested_path"] = normalised
            if actual_export_path:
                error_data["actual_export_path"] = actual_export_path
            if export_path_plan is not None:
                error_data["export_path_plan"] = export_path_plan
            if finalization is not None:
                error_data["finalization"] = finalization
            if staging_cleanup is not None:
                error_data["staging_cleanup"] = staging_cleanup
            if transaction is not None:
                error_data["target_transaction"] = transaction.as_dict()
            return _error(
                "catia_export",
                exc,
                data=error_data or None,
                warnings=warnings,
            )

    names.append("catia_export")

    @mcp.tool()
    def catia_screenshot(
        file_path: str,
        overwrite: bool = False,
        grayscale: bool = False,
        timeout_seconds: float = 15.0,
    ) -> dict[str, Any]:
        """Capture the active CATIA viewer with transactional file validation.

        Automation supports CGM, EMF, TIFF, TIFF greyscale, BMP and JPEG.
        PNG is rejected because the Viewer Automation enum has no PNG value.
        """
        normalised = ""
        transaction: Optional[_TargetTransaction] = None
        warnings: list[str] = []
        try:
            if not isinstance(grayscale, bool):
                raise ValueError("grayscale must be a boolean.")
            timeout = _positive_timeout(timeout_seconds)
            normalised = _normalised_file_path(file_path)
            extension = Path(normalised).suffix.lower()
            if extension not in CAPTURE_FORMATS:
                supported = ", ".join(sorted(CAPTURE_FORMATS))
                raise ValueError(
                    f"Unsupported capture extension '{extension or '<none>'}'. "
                    f"CATIA V5 Automation supports: {supported}; PNG is not supported."
                )
            if grayscale and extension not in {".tif", ".tiff"}:
                raise ValueError("grayscale=true is supported only for TIFF captures.")

            capture_format = 3 if grayscale else CAPTURE_FORMATS[extension]
            canonical_format = {
                ".cgm": "cgm",
                ".emf": "emf",
                ".tif": "tiff",
                ".tiff": "tiff",
                ".bmp": "bmp",
                ".jpg": "jpeg",
                ".jpeg": "jpeg",
            }[extension]

            conn.ensure_connected()
            application, application_resolution = _application(conn)
            viewer, viewer_info = _active_viewer(application)
            transaction = _TargetTransaction(normalised, overwrite)
            transaction.prepare()
            update_succeeded = True
            update_error = None
            try:
                viewer.Update()
            except Exception as exc:
                update_succeeded = False
                update_error = _format_com_error(exc)
                warnings.append(
                    f"Viewer.Update before capture failed: {update_error}"
                )
            try:
                viewer.CaptureToFile(capture_format, normalised)
            except Exception as exc:
                raise ExportOperationError(
                    f"Viewer.CaptureToFile failed: {_format_com_error(exc)}",
                    data={
                        "failure_stage": "C_CaptureToFile_call",
                        "viewer": viewer_info,
                        "capture_format_code": capture_format,
                    },
                ) from exc

            file_info, validation_warnings = _validate_generated_file(
                normalised, canonical_format, timeout
            )
            warnings.extend(validation_warnings)
            transaction.commit()
            warnings.extend(transaction.cleanup_warnings)
            return _success(
                {
                    "operation": "catia_screenshot",
                    "saved": True,
                    "path": normalised,
                    "capture_format_code": capture_format,
                    "capture_format": canonical_format,
                    "grayscale": grayscale,
                    "application_resolution": application_resolution,
                    "viewer": viewer_info,
                    "viewer_update": {
                        "succeeded": update_succeeded,
                        "error": update_error,
                    },
                    "output_file": file_info,
                    "target_transaction": transaction.as_dict(),
                    "model_modified": False,
                    "document_save_required": False,
                },
                warnings,
            )
        except Exception as exc:
            if transaction is not None and transaction.prepared and not transaction.committed:
                transaction.rollback()
            error_data: dict[str, Any] = {}
            if isinstance(exc, ExportOperationError) and isinstance(exc.data, dict):
                error_data.update(exc.data)
            if normalised:
                error_data["target_path"] = normalised
            if transaction is not None:
                error_data["target_transaction"] = transaction.as_dict()
            return _error(
                "catia_screenshot",
                exc,
                data=error_data or None,
                warnings=warnings,
            )

    names.append("catia_screenshot")

    @mcp.tool()
    def catia_fit_all() -> dict[str, Any]:
        """Fit all visible geometry in the active CATIA viewer."""
        try:
            conn.ensure_connected()
            application, application_resolution = _application(conn)
            viewer, viewer_info = _active_viewer(application)
            viewer.Reframe()
            update_succeeded = True
            update_error = None
            try:
                viewer.Update()
            except Exception as exc:
                update_succeeded = False
                update_error = _format_com_error(exc)
            return _success(
                {
                    "operation": "catia_fit_all",
                    "fit_all": True,
                    "reframe_called": True,
                    "viewer_update": {
                        "succeeded": update_succeeded,
                        "error": update_error,
                    },
                    "viewer": viewer_info,
                    "application_resolution": application_resolution,
                    "model_modified": False,
                    "document_save_required": False,
                },
                [] if update_succeeded else [
                    f"Viewer.Update after Reframe failed: {update_error}"
                ],
            )
        except Exception as exc:
            return _error("catia_fit_all", exc)

    names.append("catia_fit_all")

    @mcp.tool()
    def catia_set_view(view: str = "isometric") -> dict[str, Any]:
        """Set and verify a standard 3D viewpoint.

        Supported values: front, back, top, bottom, left, right and isometric.
        """
        try:
            key = _nonempty_text(view, "view").lower()
            views: dict[str, dict[str, tuple[float, float, float]]] = {
                "front": {"sight": (0.0, 0.0, -1.0), "up": (0.0, 1.0, 0.0)},
                "back": {"sight": (0.0, 0.0, 1.0), "up": (0.0, 1.0, 0.0)},
                "top": {"sight": (0.0, -1.0, 0.0), "up": (0.0, 0.0, -1.0)},
                "bottom": {"sight": (0.0, 1.0, 0.0), "up": (0.0, 0.0, 1.0)},
                "left": {"sight": (1.0, 0.0, 0.0), "up": (0.0, 1.0, 0.0)},
                "right": {"sight": (-1.0, 0.0, 0.0), "up": (0.0, 1.0, 0.0)},
                "isometric": {
                    "sight": (-1.0, -1.0, -1.0),
                    "up": (-1.0, -1.0, 2.0),
                },
            }
            if key not in views:
                raise ValueError(
                    f"Unsupported view '{view}'. Supported: {', '.join(views)}."
                )
            sight = _finite_vector(views[key]["sight"], "sight")
            up = _finite_vector(views[key]["up"], "up")
            sight_n = _normalise_vector(sight)
            up_n = _normalise_vector(up)
            if abs(sum(a * b for a, b in zip(sight_n, up_n))) > 1.0e-6:
                raise RuntimeError("Configured sight and up directions are not orthogonal.")

            conn.ensure_connected()
            application, application_resolution = _application(conn)
            viewer, viewer_info = _active_viewer(application, require_3d=True)
            viewpoint = viewer.Viewpoint3D
            viewpoint_evidence = _set_viewpoint_directions(
                application, viewpoint, sight, up
            )
            viewer.Reframe()
            update_succeeded = True
            update_error = None
            try:
                viewer.Update()
            except Exception as exc:
                update_succeeded = False
                update_error = _format_com_error(exc)
            return _success(
                {
                    "operation": "catia_set_view",
                    "view": key,
                    "viewpoint": viewpoint_evidence,
                    "reframed": True,
                    "viewer_update": {
                        "succeeded": update_succeeded,
                        "error": update_error,
                    },
                    "viewer": viewer_info,
                    "application_resolution": application_resolution,
                    "model_modified": False,
                    "document_save_required": False,
                },
                [] if update_succeeded else [
                    f"Viewer.Update after viewpoint change failed: {update_error}"
                ],
            )
        except Exception as exc:
            return _error("catia_set_view", exc)

    names.append("catia_set_view")
    return names
