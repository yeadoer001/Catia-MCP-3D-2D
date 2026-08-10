"""CATIA V5 document and connection MCP tools.

The functions in this module are registered lazily through ``register_tools``
so importing the module never starts CATIA.  Every mutating operation performs
both input validation and a small post-condition check before it reports
success.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Optional

from catia_mcp.connection import CATIAError, normalize_path


IMPLEMENTATION_VERSION = "document-fixed-2026-07-31-v1"


def _success(data: Any, warnings: Optional[list[str]] = None) -> dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": True,
        "status": "success_with_warnings" if warning_list else "success",
        "data": data,
        "warnings": warning_list,
    }


def _error(operation: str, exc: BaseException) -> dict[str, Any]:
    return {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": False,
        "status": "error",
        "operation": operation,
        "error": str(exc),
        "error_type": type(exc).__name__,
        "warnings": [],
    }


def _nonempty_text(value: Any, parameter_name: str) -> str:
    text = str(value).strip()
    if not text:
        raise ValueError(f"{parameter_name} cannot be empty.")
    if any(ord(character) < 32 for character in text):
        raise ValueError(f"{parameter_name} contains a control character.")
    if len(text) > 255:
        raise ValueError(f"{parameter_name} cannot exceed 255 characters.")
    return text


def _normalised_file_path(value: Any, parameter_name: str) -> str:
    raw_path = _nonempty_text(value, parameter_name)
    normalised = str(normalize_path(raw_path)).strip()
    if not normalised:
        raise ValueError(f"{parameter_name} could not be normalised.")
    return normalised


def _application(conn: Any) -> Any:
    """Return the CATIA Application exposed by the connection object."""
    for attribute_name in ("app", "application", "catia"):
        try:
            application = getattr(conn, attribute_name)
        except Exception:
            continue
        if application is not None:
            return application
    raise CATIAError("The CATIA connection does not expose its Application object.")


def _safe_attr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _document_saved(document: Any) -> Optional[bool]:
    value = _safe_attr(document, "Saved", None)
    return bool(value) if value is not None else None


def _document_full_name(document: Any) -> str:
    value = _safe_attr(document, "FullName", "")
    return str(value or "")


def _same_path(first: str, second: str) -> bool:
    return os.path.normcase(os.path.abspath(first)) == os.path.normcase(
        os.path.abspath(second)
    )


def _set_part_number(document: Any, requested_name: str) -> str:
    """Set PartNumber on the document's reference Product.

    A CATPart's ``Part`` object does not own PartNumber.  CATIA creates a
    reference ``Product`` inside every PartDocument specifically for this
    property, which is why both CATPart and CATProduct use ``document.Product``
    here.
    """
    try:
        root_product = document.Product
    except Exception as exc:
        raise CATIAError(
            "The newly created document does not expose a root Product."
        ) from exc

    try:
        root_product.PartNumber = requested_name
        actual_name = str(root_product.PartNumber)
    except Exception as exc:
        raise CATIAError(f"CATIA rejected PartNumber '{requested_name}': {exc}") from exc

    if actual_name != requested_name:
        raise CATIAError(
            "CATIA did not retain the requested PartNumber "
            f"(requested={requested_name!r}, actual={actual_name!r})."
        )
    return actual_name


def _close_created_document(document: Any) -> None:
    try:
        document.Close()
    except Exception:
        pass


def _save_as(
    conn: Any,
    document: Any,
    file_path: str,
    *,
    overwrite: bool,
) -> None:
    target_exists = os.path.exists(file_path)
    current_path = _document_full_name(document)
    same_as_current = bool(current_path) and _same_path(current_path, file_path)
    if target_exists and not same_as_current and not overwrite:
        raise FileExistsError(
            f"Target already exists: {file_path}. Set overwrite=true to replace it."
        )

    Path(file_path).parent.mkdir(parents=True, exist_ok=True)

    application = _application(conn)
    previous_alert_state = _safe_attr(application, "DisplayFileAlerts", None)
    alerts_changed = bool(overwrite and previous_alert_state is not None)
    try:
        if alerts_changed:
            application.DisplayFileAlerts = False
        document.SaveAs(file_path)
    finally:
        if alerts_changed:
            application.DisplayFileAlerts = previous_alert_state

    if not os.path.isfile(file_path):
        raise CATIAError(
            "CATIA SaveAs returned without creating the requested file: "
            f"{file_path}"
        )

    saved_path = _document_full_name(document)
    if saved_path and not _same_path(saved_path, file_path):
        raise CATIAError(
            "CATIA saved a different path than requested "
            f"(requested={file_path!r}, actual={saved_path!r})."
        )


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    """Register connection and document tools on a FastMCP-compatible server."""
    conn = ctx.conn
    names: list[str] = []

    @mcp.tool()
    def catia_start(visible: bool = True) -> dict[str, Any]:
        """Start CATIA V5 or connect to the running instance."""
        try:
            if not isinstance(visible, bool):
                raise ValueError("visible must be a boolean.")
            conn.connect(visible=visible)
            return _success(conn.get_connection_info())
        except Exception as exc:
            return _error("catia_start", exc)

    names.append("catia_start")

    @mcp.tool()
    def catia_status() -> dict[str, Any]:
        """Return connection information without starting CATIA."""
        try:
            return _success(conn.get_connection_info())
        except Exception as exc:
            return _error("catia_status", exc)

    names.append("catia_status")

    @mcp.tool()
    def catia_disconnect() -> dict[str, Any]:
        """Release Python's COM references; CATIA itself remains open."""
        try:
            conn.disconnect()
            return _success({"disconnected": True, "catia_closed": False})
        except Exception as exc:
            return _error("catia_disconnect", exc)

    names.append("catia_disconnect")

    @mcp.tool()
    def catia_new_part(name: str = "MCP_Part") -> dict[str, Any]:
        """Create a CATPart and set its reference Product.PartNumber."""
        document = None
        try:
            requested_name = _nonempty_text(name, "name")
            conn.ensure_connected()
            document = conn.documents.Add("Part")
            actual_name = _set_part_number(document, requested_name)
            _ = document.Part
            conn.refresh_display()
            data = conn.describe_document(document)
            return _success({"document": data, "part_number": actual_name})
        except Exception as exc:
            if document is not None:
                _close_created_document(document)
            return _error("catia_new_part", exc)

    names.append("catia_new_part")

    @mcp.tool()
    def catia_new_product(name: str = "MCP_Product") -> dict[str, Any]:
        """Create a CATProduct and set its root Product.PartNumber."""
        document = None
        try:
            requested_name = _nonempty_text(name, "name")
            conn.ensure_connected()
            document = conn.documents.Add("Product")
            actual_name = _set_part_number(document, requested_name)
            conn.refresh_display()
            data = conn.describe_document(document)
            return _success({"document": data, "part_number": actual_name})
        except Exception as exc:
            if document is not None:
                _close_created_document(document)
            return _error("catia_new_product", exc)

    names.append("catia_new_product")

    @mcp.tool()
    def catia_new_drawing() -> dict[str, Any]:
        """Create a CATDrawing containing at least one sheet."""
        document = None
        try:
            conn.ensure_connected()
            document = conn.documents.Add("Drawing")
            sheet_count = int(document.Sheets.Count)
            if sheet_count < 1:
                raise CATIAError("CATIA created a drawing with no sheet.")
            conn.refresh_display()
            return _success(
                {
                    "document": conn.describe_document(document),
                    "sheet_count": sheet_count,
                }
            )
        except Exception as exc:
            if document is not None:
                _close_created_document(document)
            return _error("catia_new_drawing", exc)

    names.append("catia_new_drawing")

    @mcp.tool()
    def catia_open_document(file_path: str) -> dict[str, Any]:
        """Open an existing local CATIA document."""
        try:
            normalised = _normalised_file_path(file_path, "file_path")
            if not os.path.isfile(normalised):
                raise FileNotFoundError(f"Document does not exist: {normalised}")
            conn.ensure_connected()
            document = conn.documents.Open(normalised)
            return _success(
                {
                    "document": conn.describe_document(document),
                    "path": normalised,
                }
            )
        except Exception as exc:
            return _error("catia_open_document", exc)

    names.append("catia_open_document")

    @mcp.tool()
    def catia_save_document(
        file_path: str = "",
        overwrite: bool = False,
    ) -> dict[str, Any]:
        """Save the active document, or Save As to ``file_path``.

        Existing files are not replaced unless ``overwrite`` is true.
        """
        try:
            if not isinstance(overwrite, bool):
                raise ValueError("overwrite must be a boolean.")
            conn.ensure_connected()
            document = conn.active_document

            if str(file_path).strip():
                normalised = _normalised_file_path(file_path, "file_path")
                _save_as(
                    conn,
                    document,
                    normalised,
                    overwrite=overwrite,
                )
                saved_path = normalised
                save_method = "SaveAs"
            else:
                if not _document_full_name(document):
                    raise ValueError(
                        "The active document has never been saved; file_path is required."
                    )
                document.Save()
                if _document_saved(document) is False:
                    raise CATIAError("CATIA Save returned but the document is still modified.")
                saved_path = _document_full_name(document)
                save_method = "Save"

            return _success(
                {
                    "document": conn.describe_document(document),
                    "path": saved_path,
                    "method": save_method,
                    "saved": _document_saved(document),
                }
            )
        except Exception as exc:
            return _error("catia_save_document", exc)

    names.append("catia_save_document")

    @mcp.tool()
    def catia_close_document(save: bool = False) -> dict[str, Any]:
        """Close the active document, optionally saving it first.

        When ``save`` is false, CATIA file alerts are temporarily disabled so
        an unsaved-document dialog cannot deadlock the MCP request.
        """
        try:
            if not isinstance(save, bool):
                raise ValueError("save must be a boolean.")
            conn.ensure_connected()
            document = conn.active_document
            name = str(_safe_attr(document, "Name", ""))
            was_saved = _document_saved(document)
            if save:
                if not _document_full_name(document):
                    raise ValueError(
                        "The active document has never been saved; use "
                        "catia_save_document(file_path=...) before closing it."
                    )
                document.Save()
                if _document_saved(document) is False:
                    raise CATIAError("CATIA Save returned but the document is still modified.")

            application = _application(conn)
            previous_alert_state = _safe_attr(application, "DisplayFileAlerts", None)
            alerts_changed = bool(not save and previous_alert_state is not None)
            try:
                if alerts_changed:
                    application.DisplayFileAlerts = False
                document.Close()
            finally:
                if alerts_changed:
                    application.DisplayFileAlerts = previous_alert_state

            return _success(
                {
                    "closed": True,
                    "document_name": name,
                    "saved_before_close": save,
                    "was_saved_on_entry": was_saved,
                }
            )
        except Exception as exc:
            return _error("catia_close_document", exc)

    names.append("catia_close_document")

    @mcp.tool()
    def catia_list_documents() -> dict[str, Any]:
        """Describe every open CATIA document in collection order."""
        try:
            conn.ensure_connected()
            documents = conn.documents
            result: list[dict[str, Any]] = []
            warnings: list[str] = []

            for index in range(1, int(documents.Count) + 1):
                try:
                    result.append(conn.describe_document(documents.Item(index)))
                except Exception as exc:
                    warnings.append(f"Could not describe document {index}: {exc}")
                    result.append(
                        {
                            "index": index,
                            "description_available": False,
                            "error": str(exc),
                        }
                    )

            return _success(result, warnings)
        except Exception as exc:
            return _error("catia_list_documents", exc)

    names.append("catia_list_documents")

    @mcp.tool()
    def catia_get_active_document_info() -> dict[str, Any]:
        """Describe the active CATIA document."""
        try:
            conn.ensure_connected()
            return _success(conn.describe_document(conn.active_document))
        except Exception as exc:
            return _error("catia_get_active_document_info", exc)

    names.append("catia_get_active_document_info")

    return names
