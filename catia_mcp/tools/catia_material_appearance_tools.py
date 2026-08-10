"""
catia_material_appearance_tools.py
Version: catia-material-appearance-fixed-2026-07-30-v4

CATIA V5 MCP tools for graphical appearance and material assignment.

Corrections:
- Graphic properties are changed through ActiveDocument.Selection.VisProperties.
- Transparency is converted to CATIA opacity and applied with SetRealOpacity.
- Color and opacity are read back and verified.
- The caller's existing Selection is restored when possible.
- Materials are read from CATMaterial documents, not a hard-coded CATfct path.
- Material catalog discovery uses CATIA environment paths or an explicit path.
- Materials are applied through CATMatManagerVBExt ApplyMaterialOnPart,
  ApplyMaterialOnBody, or ApplyMaterialOnProduct.
- Material application is read back and verified.
- target_scope='part' is a compatibility alias for the active MainBody,
  because the CATPart root Product may not expose stable real appearance.
- Selection restoration is completed before the result is constructed.
- Catalog lifecycle is explicit and depends on material link mode.
- link_mode=0 copies material data and closes a catalog opened by the call.
- link_mode=1 keeps one CATMaterial document open as a session-level cache
  for linked material dependencies; this is success_with_warnings, not a
  cleanup failure.
- Failed lookups and rolled-back assignments close a catalog opened by the
  call because no linked dependency remains.
- Before required close operations, material COM references are released and
  the original model document is reactivated.
- Required close operations use direct, fresh-wrapper and
  SystemService.Evaluate strategies, then verify the Documents collection.
- Material readback is repeated after lifecycle finalization.
"""

from __future__ import annotations

import gc
import math
import os
from pathlib import Path
from typing import Any, Optional

from catia_mcp.connection import CATIAError


IMPLEMENTATION_VERSION = (
    "catia-material-appearance-fixed-2026-07-30-v4"
)

_CATVB_SCRIPT_LANGUAGE = 1
_MATERIAL_MANAGER_EXTENSION = "CATMatManagerVBExt"
_ALLOWED_APPLY_TARGETS = {"part", "main_body", "product"}
_ALLOWED_COLOR_TARGETS = {"part", "main_body", "product", "named"}

_CATALOG_POLICY_CLOSE_AFTER_COPY = "close_after_copy"
_CATALOG_POLICY_CLOSE_AFTER_FAILED_LOOKUP = "close_after_failed_lookup"
_CATALOG_POLICY_CLOSE_AFTER_ROLLBACK = "close_after_rollback"
_CATALOG_POLICY_KEEP_OPEN_FOR_LINKED_MATERIAL = (
    "keep_open_for_linked_material"
)
_CATALOG_POLICY_KEEP_OPEN_FOR_UNRESOLVED_LINK = (
    "keep_open_for_unresolved_link_state"
)

# Normalized CATMaterial paths retained by this module for linked materials
# during the current CATIA/MCP Python process. CATIA Documents remains the
# authoritative cache; this set only distinguishes a tool-created session
# cache from a document that the caller had already opened.
_SESSION_LINKED_CATALOG_PATHS: set[str] = set()


# ---------------------------------------------------------------------------
# Standard results
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
# Connection and object helpers
# ---------------------------------------------------------------------------

def _get_connection(ctx: Any) -> Any:
    return getattr(ctx, "conn", ctx)


def _ensure_connected(conn: Any) -> Any:
    method = getattr(conn, "ensure_connected", None)
    if callable(method):
        result = method()
        if result is not None:
            return result

    app = getattr(conn, "app", None)
    if app is None:
        app = getattr(conn, "_app", None)
    if app is not None:
        return app

    raise CATIAError("Cannot access the CATIA Application object.")


def _get_application(conn: Any) -> Any:
    app = getattr(conn, "app", None)
    if app is None:
        app = getattr(conn, "_app", None)
    if app is not None:
        return app
    return _ensure_connected(conn)


def _active_document(conn: Any) -> Any:
    getter = getattr(conn, "get_active_document", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            pass

    application = _get_application(conn)
    try:
        return application.ActiveDocument
    except Exception as exc:
        raise CATIAError(
            "Cannot access the active CATIA document."
        ) from exc


def _active_part(conn: Any, document: Any) -> Any:
    getter = getattr(conn, "get_active_part", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            pass

    try:
        return document.Part
    except Exception as exc:
        raise CATIAError(
            "The active document does not expose a CATPart Part object."
        ) from exc


def _active_product(conn: Any, document: Any) -> Any:
    getter = getattr(conn, "get_active_product", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            pass

    try:
        return document.Product
    except Exception as exc:
        raise CATIAError(
            "The active document does not expose a Product object."
        ) from exc


def _main_body(conn: Any, part: Any) -> Any:
    getter = getattr(conn, "get_active_part_body", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            pass

    try:
        return part.MainBody
    except Exception:
        pass

    try:
        return part.PartBody
    except Exception as exc:
        raise CATIAError(
            "Cannot access the active CATPart main body."
        ) from exc


def _object_name(value: Any, fallback: str = "") -> str:
    try:
        result = str(value.Name).strip()
        return result or fallback
    except Exception:
        return fallback


def _describe_com_object(value: Any) -> dict[str, Any]:
    return {
        "python_type": type(value).__name__,
        "python_module": type(value).__module__,
        "name": _object_name(value),
        "has_oleobj": bool(hasattr(value, "_oleobj_")),
    }


def _normalise_nonempty(value: Any, parameter_name: str) -> str:
    text = str(value).strip()
    if not text:
        raise CATIAError(f"{parameter_name} cannot be empty.")
    return text


def _finite_number(value: Any, parameter_name: str) -> float:
    if isinstance(value, bool):
        raise CATIAError(f"{parameter_name} must be a finite number.")

    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(
            f"{parameter_name} must be a finite number."
        ) from exc

    if not math.isfinite(number):
        raise CATIAError(f"{parameter_name} must be finite.")

    return number


def _integer_in_range(
    value: Any,
    parameter_name: str,
    minimum: int,
    maximum: int,
) -> int:
    if isinstance(value, bool):
        raise CATIAError(
            f"{parameter_name} must be an integer from "
            f"{minimum} to {maximum}."
        )

    try:
        integer = int(value)
        numeric = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(
            f"{parameter_name} must be an integer from "
            f"{minimum} to {maximum}."
        ) from exc

    if numeric != float(integer) or not minimum <= integer <= maximum:
        raise CATIAError(
            f"{parameter_name} must be an integer from "
            f"{minimum} to {maximum}."
        )

    return integer


def _positive_integer(value: Any, parameter_name: str) -> int:
    return _integer_in_range(
        value,
        parameter_name,
        1,
        2_147_483_647,
    )


def _normalise_choice(
    value: Any,
    parameter_name: str,
    allowed: set[str],
) -> str:
    text = str(value).strip().lower()
    if text not in allowed:
        choices = ", ".join(sorted(allowed))
        raise CATIAError(
            f"{parameter_name} must be one of: {choices}."
        )
    return text


def _refresh_display(conn: Any) -> list[str]:
    warnings: list[str] = []

    method = getattr(conn, "refresh_display", None)
    if callable(method):
        try:
            method()
            return warnings
        except Exception as exc:
            warnings.append(f"Display refresh failed: {exc}")

    try:
        application = _get_application(conn)
        viewer = application.ActiveWindow.ActiveViewer
        viewer.Update()
    except Exception as exc:
        warnings.append(f"Active viewer refresh failed: {exc}")

    return warnings


# ---------------------------------------------------------------------------
# Selection handling
# ---------------------------------------------------------------------------

def _snapshot_selection(
    document: Any,
) -> tuple[list[Any], list[str]]:
    warnings: list[str] = []
    values: list[Any] = []
    selection = document.Selection

    try:
        count = int(selection.Count)
    except Exception as exc:
        return [], [f"Existing Selection count could not be read: {exc}"]

    for index in range(1, count + 1):
        try:
            item = selection.Item(index)
            values.append(item.Value)
        except Exception as exc:
            warnings.append(
                f"Selection item {index} could not be preserved: {exc}"
            )

    return values, warnings


def _restore_selection(
    document: Any,
    previous_values: list[Any],
) -> tuple[bool, list[str]]:
    warnings: list[str] = []
    selection = document.Selection

    try:
        selection.Clear()
    except Exception as exc:
        return False, [f"Selection could not be cleared: {exc}"]

    restored = 0
    for index, value in enumerate(previous_values, start=1):
        try:
            selection.Add(value)
            restored += 1
        except Exception as exc:
            warnings.append(
                f"Previous Selection item {index} could not be restored: "
                f"{exc}"
            )

    return restored == len(previous_values), warnings


def _selection_count(document: Any) -> Optional[int]:
    try:
        return int(document.Selection.Count)
    except Exception:
        return None


def _search_named_object(
    document: Any,
    name: str,
    occurrence: int,
    require_unique: bool,
) -> tuple[Any, dict[str, Any]]:
    target_name = _normalise_nonempty(name, "target_object_name")
    index = _positive_integer(occurrence, "occurrence")
    selection = document.Selection

    selection.Clear()
    try:
        selection.Search(f"Name={target_name},all")
        match_count = int(selection.Count)

        if match_count == 0:
            raise CATIAError(
                f"Target object not found: {target_name}"
            )

        if require_unique and match_count != 1:
            raise CATIAError(
                f"Target name '{target_name}' is ambiguous: "
                f"{match_count} objects matched. Rename the objects "
                "uniquely or set require_unique=false and provide "
                "an explicit occurrence."
            )

        if index > match_count:
            raise CATIAError(
                f"Occurrence {index} was requested for "
                f"'{target_name}', but only {match_count} objects matched."
            )

        value = selection.Item(index).Value
        return value, {
            "lookup_mode": "name",
            "requested_name": target_name,
            "resolved_name": _object_name(value, target_name),
            "match_count": match_count,
            "occurrence": index,
            "require_unique": bool(require_unique),
            "object": _describe_com_object(value),
        }
    finally:
        selection.Clear()


def _resolve_color_target(
    conn: Any,
    document: Any,
    target_scope: str,
    target_object_name: str,
    occurrence: int,
    require_unique: bool,
) -> tuple[Any, dict[str, Any], list[str]]:
    scope = _normalise_choice(
        target_scope,
        "target_scope",
        _ALLOWED_COLOR_TARGETS,
    )
    warnings: list[str] = []

    if scope == "named":
        value, details = _search_named_object(
            document,
            target_object_name,
            occurrence,
            require_unique,
        )
        details["target_scope"] = scope
        details["scope_alias_applied"] = False
        return value, details, warnings

    if target_object_name.strip():
        raise CATIAError(
            "target_object_name is only valid when "
            "target_scope='named'."
        )

    if scope == "part":
        # A CATPart document also exposes a root Product, but that Product can
        # inherit its displayed appearance from the Body and can therefore
        # reject stable real-color readback. For visual appearance, "part"
        # is treated as a compatibility alias for the active MainBody.
        part = _active_part(conn, document)
        value = _main_body(conn, part)
        resolved_scope = "active_part_main_body"
        warnings.append(
            "target_scope='part' is applied to the active MainBody "
            "visual representation. Use target_scope='named' for other "
            "Bodies or features in a multi-body CATPart."
        )
        scope_alias_applied = True
    elif scope == "main_body":
        part = _active_part(conn, document)
        value = _main_body(conn, part)
        resolved_scope = "active_part_main_body"
        scope_alias_applied = False
    else:
        value = _active_product(conn, document)
        resolved_scope = "active_product"
        scope_alias_applied = False

    return value, {
        "lookup_mode": "direct",
        "target_scope": scope,
        "resolved_scope": resolved_scope,
        "scope_alias_applied": scope_alias_applied,
        "requested_name": None,
        "resolved_name": _object_name(value),
        "match_count": 1,
        "occurrence": 1,
        "require_unique": True,
        "object": _describe_com_object(value),
    }, warnings


# ---------------------------------------------------------------------------
# CATIA SystemService.Evaluate helpers
# ---------------------------------------------------------------------------

def _get_system_service(application: Any) -> Any:
    try:
        return application.SystemService
    except Exception as exc:
        raise CATIAError(
            f"Cannot access CATIA SystemService: {exc}"
        ) from exc


def _evaluate(
    application: Any,
    script: str,
    function_name: str,
    parameters: list[Any],
) -> Any:
    service = _get_system_service(application)
    method = getattr(service, "Evaluate", None)
    if not callable(method):
        raise CATIAError(
            "CATIA SystemService does not expose Evaluate."
        )

    return method(
        script,
        _CATVB_SCRIPT_LANGUAGE,
        function_name,
        parameters,
    )


def _numeric_sequence(
    value: Any,
    minimum_length: int,
) -> list[float]:
    try:
        sequence = list(value)
    except Exception as exc:
        raise CATIAError(
            f"CATIA did not return an array: {exc}"
        ) from exc

    if len(sequence) < minimum_length:
        raise CATIAError(
            f"CATIA returned {len(sequence)} values; "
            f"{minimum_length} were required."
        )

    result: list[float] = []
    for item in sequence:
        number = float(item)
        if not math.isfinite(number):
            raise CATIAError(
                "CATIA returned a non-finite appearance value."
            )
        result.append(number)

    return result


# ---------------------------------------------------------------------------
# Appearance readback
# ---------------------------------------------------------------------------

def _read_selection_appearance(
    application: Any,
    selection: Any,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    result: dict[str, Any] = {
        "real_color": None,
        "visible_color": None,
        "real_opacity_255": None,
        "visible_opacity_255": None,
        "readback_method": (
            "CATIA.Application.SystemService.Evaluate "
            "with fixed in-memory CATVBScript"
        ),
        "external_macro_file": False,
        "user_script_input": False,
    }

    color_script = (
        "Public Function MCP_ReadVisColor(selectionObject)\n"
        "    Dim visProperties\n"
        "    Dim realR, realG, realB\n"
        "    Dim visibleR, visibleG, visibleB\n"
        "    Dim realStatus, visibleStatus\n"
        "    Set visProperties = selectionObject.VisProperties\n"
        "    realStatus = visProperties.GetRealColor("
        "realR, realG, realB)\n"
        "    visibleStatus = visProperties.GetVisibleColor("
        "visibleR, visibleG, visibleB)\n"
        "    MCP_ReadVisColor = Array("
        "realR, realG, realB, realStatus, "
        "visibleR, visibleG, visibleB, visibleStatus)\n"
        "End Function"
    )

    opacity_script = (
        "Public Function MCP_ReadVisOpacity(selectionObject)\n"
        "    Dim visProperties\n"
        "    Dim realOpacity, visibleOpacity\n"
        "    Dim realStatus, visibleStatus\n"
        "    Set visProperties = selectionObject.VisProperties\n"
        "    realStatus = visProperties.GetRealOpacity(realOpacity)\n"
        "    visibleStatus = "
        "visProperties.GetVisibleOpacity(visibleOpacity)\n"
        "    MCP_ReadVisOpacity = Array("
        "realOpacity, realStatus, visibleOpacity, visibleStatus)\n"
        "End Function"
    )

    try:
        values = _numeric_sequence(
            _evaluate(
                application,
                color_script,
                "MCP_ReadVisColor",
                [selection],
            ),
            8,
        )
        result["real_color"] = {
            "r": int(round(values[0])),
            "g": int(round(values[1])),
            "b": int(round(values[2])),
            "status_code": int(round(values[3])),
        }
        result["visible_color"] = {
            "r": int(round(values[4])),
            "g": int(round(values[5])),
            "b": int(round(values[6])),
            "status_code": int(round(values[7])),
        }
    except Exception as exc:
        warnings.append(f"Color readback failed: {exc}")

    try:
        values = _numeric_sequence(
            _evaluate(
                application,
                opacity_script,
                "MCP_ReadVisOpacity",
                [selection],
            ),
            4,
        )
        result["real_opacity_255"] = {
            "value": int(round(values[0])),
            "status_code": int(round(values[1])),
        }
        result["visible_opacity_255"] = {
            "value": int(round(values[2])),
            "status_code": int(round(values[3])),
        }
    except Exception as exc:
        warnings.append(f"Opacity readback failed: {exc}")

    return result, warnings


def _appearance_matches(
    appearance: dict[str, Any],
    rgb: tuple[int, int, int],
    opacity_255: int,
) -> tuple[bool, list[str]]:
    reasons: list[str] = []

    real_color = appearance.get("real_color")
    if not isinstance(real_color, dict):
        reasons.append("Real color could not be read back.")
    else:
        actual_rgb = (
            int(real_color.get("r", -1)),
            int(real_color.get("g", -1)),
            int(real_color.get("b", -1)),
        )
        if actual_rgb != rgb:
            reasons.append(
                f"Real color readback was {actual_rgb}, expected {rgb}."
            )

    real_opacity = appearance.get("real_opacity_255")
    if not isinstance(real_opacity, dict):
        reasons.append("Real opacity could not be read back.")
    else:
        actual_opacity = int(real_opacity.get("value", -1))
        if actual_opacity != opacity_255:
            reasons.append(
                f"Real opacity readback was {actual_opacity}, "
                f"expected {opacity_255}."
            )

    return not reasons, reasons


# ---------------------------------------------------------------------------
# Material catalog discovery
# ---------------------------------------------------------------------------

def _normalise_path(path_value: Any) -> str:
    text = os.path.expandvars(str(path_value).strip().strip('"'))
    return os.path.normpath(text) if text else ""


def _path_exists(application: Any, path: str) -> bool:
    try:
        file_system = application.FileSystem
        return bool(file_system.FileExists(path))
    except Exception:
        return Path(path).is_file()


def _split_environment_paths(value: Any) -> list[str]:
    text = str(value or "").strip()
    if not text:
        return []

    candidates: list[str] = []
    for item in text.replace("\n", ";").split(";"):
        path = _normalise_path(item)
        if path and path not in candidates:
            candidates.append(path)
    return candidates


def _catia_environment(
    application: Any,
    variable_name: str,
) -> tuple[Optional[str], Optional[str]]:
    try:
        value = application.SystemService.Environ(variable_name)
        return str(value), None
    except Exception as exc:
        return None, str(exc)


def _open_catmaterial_documents(
    application: Any,
) -> list[Any]:
    result: list[Any] = []
    documents = application.Documents

    try:
        count = int(documents.Count)
    except Exception:
        return result

    for index in range(1, count + 1):
        try:
            document = documents.Item(index)
            name = str(document.Name)
            if name.lower().endswith(".catmaterial"):
                result.append(document)
        except Exception:
            continue

    return result


def _document_full_path(document: Any) -> Optional[str]:
    try:
        full_name = str(document.FullName).strip()
        if full_name:
            return full_name
    except Exception:
        pass

    try:
        path = str(document.Path).strip()
        name = str(document.Name).strip()
        if path and name:
            return os.path.join(path, name)
    except Exception:
        pass

    return None


def _catalog_candidates(
    application: Any,
    explicit_path: str,
) -> tuple[list[str], dict[str, Any]]:
    diagnostics: dict[str, Any] = {
        "explicit_path": explicit_path or None,
        "environment_values": {},
        "candidates": [],
    }
    candidates: list[str] = []

    def add(path_value: Any) -> None:
        path = _normalise_path(path_value)
        if path and path not in candidates:
            candidates.append(path)

    if explicit_path:
        add(explicit_path)

    for open_document in _open_catmaterial_documents(application):
        full_path = _document_full_path(open_document)
        if full_path:
            add(full_path)

    environment_names = (
        "CATStartupPath",
        "CATDocView",
        "CATInstallPath",
        "CATReferenceSettingPath",
    )

    relative_candidates = (
        os.path.join("materials", "Catalog.CATMaterial"),
        os.path.join("startup", "materials", "Catalog.CATMaterial"),
        "Catalog.CATMaterial",
    )

    for variable_name in environment_names:
        value, error = _catia_environment(
            application,
            variable_name,
        )
        diagnostics["environment_values"][variable_name] = {
            "value": value,
            "error": error,
        }

        for root in _split_environment_paths(value):
            if root.lower().endswith(".catmaterial"):
                add(root)
            else:
                for relative in relative_candidates:
                    add(os.path.join(root, relative))

    # Common install roots are only fallback candidates; no release number is
    # hard-coded. Existing CATIA environment paths remain the primary source.
    for variable_name in ("ProgramFiles", "ProgramW6432"):
        root = os.environ.get(variable_name)
        if not root:
            continue

        dassault_root = os.path.join(root, "Dassault Systemes")
        try:
            for release_dir in Path(dassault_root).glob("*"):
                add(
                    release_dir
                    / "win_b64"
                    / "startup"
                    / "materials"
                    / "Catalog.CATMaterial"
                )
                add(
                    release_dir
                    / "intel_a"
                    / "startup"
                    / "materials"
                    / "Catalog.CATMaterial"
                )
        except Exception:
            continue

    diagnostics["candidates"] = [
        {
            "path": path,
            "exists": _path_exists(application, path),
        }
        for path in candidates
    ]

    return candidates, diagnostics


def _find_catalog_document(
    application: Any,
    catalog_path: str,
) -> tuple[Any, dict[str, Any], bool]:
    candidates, diagnostics = _catalog_candidates(
        application,
        catalog_path,
    )

    open_documents = _open_catmaterial_documents(application)
    open_by_path: dict[str, Any] = {}
    for document in open_documents:
        full_path = _document_full_path(document)
        if full_path:
            open_by_path[
                os.path.normcase(os.path.normpath(full_path))
            ] = document

    for candidate in candidates:
        if not _path_exists(application, candidate):
            continue

        key = os.path.normcase(os.path.normpath(candidate))
        if key in open_by_path:
            diagnostics["selected_path"] = candidate
            diagnostics["open_method"] = "reuse_open_CATMaterial"
            return open_by_path[key], diagnostics, False

        try:
            document = application.Documents.Read(candidate)
        except Exception:
            try:
                document = application.Documents.Open(candidate)
            except Exception as exc:
                diagnostics.setdefault("open_errors", []).append(
                    {
                        "path": candidate,
                        "error": str(exc),
                    }
                )
                continue

        diagnostics["selected_path"] = candidate
        diagnostics["open_method"] = "Documents.Read_or_Open"
        return document, diagnostics, True

    raise CATIAError(
        "No CATMaterial catalog could be located. Supply catalog_path "
        "explicitly or configure CATStartupPath so that "
        "materials\\Catalog.CATMaterial can be found."
    )


def _iter_materials(
    catalog_document: Any,
) -> list[tuple[str, str, Any]]:
    result: list[tuple[str, str, Any]] = []
    families = catalog_document.Families

    for family_index in range(1, int(families.Count) + 1):
        family = families.Item(family_index)
        family_name = _object_name(
            family,
            f"Family[{family_index}]",
        )
        materials = family.Materials

        for material_index in range(
            1,
            int(materials.Count) + 1,
        ):
            material = materials.Item(material_index)
            material_name = _object_name(
                material,
                f"Material[{material_index}]",
            )
            result.append(
                (family_name, material_name, material)
            )

    return result


def _find_material(
    catalog_document: Any,
    material_name: str,
    family_name: str,
) -> tuple[Any, dict[str, Any]]:
    requested_material = _normalise_nonempty(
        material_name,
        "material_name",
    )
    requested_family = str(family_name).strip()

    all_materials = _iter_materials(catalog_document)
    exact_matches: list[tuple[str, str, Any]] = []

    for family, material, value in all_materials:
        if requested_family and (
            family.casefold() != requested_family.casefold()
        ):
            continue
        if material.casefold() == requested_material.casefold():
            exact_matches.append((family, material, value))

    if not exact_matches:
        samples = [
            {
                "family": family,
                "material": material,
            }
            for family, material, _ in all_materials[:50]
        ]
        family_note = (
            f" in family '{requested_family}'"
            if requested_family
            else ""
        )
        error = CATIAError(
            f"Material '{requested_material}' was not found"
            f"{family_note}."
        )
        setattr(error, "available_material_samples", samples)
        raise error

    if len(exact_matches) > 1 and not requested_family:
        families = sorted(
            {family for family, _, _ in exact_matches}
        )
        raise CATIAError(
            f"Material name '{requested_material}' exists in multiple "
            f"families: {families}. Supply family_name."
        )

    family, material, value = exact_matches[0]
    return value, {
        "requested_material_name": requested_material,
        "requested_family_name": requested_family or None,
        "resolved_material_name": material,
        "resolved_family_name": family,
        "exact_case_insensitive_match": True,
        "catalog_material_count": len(all_materials),
    }


# ---------------------------------------------------------------------------
# Material manager and readback
# ---------------------------------------------------------------------------

def _get_material_manager(
    target_owner: Any,
) -> tuple[Any, dict[str, Any]]:
    errors: list[str] = []
    candidates = (
        target_owner,
        getattr(target_owner, "Parent", None),
    )

    for candidate in candidates:
        if candidate is None:
            continue
        try:
            manager = candidate.GetItem(
                _MATERIAL_MANAGER_EXTENSION
            )
            return manager, {
                "extension": _MATERIAL_MANAGER_EXTENSION,
                "owner": _describe_com_object(candidate),
            }
        except Exception as exc:
            errors.append(
                f"{type(candidate).__name__}.GetItem: {exc}"
            )

    raise CATIAError(
        "Cannot obtain CATMatManagerVBExt. "
        + "; ".join(errors)
    )


def _resolve_material_target(
    conn: Any,
    document: Any,
    apply_to: str,
) -> tuple[Any, Any, str, dict[str, Any]]:
    scope = _normalise_choice(
        apply_to,
        "apply_to",
        _ALLOWED_APPLY_TARGETS,
    )

    if scope == "part":
        part = _active_part(conn, document)
        manager_owner = part
        method_suffix = "Part"
        target = part
    elif scope == "main_body":
        part = _active_part(conn, document)
        manager_owner = part
        method_suffix = "Body"
        target = _main_body(conn, part)
    else:
        product = _active_product(conn, document)
        manager_owner = product
        method_suffix = "Product"
        target = product

    return target, manager_owner, method_suffix, {
        "apply_to": scope,
        "target_name": _object_name(target),
        "target": _describe_com_object(target),
        "manager_owner": _describe_com_object(manager_owner),
    }


def _material_name_readback(
    application: Any,
    manager: Any,
    target: Any,
    method_suffix: str,
) -> tuple[Optional[str], dict[str, Any], Optional[str]]:
    function_name = f"MCP_GetMaterialOn{method_suffix}"
    script = (
        f"Public Function {function_name}(managerObject, targetObject)\n"
        "    Dim appliedMaterial\n"
        "    Set appliedMaterial = Nothing\n"
        f"    managerObject.GetMaterialOn{method_suffix} "
        "targetObject, appliedMaterial\n"
        "    If appliedMaterial Is Nothing Then\n"
        f"        {function_name} = \"\"\n"
        "    Else\n"
        f"        {function_name} = appliedMaterial.Name\n"
        "    End If\n"
        "End Function"
    )

    diagnostics = {
        "method": f"GetMaterialOn{method_suffix}",
        "readback_method": (
            "CATIA.Application.SystemService.Evaluate "
            "with fixed in-memory CATVBScript"
        ),
        "external_macro_file": False,
        "user_script_input": False,
        "success": False,
    }

    try:
        value = _evaluate(
            application,
            script,
            function_name,
            [manager, target],
        )
        material_name = str(value or "").strip() or None
        diagnostics["success"] = True
        diagnostics["material_name"] = material_name
        return material_name, diagnostics, None
    except Exception as exc:
        diagnostics["error"] = str(exc)
        return None, diagnostics, str(exc)


def _catalog_document_key(document: Any) -> tuple[str, str]:
    full_path = _document_full_path(document)
    name = ""
    try:
        name = str(document.Name).strip()
    except Exception:
        pass

    return (
        os.path.normcase(os.path.normpath(full_path))
        if full_path
        else "",
        name.casefold(),
    )


def _catalog_saved_state(document: Any) -> Optional[bool]:
    try:
        return bool(document.Saved)
    except Exception:
        return None


def _active_document_name(application: Any) -> Optional[str]:
    try:
        return str(application.ActiveDocument.Name)
    except Exception:
        return None


def _find_open_document_by_key(
    application: Any,
    expected_path: str,
    expected_name: str,
) -> Optional[Any]:
    try:
        documents = application.Documents
        count = int(documents.Count)
    except Exception:
        return None

    for index in range(1, count + 1):
        try:
            candidate = documents.Item(index)
            candidate_path, candidate_name = _catalog_document_key(
                candidate
            )
            if expected_path and candidate_path == expected_path:
                return candidate
            if (
                not expected_path
                and expected_name
                and candidate_name == expected_name
            ):
                return candidate
        except Exception:
            continue

    return None


def _catalog_is_open_by_key(
    application: Any,
    expected_path: str,
    expected_name: str,
) -> bool:
    return (
        _find_open_document_by_key(
            application,
            expected_path,
            expected_name,
        )
        is not None
    )


def _catalog_is_open(
    application: Any,
    catalog_document: Any,
) -> bool:
    expected_path, expected_name = _catalog_document_key(
        catalog_document
    )
    return _catalog_is_open_by_key(
        application,
        expected_path,
        expected_name,
    )


def _activate_owner_document(
    application: Any,
    owner_document: Any,
) -> tuple[bool, Optional[str]]:
    if owner_document is None:
        return False, "No owner document was available."

    try:
        owner_document.Activate()
        return True, None
    except Exception as exc:
        return False, str(exc)


def _release_unused_com_references() -> dict[str, Any]:
    details: dict[str, Any] = {
        "gc_collect_called": True,
        "gc_collected_objects": None,
        "co_free_unused_libraries_called": False,
        "co_free_unused_libraries_error": None,
    }

    try:
        details["gc_collected_objects"] = int(gc.collect())
    except Exception as exc:
        details["gc_error"] = str(exc)

    try:
        import pythoncom  # type: ignore

        method = getattr(
            pythoncom,
            "CoFreeUnusedLibraries",
            None,
        )
        if callable(method):
            method()
            details["co_free_unused_libraries_called"] = True
    except Exception as exc:
        details["co_free_unused_libraries_error"] = str(exc)

    return details


def _close_catalog_via_evaluate(
    application: Any,
    catalog_document: Any,
) -> tuple[bool, Optional[str], Any]:
    function_name = "MCP_CloseCATMaterialDocument"
    script = (
        "Public Function MCP_CloseCATMaterialDocument(documentObject)\n"
        "    On Error Resume Next\n"
        "    Err.Clear\n"
        "    documentObject.Close\n"
        "    If Err.Number = 0 Then\n"
        "        MCP_CloseCATMaterialDocument = \"OK\"\n"
        "    Else\n"
        "        MCP_CloseCATMaterialDocument = "
        "\"ERROR:\" & CStr(Err.Number) & \":\" & Err.Description\n"
        "    End If\n"
        "    On Error GoTo 0\n"
        "End Function"
    )

    try:
        result = _evaluate(
            application,
            script,
            function_name,
            [catalog_document],
        )
        result_text = str(result or "").strip()
        if result_text.upper() == "OK":
            return True, None, result
        return False, result_text or "Evaluate returned no status.", result
    except Exception as exc:
        return False, str(exc), None


def _attempt_catalog_close(
    application: Any,
    expected_path: str,
    expected_name: str,
    strategy: str,
    document: Any,
) -> dict[str, Any]:
    attempt: dict[str, Any] = {
        "strategy": strategy,
        "attempted": True,
        "call_succeeded": False,
        "open_after_attempt": True,
        "error": None,
        "evaluate_result": None,
    }

    try:
        if strategy == "SystemService.Evaluate":
            (
                call_succeeded,
                error,
                evaluate_result,
            ) = _close_catalog_via_evaluate(
                application,
                document,
            )
            attempt["call_succeeded"] = call_succeeded
            attempt["error"] = error
            attempt["evaluate_result"] = evaluate_result
        else:
            document.Close()
            attempt["call_succeeded"] = True
    except Exception as exc:
        attempt["error"] = str(exc)

    attempt["open_after_attempt"] = _catalog_is_open_by_key(
        application,
        expected_path,
        expected_name,
    )
    attempt["verified_closed"] = not attempt[
        "open_after_attempt"
    ]

    return attempt


def _attempt_close_with_alerts_suppressed(
    application: Any,
    expected_path: str,
    expected_name: str,
    document: Any,
) -> dict[str, Any]:
    attempt: dict[str, Any] = {
        "strategy": "fresh_wrapper.DisplayFileAlerts_false",
        "attempted": True,
        "call_succeeded": False,
        "open_after_attempt": True,
        "error": None,
        "display_file_alerts_before": None,
        "display_file_alerts_restored": None,
    }

    previous_alerts: Optional[bool] = None

    try:
        previous_alerts = bool(application.DisplayFileAlerts)
        attempt["display_file_alerts_before"] = previous_alerts
    except Exception as exc:
        attempt["display_file_alerts_read_error"] = str(exc)

    try:
        application.DisplayFileAlerts = False
        document.Close()
        attempt["call_succeeded"] = True
    except Exception as exc:
        attempt["error"] = str(exc)
    finally:
        if previous_alerts is not None:
            try:
                application.DisplayFileAlerts = previous_alerts
                attempt["display_file_alerts_restored"] = True
            except Exception as exc:
                attempt["display_file_alerts_restored"] = False
                attempt["display_file_alerts_restore_error"] = str(
                    exc
                )

    attempt["open_after_attempt"] = _catalog_is_open_by_key(
        application,
        expected_path,
        expected_name,
    )
    attempt["verified_closed"] = not attempt[
        "open_after_attempt"
    ]

    return attempt


def _normalised_catalog_cache_key(
    expected_path: str,
    expected_name: str,
) -> str:
    if expected_path:
        return f"path:{expected_path}"
    if expected_name:
        return f"name:{expected_name.casefold()}"
    return ""


def _catalog_lifecycle_decision(
    link_mode: int,
    apply_started: bool,
    apply_verified: bool,
    rollback_succeeded: Optional[bool],
) -> dict[str, Any]:
    if int(link_mode) == 0:
        return {
            "policy": _CATALOG_POLICY_CLOSE_AFTER_COPY,
            "requires_close_if_opened_by_call": True,
            "keep_open_for_linked_material": False,
            "reason": (
                "link_mode=0 copies the material; a catalog opened by "
                "this call is not required after assignment."
            ),
        }

    if apply_verified:
        return {
            "policy": _CATALOG_POLICY_KEEP_OPEN_FOR_LINKED_MATERIAL,
            "requires_close_if_opened_by_call": False,
            "keep_open_for_linked_material": True,
            "reason": (
                "link_mode=1 creates a linked material dependency. "
                "The CATMaterial document is retained as a session-level "
                "cache and reused by later linked-material calls."
            ),
        }

    if not apply_started:
        return {
            "policy": _CATALOG_POLICY_CLOSE_AFTER_FAILED_LOOKUP,
            "requires_close_if_opened_by_call": True,
            "keep_open_for_linked_material": False,
            "reason": (
                "The material was not applied, so no linked dependency "
                "was created."
            ),
        }

    if rollback_succeeded is True:
        return {
            "policy": _CATALOG_POLICY_CLOSE_AFTER_ROLLBACK,
            "requires_close_if_opened_by_call": True,
            "keep_open_for_linked_material": False,
            "reason": (
                "The linked assignment was rolled back successfully, "
                "so no linked dependency remains."
            ),
        }

    return {
        "policy": _CATALOG_POLICY_KEEP_OPEN_FOR_UNRESOLVED_LINK,
        "requires_close_if_opened_by_call": False,
        "keep_open_for_linked_material": True,
        "reason": (
            "A link-mode assignment started but could not be verified or "
            "rolled back. The catalog is retained to avoid invalidating a "
            "possible linked material state."
        ),
    }


def _finalize_catalog_document(
    application: Any,
    catalog_document: Any,
    opened_by_tool: bool,
    owner_document: Any = None,
    lifecycle_policy: str = _CATALOG_POLICY_CLOSE_AFTER_COPY,
    requires_close_if_opened_by_call: bool = True,
    keep_open_reason: str = "",
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []

    expected_path, expected_name = (
        _catalog_document_key(catalog_document)
        if catalog_document is not None
        else ("", "")
    )
    cache_key = _normalised_catalog_cache_key(
        expected_path,
        expected_name,
    )

    # Remove stale cache metadata if CATIA no longer has the document open.
    if (
        cache_key
        and cache_key in _SESSION_LINKED_CATALOG_PATHS
        and not _catalog_is_open_by_key(
            application,
            expected_path,
            expected_name,
        )
    ):
        _SESSION_LINKED_CATALOG_PATHS.discard(cache_key)

    was_tool_session_cache = bool(
        cache_key
        and cache_key in _SESSION_LINKED_CATALOG_PATHS
        and not opened_by_tool
    )

    if opened_by_tool:
        ownership = "opened_by_current_call"
    elif was_tool_session_cache:
        ownership = "tool_session_cache"
    else:
        ownership = "preexisting_user_document"

    details: dict[str, Any] = {
        "lifecycle_policy": lifecycle_policy,
        "lifecycle_reason": keep_open_reason or None,
        "requires_close_if_opened_by_call": bool(
            requires_close_if_opened_by_call
        ),
        "catalog_ownership": ownership,
        "opened_by_tool": bool(opened_by_tool),
        "was_tool_session_cache": was_tool_session_cache,
        "catalog_path": expected_path or None,
        "catalog_name": expected_name or None,
        "catalog_saved_before_close": (
            _catalog_saved_state(catalog_document)
            if catalog_document is not None
            else None
        ),
        "active_document_before_finalize": (
            _active_document_name(application)
            if application is not None
            else None
        ),
        "owner_document_reactivated": False,
        "owner_document_activation_error": None,
        "python_reference_cleanup": None,
        "close_attempted": False,
        "close_succeeded": None,
        "selected_close_strategy": None,
        "close_attempts": [],
        "left_open_intentionally": False,
        "catalog_kept_open_by_design": False,
        "keep_open_reason": None,
        "session_cache_registered": bool(was_tool_session_cache),
        "session_cache_key": cache_key or None,
        "session_cache_reuse_expected": False,
        "open_after_finalize": None,
        "active_document_after_finalize": None,
    }

    if catalog_document is None:
        details["open_after_finalize"] = False
        return details, warnings

    activated, activation_error = _activate_owner_document(
        application,
        owner_document,
    )
    details["owner_document_reactivated"] = activated
    details["owner_document_activation_error"] = activation_error

    if activation_error:
        warnings.append(
            "The original model document could not be reactivated "
            f"during catalog lifecycle finalization: {activation_error}"
        )

    # A document not opened by this call is never closed automatically.
    # It may be a caller-owned document or a linked-material cache created
    # by an earlier call in this MCP process.
    if not opened_by_tool:
        details["left_open_intentionally"] = True
        details["catalog_kept_open_by_design"] = bool(
            was_tool_session_cache
            or not requires_close_if_opened_by_call
        )
        details["keep_open_reason"] = (
            "existing_tool_session_cache"
            if was_tool_session_cache
            else "preexisting_document_preserved"
        )
        details["session_cache_reuse_expected"] = bool(
            was_tool_session_cache
        )
        details["open_after_finalize"] = (
            _catalog_is_open_by_key(
                application,
                expected_path,
                expected_name,
            )
        )
        details["active_document_after_finalize"] = (
            _active_document_name(application)
        )
        return details, warnings

    # Linked-material success is intentionally retained. CATIA Documents is
    # the session cache and later calls reuse this same CATMaterial document.
    if not requires_close_if_opened_by_call:
        details["left_open_intentionally"] = True
        details["catalog_kept_open_by_design"] = True
        details["keep_open_reason"] = (
            keep_open_reason
            or "linked_material_session_dependency"
        )
        details["selected_close_strategy"] = (
            "session_documents_cache"
        )
        details["session_cache_reuse_expected"] = True
        if cache_key:
            _SESSION_LINKED_CATALOG_PATHS.add(cache_key)
            details["session_cache_registered"] = True

        details["open_after_finalize"] = (
            _catalog_is_open_by_key(
                application,
                expected_path,
                expected_name,
            )
        )
        details["active_document_after_finalize"] = (
            _active_document_name(application)
        )

        if not details["open_after_finalize"]:
            warnings.append(
                "The linked-material catalog was expected to remain open "
                "as a session cache, but it is no longer present in the "
                "CATIA Documents collection."
            )
        return details, warnings

    details["close_attempted"] = True
    details["python_reference_cleanup"] = (
        _release_unused_com_references()
    )

    # Strategy 1: original wrapper.
    attempt = _attempt_catalog_close(
        application,
        expected_path,
        expected_name,
        "original_wrapper.Close",
        catalog_document,
    )
    details["close_attempts"].append(attempt)

    if attempt["verified_closed"]:
        details["close_succeeded"] = True
        details["selected_close_strategy"] = attempt["strategy"]
    else:
        # Strategy 2: fresh Documents wrapper.
        fresh_document = _find_open_document_by_key(
            application,
            expected_path,
            expected_name,
        )
        if fresh_document is not None:
            attempt = _attempt_catalog_close(
                application,
                expected_path,
                expected_name,
                "fresh_documents_wrapper.Close",
                fresh_document,
            )
            details["close_attempts"].append(attempt)

            if attempt["verified_closed"]:
                details["close_succeeded"] = True
                details["selected_close_strategy"] = attempt[
                    "strategy"
                ]

    if details["close_succeeded"] is not True:
        # Strategy 3: execute Close inside CATIA.
        fresh_document = _find_open_document_by_key(
            application,
            expected_path,
            expected_name,
        )
        if fresh_document is not None:
            attempt = _attempt_catalog_close(
                application,
                expected_path,
                expected_name,
                "SystemService.Evaluate",
                fresh_document,
            )
            details["close_attempts"].append(attempt)

            if attempt["verified_closed"]:
                details["close_succeeded"] = True
                details["selected_close_strategy"] = attempt[
                    "strategy"
                ]

    if (
        details["close_succeeded"] is not True
        and details["catalog_saved_before_close"] is True
    ):
        # Strategy 4: alerts may only be suppressed for a saved catalog.
        fresh_document = _find_open_document_by_key(
            application,
            expected_path,
            expected_name,
        )
        if fresh_document is not None:
            attempt = _attempt_close_with_alerts_suppressed(
                application,
                expected_path,
                expected_name,
                fresh_document,
            )
            details["close_attempts"].append(attempt)

            if attempt["verified_closed"]:
                details["close_succeeded"] = True
                details["selected_close_strategy"] = attempt[
                    "strategy"
                ]

    details["open_after_finalize"] = (
        _catalog_is_open_by_key(
            application,
            expected_path,
            expected_name,
        )
    )
    details["active_document_after_finalize"] = (
        _active_document_name(application)
    )

    if details["open_after_finalize"]:
        details["close_succeeded"] = False
        attempt_errors = [
            {
                "strategy": item.get("strategy"),
                "error": item.get("error"),
            }
            for item in details["close_attempts"]
            if item.get("error")
        ]
        warnings.append(
            "CATMaterial catalog remains open after all required close "
            f"strategies. Errors: {attempt_errors}"
        )
    else:
        if cache_key:
            _SESSION_LINKED_CATALOG_PATHS.discard(cache_key)
        details["session_cache_registered"] = False

    return details, warnings


# ---------------------------------------------------------------------------
# MCP tools
# ---------------------------------------------------------------------------

def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = _get_connection(ctx)
    names: list[str] = []

    @mcp.tool()
    def catia_set_color(
        r: int,
        g: int,
        b: int,
        transparency: float = 0.0,
        target_scope: str = "main_body",
        target_object_name: str = "",
        occurrence: int = 1,
        require_unique: bool = True,
        inheritance: int = 1,
    ) -> dict[str, Any]:
        """Set RGB color and transparency on an active CATIA object.

        transparency is a fraction from 0.0 (fully opaque) to 1.0
        (fully transparent). CATIA stores opacity from 255 (opaque) to
        0 (transparent).

        target_scope:
        - main_body: active CATPart MainBody (default)
        - part: compatibility alias for active CATPart MainBody
        - product: active Product
        - named: an object resolved by target_object_name
        """

        warnings: list[str] = []
        document = None
        application = None
        previous_selection: list[Any] = []
        selection_snapshot_taken = False
        before: Optional[dict[str, Any]] = None
        after: Optional[dict[str, Any]] = None
        target_details: Optional[dict[str, Any]] = None
        vis_properties = None
        assignment_started = False
        result_payload: dict[str, Any] = {}
        result_error: Optional[str] = None
        result_status = "error"
        result_ok = False

        red = r
        green = g
        blue = b
        transparency_value = transparency
        opacity_255: Optional[int] = None
        inheritance_value = inheritance
        rollback_attempted = False
        rollback_succeeded: Optional[bool] = None

        try:
            red = _integer_in_range(r, "r", 0, 255)
            green = _integer_in_range(g, "g", 0, 255)
            blue = _integer_in_range(b, "b", 0, 255)
            transparency_value = _finite_number(
                transparency,
                "transparency",
            )
            if not 0.0 <= transparency_value <= 1.0:
                raise CATIAError(
                    "transparency must be from 0.0 "
                    "(fully opaque) to 1.0 "
                    "(fully transparent)."
                )

            inheritance_value = _integer_in_range(
                inheritance,
                "inheritance",
                0,
                1,
            )
            opacity_255 = int(
                round((1.0 - transparency_value) * 255.0)
            )

            application = _ensure_connected(conn)
            document = _active_document(conn)
            previous_selection, selection_warnings = (
                _snapshot_selection(document)
            )
            selection_snapshot_taken = True
            warnings.extend(selection_warnings)

            target, target_details, target_warnings = (
                _resolve_color_target(
                    conn,
                    document,
                    target_scope,
                    str(target_object_name),
                    occurrence,
                    bool(require_unique),
                )
            )
            warnings.extend(target_warnings)

            selection = document.Selection
            selection.Clear()
            selection.Add(target)

            before, before_warnings = _read_selection_appearance(
                application,
                selection,
            )
            warnings.extend(before_warnings)

            vis_properties = selection.VisProperties
            vis_properties.SetRealColor(
                red,
                green,
                blue,
                inheritance_value,
            )
            assignment_started = True
            vis_properties.SetRealOpacity(
                opacity_255,
                inheritance_value,
            )

            after, after_warnings = _read_selection_appearance(
                application,
                selection,
            )
            warnings.extend(after_warnings)

            verified, verification_reasons = _appearance_matches(
                after,
                (red, green, blue),
                opacity_255,
            )

            if not verified:
                rollback_attempted = True
                rollback_succeeded = False

                try:
                    old_color = (
                        before.get("real_color")
                        if before
                        else None
                    )
                    old_opacity = (
                        before.get("real_opacity_255")
                        if before
                        else None
                    )

                    if not isinstance(old_color, dict):
                        raise CATIAError(
                            "Previous real color was unavailable."
                        )
                    if not isinstance(old_opacity, dict):
                        raise CATIAError(
                            "Previous real opacity was unavailable."
                        )

                    vis_properties.SetRealColor(
                        int(old_color["r"]),
                        int(old_color["g"]),
                        int(old_color["b"]),
                        inheritance_value,
                    )
                    vis_properties.SetRealOpacity(
                        int(old_opacity["value"]),
                        inheritance_value,
                    )
                    rollback_succeeded = True
                except Exception as rollback_exc:
                    warnings.append(
                        f"Appearance rollback failed: {rollback_exc}"
                    )

                result_error = (
                    "CATIA appearance readback did not match the "
                    "requested values: "
                    + "; ".join(verification_reasons)
                )
                result_status = (
                    "error"
                    if rollback_succeeded
                    else "partial_success"
                )
                result_ok = False
                result_payload = {
                    "requested": {
                        "color_rgb": [red, green, blue],
                        "transparency": transparency_value,
                        "opacity_255": opacity_255,
                        "inheritance": inheritance_value,
                    },
                    "target": target_details,
                    "before": before,
                    "after_assignment": after,
                    "assignment_started": True,
                    "assignment_verified": False,
                    "rollback_attempted": True,
                    "rollback_succeeded": rollback_succeeded,
                    "feature_created": False,
                    "model_modified": (
                        rollback_succeeded is not True
                    ),
                    "document_save_required": (
                        rollback_succeeded is not True
                    ),
                }
            else:
                result_ok = True
                result_status = "success"
                result_payload = {
                    "requested": {
                        "color_rgb": [red, green, blue],
                        "transparency": transparency_value,
                        "opacity_255": opacity_255,
                        "inheritance": inheritance_value,
                    },
                    "target": target_details,
                    "before": before,
                    "after": after,
                    "assignment_started": True,
                    "assignment_verified": True,
                    "assignment_method": {
                        "selection": (
                            "ActiveDocument.Selection"
                        ),
                        "vis_properties": (
                            "Selection.VisProperties"
                        ),
                        "color_method": (
                            "VisPropertySet.SetRealColor"
                        ),
                        "opacity_method": (
                            "VisPropertySet.SetRealOpacity"
                        ),
                    },
                    "unit_semantics": {
                        "rgb": "integer 0..255",
                        "transparency": (
                            "fraction 0.0 opaque .. "
                            "1.0 transparent"
                        ),
                        "opacity_255": (
                            "CATIA opacity 0 transparent .. "
                            "255 opaque"
                        ),
                    },
                    "feature_created": False,
                    "model_modified": True,
                    "document_save_required": True,
                    "rollback_attempted": False,
                    "rollback_succeeded": None,
                }

            warnings.extend(_refresh_display(conn))
        except Exception as exc:
            if assignment_started and vis_properties is not None:
                rollback_attempted = True
                rollback_succeeded = False
                try:
                    old_color = (
                        before.get("real_color")
                        if before
                        else None
                    )
                    old_opacity = (
                        before.get("real_opacity_255")
                        if before
                        else None
                    )
                    if not isinstance(old_color, dict):
                        raise CATIAError(
                            "Previous real color was unavailable."
                        )
                    if not isinstance(old_opacity, dict):
                        raise CATIAError(
                            "Previous real opacity was unavailable."
                        )

                    vis_properties.SetRealColor(
                        int(old_color["r"]),
                        int(old_color["g"]),
                        int(old_color["b"]),
                        int(inheritance_value),
                    )
                    vis_properties.SetRealOpacity(
                        int(old_opacity["value"]),
                        int(inheritance_value),
                    )
                    rollback_succeeded = True
                except Exception as rollback_exc:
                    warnings.append(
                        f"Appearance rollback failed: {rollback_exc}"
                    )

            result_ok = False
            result_status = (
                "error"
                if (
                    not assignment_started
                    or rollback_succeeded is True
                )
                else "partial_success"
            )
            result_error = str(exc)
            result_payload = {
                "requested": {
                    "color_rgb": [r, g, b],
                    "transparency": transparency,
                    "opacity_255": opacity_255,
                    "target_scope": target_scope,
                    "target_object_name": (
                        target_object_name or None
                    ),
                    "inheritance": inheritance,
                },
                "target": target_details,
                "before": before,
                "after": after,
                "assignment_started": assignment_started,
                "assignment_verified": False,
                "rollback_attempted": rollback_attempted,
                "rollback_succeeded": rollback_succeeded,
                "feature_created": False,
                "model_modified": bool(
                    assignment_started
                    and rollback_succeeded is not True
                ),
                "document_save_required": bool(
                    assignment_started
                    and rollback_succeeded is not True
                ),
            }

        selection_state = {
            "snapshot_taken": selection_snapshot_taken,
            "before_count": len(previous_selection),
            "restored": None,
            "after_count": None,
            "initial_selection_was_empty": (
                len(previous_selection) == 0
                if selection_snapshot_taken
                else None
            ),
        }

        if document is not None and selection_snapshot_taken:
            restored, restore_warnings = _restore_selection(
                document,
                previous_selection,
            )
            warnings.extend(restore_warnings)
            selection_state["restored"] = restored
            selection_state["after_count"] = _selection_count(
                document
            )

            if not restored:
                selection_message = (
                    "The caller's previous CATIA Selection could not "
                    "be fully restored."
                )
                if result_ok:
                    result_ok = False
                    result_status = "partial_success"
                    result_error = selection_message
                else:
                    warnings.append(selection_message)

        result_payload["selection_state"] = selection_state

        if result_ok:
            return _success(result_payload, warnings)

        return _error(
            result_error or "Color assignment failed.",
            data=result_payload,
            warnings=warnings,
            status=result_status,
        )

    names.append("catia_set_color")

    @mcp.tool()
    def catia_set_material(
        material_name: str,
        family_name: str = "",
        catalog_path: str = "",
        apply_to: str = "part",
        link_mode: int = 1,
    ) -> dict[str, Any]:
        """Apply a CATIA material from a CATMaterial catalog.

        catalog_path is optional. When omitted, the tool checks open
        CATMaterial documents and CATIA environment paths such as
        CATStartupPath.

        apply_to may be part, main_body, or product.

        link_mode:
        - 0: copy material data. A CATMaterial opened by this call is
          closed after application.
        - 1: create/retain a linked material. One CATMaterial document is
          kept open as a session-level cache and reused by later calls.
        """

        warnings: list[str] = []
        application = None
        document = None
        catalog_document = None
        material = None
        target = None
        manager = None
        method_suffix: Optional[str] = None
        catalog_opened_by_tool = False
        catalog_diagnostics: Optional[dict[str, Any]] = None
        catalog_cleanup: Optional[dict[str, Any]] = None
        material_details: Optional[dict[str, Any]] = None
        target_details: Optional[dict[str, Any]] = None
        manager_details: Optional[dict[str, Any]] = None
        before_name: Optional[str] = None
        after_name: Optional[str] = None
        before_readback: Optional[dict[str, Any]] = None
        after_readback: Optional[dict[str, Any]] = None
        apply_started = False
        result_ok = False
        result_status = "error"
        result_error: Optional[str] = None
        result_payload: dict[str, Any] = {}
        rollback_attempted = False
        rollback_succeeded: Optional[bool] = None

        requested_material = material_name
        requested_family = family_name
        requested_catalog = catalog_path
        requested_link_mode = link_mode

        try:
            requested_material = _normalise_nonempty(
                material_name,
                "material_name",
            )
            requested_family = str(family_name).strip()
            requested_catalog = _normalise_path(catalog_path)
            requested_link_mode = _integer_in_range(
                link_mode,
                "link_mode",
                0,
                1,
            )

            application = _ensure_connected(conn)
            document = _active_document(conn)

            (
                target,
                manager_owner,
                method_suffix,
                target_details,
            ) = _resolve_material_target(
                conn,
                document,
                apply_to,
            )

            manager, manager_details = _get_material_manager(
                manager_owner
            )

            (
                before_name,
                before_readback,
                before_error,
            ) = _material_name_readback(
                application,
                manager,
                target,
                method_suffix,
            )
            if before_error:
                warnings.append(
                    "Existing material could not be read before "
                    f"assignment: {before_error}"
                )

            (
                catalog_document,
                catalog_diagnostics,
                catalog_opened_by_tool,
            ) = _find_catalog_document(
                application,
                requested_catalog,
            )

            material, material_details = _find_material(
                catalog_document,
                requested_material,
                requested_family,
            )

            apply_method_name = (
                f"ApplyMaterialOn{method_suffix}"
            )
            apply_method = getattr(
                manager,
                apply_method_name,
                None,
            )
            if not callable(apply_method):
                raise CATIAError(
                    f"Material manager does not expose "
                    f"{apply_method_name}."
                )

            apply_method(
                target,
                material,
                requested_link_mode,
            )
            apply_started = True

            (
                after_name,
                after_readback,
                after_error,
            ) = _material_name_readback(
                application,
                manager,
                target,
                method_suffix,
            )

            verified = bool(
                after_name
                and after_name.casefold()
                == material_details[
                    "resolved_material_name"
                ].casefold()
            )

            if not verified:
                rollback_attempted = True
                rollback_succeeded = False

                if before_name is None:
                    try:
                        apply_method(
                            target,
                            None,
                            requested_link_mode,
                        )
                        rollback_succeeded = True
                    except Exception as rollback_exc:
                        warnings.append(
                            f"Material rollback failed: {rollback_exc}"
                        )
                else:
                    warnings.append(
                        "A prior material existed. Automatic rollback "
                        "was not attempted because its original catalog "
                        "object was not available."
                    )

                verification_message = (
                    after_error
                    or (
                        f"Material readback was {after_name!r}, "
                        f"expected "
                        f"{material_details['resolved_material_name']!r}."
                    )
                )

                result_ok = False
                result_status = (
                    "error"
                    if rollback_succeeded
                    else "partial_success"
                )
                result_error = (
                    "Material application could not be verified: "
                    + verification_message
                )
            else:
                result_ok = True
                result_status = "success"

            result_payload = {
                "requested": {
                    "material_name": requested_material,
                    "family_name": (
                        requested_family or None
                    ),
                    "catalog_path": (
                        requested_catalog or None
                    ),
                    "apply_to": apply_to,
                    "link_mode": requested_link_mode,
                },
                "catalog": catalog_diagnostics,
                "material": material_details,
                "target": target_details,
                "manager": manager_details,
                "before_material_name": before_name,
                "after_material_name": after_name,
                "before_readback": before_readback,
                "after_readback": after_readback,
                "apply_started": apply_started,
                "apply_verified": bool(result_ok),
                "apply_method": (
                    f"CATMatManagerVBExt."
                    f"ApplyMaterialOn{method_suffix}"
                ),
                "readback_method": (
                    f"CATMatManagerVBExt."
                    f"GetMaterialOn{method_suffix} "
                    "through SystemService.Evaluate"
                ),
                "catalog_document_opened_by_tool": (
                    catalog_opened_by_tool
                ),
                "feature_created": False,
                "model_modified": (
                    True
                    if result_ok
                    else (
                        rollback_succeeded is not True
                        and apply_started
                    )
                ),
                "document_save_required": (
                    True
                    if result_ok
                    else (
                        rollback_succeeded is not True
                        and apply_started
                    )
                ),
                "rollback_attempted": rollback_attempted,
                "rollback_succeeded": rollback_succeeded,
            }

            warnings.extend(_refresh_display(conn))
        except Exception as exc:
            samples = getattr(
                exc,
                "available_material_samples",
                None,
            )
            result_ok = False
            result_status = (
                "partial_success"
                if apply_started
                else "error"
            )
            result_error = str(exc)
            result_payload = {
                "requested": {
                    "material_name": material_name,
                    "family_name": family_name or None,
                    "catalog_path": catalog_path or None,
                    "apply_to": apply_to,
                    "link_mode": link_mode,
                },
                "catalog": catalog_diagnostics,
                "material": material_details,
                "available_material_samples": samples,
                "target": target_details,
                "manager": manager_details,
                "before_material_name": before_name,
                "after_material_name": after_name,
                "before_readback": before_readback,
                "after_readback": after_readback,
                "apply_started": apply_started,
                "apply_verified": False,
                "feature_created": False,
                "model_modified": apply_started,
                "document_save_required": apply_started,
                "rollback_attempted": rollback_attempted,
                "rollback_succeeded": rollback_succeeded,
            }

        lifecycle_link_mode = (
            requested_link_mode
            if (
                isinstance(requested_link_mode, int)
                and not isinstance(requested_link_mode, bool)
                and requested_link_mode in (0, 1)
            )
            else 0
        )
        catalog_lifecycle = _catalog_lifecycle_decision(
            lifecycle_link_mode,
            bool(apply_started),
            bool(result_ok),
            rollback_succeeded,
        )

        result_payload["catalog_lifecycle_policy"] = (
            catalog_lifecycle["policy"]
        )
        result_payload["catalog_lifecycle_reason"] = (
            catalog_lifecycle["reason"]
        )
        result_payload["catalog_kept_open_by_design"] = bool(
            catalog_lifecycle["keep_open_for_linked_material"]
        )

        if (
            result_ok
            and catalog_lifecycle[
                "keep_open_for_linked_material"
            ]
        ):
            warnings.append(
                "link_mode=1 retains the CATMaterial document as a "
                "session-level linked-material cache. This is expected "
                "behavior; subsequent linked-material calls must reuse "
                "the same open document without increasing the document "
                "count."
            )

        material_reference_was_held = material is not None
        material = None
        # _find_material no longer holds its internal family/material list
        # after return. Dropping the selected Material wrapper here allows
        # CATIA to close the CATMaterial document.
        reference_release = _release_unused_com_references()

        if application is not None and catalog_document is not None:
            catalog_cleanup, cleanup_warnings = (
                _finalize_catalog_document(
                    application,
                    catalog_document,
                    catalog_opened_by_tool,
                    owner_document=document,
                    lifecycle_policy=catalog_lifecycle["policy"],
                    requires_close_if_opened_by_call=bool(
                        catalog_lifecycle[
                            "requires_close_if_opened_by_call"
                        ]
                    ),
                    keep_open_reason=catalog_lifecycle["reason"],
                )
            )
            warnings.extend(cleanup_warnings)
        else:
            catalog_cleanup = {
                "lifecycle_policy": catalog_lifecycle["policy"],
                "lifecycle_reason": catalog_lifecycle["reason"],
                "requires_close_if_opened_by_call": bool(
                    catalog_lifecycle[
                        "requires_close_if_opened_by_call"
                    ]
                ),
                "catalog_ownership": None,
                "opened_by_tool": bool(
                    catalog_opened_by_tool
                ),
                "was_tool_session_cache": False,
                "catalog_path": None,
                "catalog_name": None,
                "catalog_saved_before_close": None,
                "active_document_before_finalize": None,
                "owner_document_reactivated": False,
                "owner_document_activation_error": None,
                "python_reference_cleanup": None,
                "close_attempted": False,
                "close_succeeded": None,
                "selected_close_strategy": None,
                "close_attempts": [],
                "left_open_intentionally": False,
                "catalog_kept_open_by_design": False,
                "keep_open_reason": None,
                "session_cache_registered": False,
                "session_cache_key": None,
                "session_cache_reuse_expected": False,
                "open_after_finalize": False,
                "active_document_after_finalize": None,
            }

        result_payload["material_reference_release"] = {
            "material_reference_was_held": (
                material_reference_was_held
            ),
            "material_reference_set_to_none": True,
            "python_reference_cleanup": reference_release,
        }
        result_payload["catalog_cleanup"] = catalog_cleanup

        post_cleanup_material_name: Optional[str] = None
        post_cleanup_readback: Optional[dict[str, Any]] = None
        post_cleanup_error: Optional[str] = None
        post_cleanup_verified: Optional[bool] = None

        if (
            apply_started
            and application is not None
            and manager is not None
            and target is not None
            and method_suffix is not None
        ):
            (
                post_cleanup_material_name,
                post_cleanup_readback,
                post_cleanup_error,
            ) = _material_name_readback(
                application,
                manager,
                target,
                method_suffix,
            )
            expected_material_name = (
                material_details.get("resolved_material_name")
                if material_details
                else after_name
            )
            post_cleanup_verified = bool(
                expected_material_name
                and post_cleanup_material_name
                and post_cleanup_material_name.casefold()
                == str(expected_material_name).casefold()
            )

        result_payload["post_cleanup_material_name"] = (
            post_cleanup_material_name
        )
        result_payload["post_cleanup_readback"] = (
            post_cleanup_readback
        )
        result_payload["post_cleanup_error"] = (
            post_cleanup_error
        )
        result_payload["post_cleanup_verified"] = (
            post_cleanup_verified
        )

        close_was_required = bool(
            catalog_cleanup.get(
                "requires_close_if_opened_by_call"
            )
        )
        cleanup_failed = bool(
            catalog_cleanup.get("opened_by_tool")
            and close_was_required
            and (
                catalog_cleanup.get("close_succeeded") is not True
                or catalog_cleanup.get("open_after_finalize")
            )
        )

        linked_cache_missing = bool(
            catalog_cleanup.get("opened_by_tool")
            and not close_was_required
            and catalog_cleanup.get(
                "catalog_kept_open_by_design"
            )
            and not catalog_cleanup.get("open_after_finalize")
        )

        if cleanup_failed:
            cleanup_message = (
                "The material operation completed, but the CATMaterial "
                "document that this lifecycle policy required to close "
                "remains open."
            )
            if result_ok:
                result_ok = False
                result_status = "partial_success"
                result_error = cleanup_message
            else:
                warnings.append(cleanup_message)

        if linked_cache_missing:
            cache_message = (
                "The linked material was applied, but its CATMaterial "
                "session cache is not open after lifecycle finalization."
            )
            if result_ok:
                result_ok = False
                result_status = "partial_success"
                result_error = cache_message
            else:
                warnings.append(cache_message)

        if (
            apply_started
            and post_cleanup_verified is not True
        ):
            persistence_message = (
                "The applied material could not be verified after "
                "catalog lifecycle finalization."
            )
            if post_cleanup_error:
                persistence_message += (
                    f" Readback error: {post_cleanup_error}"
                )
            if result_ok:
                result_ok = False
                result_status = "partial_success"
                result_error = persistence_message
            else:
                warnings.append(persistence_message)

        if result_ok:
            return _success(result_payload, warnings)

        return _error(
            result_error or "Material assignment failed.",
            data=result_payload,
            warnings=warnings,
            status=result_status,
        )

    names.append("catia_set_material")

    return names

