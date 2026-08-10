"""
catia_measurement_tools.py
Version: catia-measurement-tools-fixed-2026-07-30-v1

Focused CATIA V5 MCP tools for:
- minimum distance between two named measurable references
- radius of a named circular/cylindrical/spherical measurable reference

Important behavior:
- SPAWorkbench is obtained from the active Document.
- Measurable objects are created from CATIA References.
- Named-object lookup rejects ambiguous duplicates by default.
- A CATIA generic reference label may be supplied with lookup_mode="reference".
- Selection is cleared on every success and error path.
- The thickness tool reports minimum separation. It is only a wall-thickness
  interpretation when the caller deliberately supplies opposing wall faces or
  equivalent extracted surfaces.
- No temporary Measure features or geometry are created.
"""

from __future__ import annotations

import math
from typing import Any, Optional

from catia_mcp.connection import CATIAError


IMPLEMENTATION_VERSION = (
    "catia-measurement-tools-fixed-2026-07-30-v1"
)

_ALLOWED_LOOKUP_MODES = {"name", "reference"}


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
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "implementation_version": IMPLEMENTATION_VERSION,
        "ok": False,
        "status": "error",
        "error": str(message),
        "warnings": list(warnings or []),
    }
    if data is not None:
        result["data"] = data
    return result


# ---------------------------------------------------------------------------
# Connection and COM helpers
# ---------------------------------------------------------------------------

def _get_connection(ctx: Any) -> Any:
    return getattr(ctx, "conn", ctx)


def _ensure_connected(conn: Any) -> Any:
    method = getattr(conn, "ensure_connected", None)
    if callable(method):
        return method()
    return None


def _get_application(conn: Any) -> Any:
    app = getattr(conn, "app", None)
    if app is None:
        app = getattr(conn, "_app", None)

    if app is not None:
        return app

    connected = _ensure_connected(conn)
    if connected is not None:
        return connected

    raise CATIAError("Cannot access the CATIA Application object.")


def _active_document(conn: Any) -> Any:
    app = _get_application(conn)
    try:
        document = app.ActiveDocument
    except Exception as exc:
        raise CATIAError(
            f"Cannot access CATIA.ActiveDocument: {exc}"
        ) from exc

    if document is None:
        raise CATIAError("CATIA has no active document.")

    return document


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
            "The active document is not a CATPart document."
        ) from exc


def _dispatch_if_possible(value: Any) -> tuple[Any, dict[str, Any]]:
    details: dict[str, Any] = {
        "dispatch_used": False,
        "raw_python_type": type(value).__name__,
        "raw_python_module": type(value).__module__,
    }

    try:
        import win32com.client  # type: ignore

        dispatched = win32com.client.Dispatch(value)
        details["dispatch_used"] = True
        details["resolved_python_type"] = type(dispatched).__name__
        details["resolved_python_module"] = type(dispatched).__module__
        return dispatched, details
    except Exception as exc:
        details["dispatch_error"] = str(exc)
        details["resolved_python_type"] = type(value).__name__
        details["resolved_python_module"] = type(value).__module__
        return value, details


def _get_spa_workbench(
    document: Any,
) -> tuple[Any, dict[str, Any]]:
    try:
        raw_spa = document.GetWorkbench("SPAWorkbench")
    except Exception as exc:
        raise CATIAError(
            "Cannot obtain SPAWorkbench from the active Document. "
            "ActiveDocument.GetWorkbench('SPAWorkbench') failed: "
            f"{exc}"
        ) from exc

    spa, dispatch_details = _dispatch_if_possible(raw_spa)
    return spa, {
        "acquisition": (
            "ActiveDocument.GetWorkbench('SPAWorkbench')"
        ),
        "interface": dispatch_details,
    }


def _get_measurable(
    spa: Any,
    reference: Any,
) -> tuple[Any, dict[str, Any]]:
    candidates: list[tuple[str, Any]] = [("raw", spa)]
    dispatched, dispatch_details = _dispatch_if_possible(spa)
    if dispatched is not spa:
        candidates.append(("dispatched", dispatched))

    errors: list[str] = []
    for candidate_name, candidate in candidates:
        method = getattr(candidate, "GetMeasurable", None)
        if not callable(method):
            continue

        try:
            raw_measurable = method(reference)
            measurable, measurable_dispatch = (
                _dispatch_if_possible(raw_measurable)
            )
            return measurable, {
                "method": "GetMeasurable",
                "candidate": candidate_name,
                "spa_dispatch": dispatch_details,
                "measurable_dispatch": measurable_dispatch,
            }
        except Exception as exc:
            errors.append(
                f"{candidate_name}.GetMeasurable: {exc}"
            )

    detail = (
        "; ".join(errors)
        if errors
        else "GetMeasurable is not exposed."
    )
    raise CATIAError(
        f"Cannot create a Measurable object: {detail}"
    )


def _clear_selection(document: Any) -> None:
    try:
        document.Selection.Clear()
    except Exception:
        pass


def _selection_count(selection: Any) -> int:
    for attribute in ("Count2", "Count"):
        try:
            return int(getattr(selection, attribute))
        except Exception:
            continue
    raise CATIAError("Cannot read CATIA Selection count.")


def _selection_item(selection: Any, index: int) -> Any:
    for method_name in ("Item2", "Item"):
        method = getattr(selection, method_name, None)
        if callable(method):
            try:
                return method(index)
            except Exception:
                continue
    raise CATIAError(
        f"Cannot access Selection item {index}."
    )


def _object_name(value: Any, fallback: str = "") -> str:
    try:
        name = str(value.Name).strip()
        return name or fallback
    except Exception:
        return fallback


def _object_type(value: Any) -> dict[str, str]:
    result = {
        "python_type": type(value).__name__,
        "python_module": type(value).__module__,
    }

    try:
        result["catia_name"] = str(value.Name)
    except Exception:
        pass

    try:
        result["catia_type"] = str(value.Type)
    except Exception:
        pass

    return result


# ---------------------------------------------------------------------------
# Validation and reference resolution
# ---------------------------------------------------------------------------

def _non_empty_text(value: Any, parameter_name: str) -> str:
    text = str(value).strip()
    if not text:
        raise CATIAError(f"{parameter_name} cannot be empty.")
    return text


def _positive_integer(value: Any, parameter_name: str) -> int:
    if isinstance(value, bool):
        raise CATIAError(
            f"{parameter_name} must be a positive integer."
        )

    try:
        integer = int(value)
        numeric = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(
            f"{parameter_name} must be a positive integer."
        ) from exc

    if integer <= 0 or numeric != float(integer):
        raise CATIAError(
            f"{parameter_name} must be a positive integer."
        )

    return integer


def _normalise_lookup_mode(value: Any) -> str:
    mode = str(value).strip().lower()
    if mode not in _ALLOWED_LOOKUP_MODES:
        allowed = ", ".join(sorted(_ALLOWED_LOOKUP_MODES))
        raise CATIAError(
            f"lookup_mode must be one of: {allowed}."
        )
    return mode


def _search_named_object(
    document: Any,
    identifier: str,
    occurrence: int,
    require_unique: bool,
) -> tuple[Any, dict[str, Any]]:
    selection = document.Selection
    _clear_selection(document)

    try:
        selection.Search(f"Name={identifier},all")
        match_count = _selection_count(selection)

        if match_count == 0:
            raise CATIAError(
                f"Element not found by name: {identifier}"
            )

        if bool(require_unique) and match_count != 1:
            raise CATIAError(
                f"Element name '{identifier}' is ambiguous: "
                f"{match_count} objects matched. Rename the objects "
                "uniquely or call with require_unique=false and an "
                "explicit occurrence."
            )

        if occurrence > match_count:
            raise CATIAError(
                f"Occurrence {occurrence} was requested for "
                f"'{identifier}', but only {match_count} objects "
                "matched."
            )

        selected = _selection_item(selection, occurrence)
        try:
            obj = selected.Value
        except Exception as exc:
            raise CATIAError(
                f"Cannot obtain the selected CATIA object: {exc}"
            ) from exc

        return obj, {
            "lookup_mode": "name",
            "requested_identifier": identifier,
            "resolved_name": _object_name(obj, identifier),
            "match_count": match_count,
            "occurrence": occurrence,
            "require_unique": bool(require_unique),
            "object": _object_type(obj),
        }
    finally:
        _clear_selection(document)


def _reference_from_generic_name(
    part: Any,
    identifier: str,
) -> tuple[Any, dict[str, Any]]:
    try:
        reference = part.CreateReferenceFromName(identifier)
    except Exception as exc:
        raise CATIAError(
            "CATIA could not create a Reference from the supplied "
            f"generic label: {exc}"
        ) from exc

    if reference is None:
        raise CATIAError(
            "CreateReferenceFromName returned no Reference."
        )

    return reference, {
        "lookup_mode": "reference",
        "requested_identifier": identifier,
        "match_count": None,
        "occurrence": None,
        "require_unique": None,
        "reference_created": True,
    }


def _resolve_reference(
    document: Any,
    part: Any,
    identifier: Any,
    *,
    occurrence: Any,
    require_unique: bool,
    lookup_mode: Any,
) -> tuple[Any, dict[str, Any]]:
    resolved_identifier = _non_empty_text(
        identifier,
        "element identifier",
    )
    mode = _normalise_lookup_mode(lookup_mode)

    if mode == "reference":
        return _reference_from_generic_name(
            part,
            resolved_identifier,
        )

    resolved_occurrence = _positive_integer(
        occurrence,
        "occurrence",
    )
    obj, details = _search_named_object(
        document,
        resolved_identifier,
        resolved_occurrence,
        bool(require_unique),
    )

    try:
        reference = part.CreateReferenceFromObject(obj)
    except Exception as exc:
        raise CATIAError(
            "Cannot create a CATIA Reference from the named "
            f"object '{resolved_identifier}': {exc}"
        ) from exc

    details["reference_created"] = True
    return reference, details


# ---------------------------------------------------------------------------
# Measurement helpers
# ---------------------------------------------------------------------------

def _finite_property(
    value: Any,
    property_name: str,
) -> tuple[Optional[float], Optional[str]]:
    try:
        result = getattr(value, property_name)
        if callable(result):
            result = result()
        number = float(result)
    except Exception as exc:
        return None, str(exc)

    if not math.isfinite(number):
        return None, (
            f"{property_name} returned a non-finite value."
        )

    return number, None


def _geometry_name(
    measurable: Any,
) -> tuple[Optional[int], Optional[str]]:
    try:
        value = getattr(measurable, "GeometryName")
        if callable(value):
            value = value()
        return int(value), None
    except Exception as exc:
        return None, str(exc)


def _minimum_distance(
    measurable: Any,
    second_reference: Any,
) -> float:
    method = getattr(
        measurable,
        "GetMinimumDistance",
        None,
    )
    if not callable(method):
        dispatched, _ = _dispatch_if_possible(measurable)
        method = getattr(
            dispatched,
            "GetMinimumDistance",
            None,
        )

    if not callable(method):
        raise CATIAError(
            "Measurable.GetMinimumDistance is not available."
        )

    try:
        distance = float(method(second_reference))
    except Exception as exc:
        raise CATIAError(
            "GetMinimumDistance failed. CATIA does not support "
            "minimum-distance measurement between some container "
            "objects, including Body and HybridBody. Use measurable "
            "points, curves, edges, faces, planes or surfaces instead. "
            f"Detail: {exc}"
        ) from exc

    if not math.isfinite(distance):
        raise CATIAError(
            "GetMinimumDistance returned a non-finite value."
        )

    if distance < 0.0:
        raise CATIAError(
            "GetMinimumDistance returned a negative value."
        )

    return distance


def _radius(
    measurable: Any,
) -> float:
    radius, radius_error = _finite_property(
        measurable,
        "Radius",
    )

    if radius is None:
        geometry_code, geometry_error = _geometry_name(
            measurable
        )
        geometry_detail = (
            f" GeometryName={geometry_code}."
            if geometry_code is not None
            else (
                f" GeometryName could not be read: "
                f"{geometry_error}."
            )
        )
        raise CATIAError(
            "The selected object does not expose a measurable "
            "Radius. Radius is supported for circular/arc, "
            "cylindrical or spherical measurable geometry."
            f"{geometry_detail} Detail: {radius_error}"
        )

    if radius <= 0.0:
        raise CATIAError(
            f"Measurable.Radius returned {radius}; a positive "
            "radius was expected."
        )

    return radius


# ---------------------------------------------------------------------------
# MCP registration
# ---------------------------------------------------------------------------

def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = _get_connection(ctx)
    names: list[str] = []

    @mcp.tool()
    def catia_measure_thickness(
        face1_name: str,
        face2_name: str,
        occurrence1: int = 1,
        occurrence2: int = 1,
        require_unique: bool = True,
        lookup_mode: str = "name",
    ) -> dict[str, Any]:
        """Measure minimum distance between two CATPart references.

        The legacy tool name uses "thickness", but the CATIA operation is
        minimum distance between two references. It is a wall-thickness result
        only when face1_name and face2_name deliberately identify opposing wall
        faces or equivalent extracted surfaces.

        lookup_mode="name" resolves specification-tree objects by exact name.
        lookup_mode="reference" treats both identifiers as CATIA generic
        reference labels accepted by Part.CreateReferenceFromName.
        """

        warnings: list[str] = []
        document = None

        try:
            identifier1 = _non_empty_text(
                face1_name,
                "face1_name",
            )
            identifier2 = _non_empty_text(
                face2_name,
                "face2_name",
            )
            mode = _normalise_lookup_mode(lookup_mode)

            _ensure_connected(conn)
            document = _active_document(conn)
            part = _active_part(conn, document)
            _clear_selection(document)

            reference1, resolution1 = _resolve_reference(
                document,
                part,
                identifier1,
                occurrence=occurrence1,
                require_unique=bool(require_unique),
                lookup_mode=mode,
            )
            reference2, resolution2 = _resolve_reference(
                document,
                part,
                identifier2,
                occurrence=occurrence2,
                require_unique=bool(require_unique),
                lookup_mode=mode,
            )

            spa, spa_details = _get_spa_workbench(document)
            measurable1, measurable_details = _get_measurable(
                spa,
                reference1,
            )
            distance_mm = _minimum_distance(
                measurable1,
                reference2,
            )

            geometry1, geometry1_error = _geometry_name(
                measurable1
            )

            geometry2 = None
            geometry2_error = None
            try:
                measurable2, _ = _get_measurable(
                    spa,
                    reference2,
                )
                geometry2, geometry2_error = _geometry_name(
                    measurable2
                )
            except Exception as exc:
                geometry2_error = str(exc)

            warnings.append(
                "This result is the minimum distance between the "
                "two supplied references. It is not a full variable "
                "wall-thickness analysis and does not prove uniform "
                "thickness over either face."
            )

            return _success(
                {
                    "thickness_mm": distance_mm,
                    "minimum_distance_mm": distance_mm,
                    "measurement_kind": (
                        "minimum_distance_between_references"
                    ),
                    "wall_thickness_interpretation": (
                        "valid_only_for_deliberately_paired_"
                        "opposing_faces_or_surfaces"
                    ),
                    "is_full_wall_thickness_analysis": False,
                    "element1": {
                        **resolution1,
                        "measurable_geometry_code": geometry1,
                        "geometry_code_error": geometry1_error,
                    },
                    "element2": {
                        **resolution2,
                        "measurable_geometry_code": geometry2,
                        "geometry_code_error": geometry2_error,
                    },
                    "units": {
                        "thickness_mm": "mm",
                        "minimum_distance_mm": "mm",
                    },
                    "measurement_method": {
                        "spa": spa_details,
                        "measurable": measurable_details,
                        "distance_method": (
                            "Measurable.GetMinimumDistance"
                        ),
                    },
                    "selection_cleared": True,
                    "feature_created": False,
                    "model_modified": False,
                },
                warnings,
            )
        except Exception as exc:
            return _error(
                str(exc),
                data={
                    "face1_name": str(face1_name),
                    "face2_name": str(face2_name),
                    "lookup_mode": str(lookup_mode),
                    "selection_cleared": True,
                    "feature_created": False,
                    "model_modified": False,
                },
                warnings=warnings,
            )
        finally:
            if document is not None:
                _clear_selection(document)

    names.append("catia_measure_thickness")

    @mcp.tool()
    def catia_measure_radius(
        edge_or_circle_name: str,
        occurrence: int = 1,
        require_unique: bool = True,
        lookup_mode: str = "name",
    ) -> dict[str, Any]:
        """Measure radius of a circular, cylindrical or spherical reference.

        lookup_mode="name" resolves a uniquely named specification-tree
        object. For unstable topology, create and name an extracted edge/face,
        or use lookup_mode="reference" with a CATIA generic reference label.
        """

        warnings: list[str] = []
        document = None

        try:
            identifier = _non_empty_text(
                edge_or_circle_name,
                "edge_or_circle_name",
            )
            mode = _normalise_lookup_mode(lookup_mode)

            _ensure_connected(conn)
            document = _active_document(conn)
            part = _active_part(conn, document)
            _clear_selection(document)

            reference, resolution = _resolve_reference(
                document,
                part,
                identifier,
                occurrence=occurrence,
                require_unique=bool(require_unique),
                lookup_mode=mode,
            )

            spa, spa_details = _get_spa_workbench(document)
            measurable, measurable_details = _get_measurable(
                spa,
                reference,
            )

            geometry_code, geometry_error = _geometry_name(
                measurable
            )
            radius_mm = _radius(measurable)

            if mode == "name":
                warnings.append(
                    "CATIA topological Edge/Face names may change "
                    "after model updates. Prefer a uniquely named "
                    "extracted curve/surface or a validated generic "
                    "reference label for persistent automation."
                )

            return _success(
                {
                    "radius_mm": radius_mm,
                    "diameter_mm": radius_mm * 2.0,
                    "measurement_kind": (
                        "radius_of_measurable_geometry"
                    ),
                    "element": {
                        **resolution,
                        "measurable_geometry_code": geometry_code,
                        "geometry_code_error": geometry_error,
                    },
                    "units": {
                        "radius_mm": "mm",
                        "diameter_mm": "mm",
                    },
                    "supported_geometry_note": (
                        "Measurable.Radius is intended for "
                        "circle/arc, cylinder or sphere geometry."
                    ),
                    "measurement_method": {
                        "spa": spa_details,
                        "measurable": measurable_details,
                        "radius_property": "Measurable.Radius",
                    },
                    "selection_cleared": True,
                    "feature_created": False,
                    "model_modified": False,
                },
                warnings,
            )
        except Exception as exc:
            return _error(
                str(exc),
                data={
                    "edge_or_circle_name": str(
                        edge_or_circle_name
                    ),
                    "lookup_mode": str(lookup_mode),
                    "selection_cleared": True,
                    "feature_created": False,
                    "model_modified": False,
                },
                warnings=warnings,
            )
        finally:
            if document is not None:
                _clear_selection(document)

    names.append("catia_measure_radius")

    return names


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
