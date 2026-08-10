"""
advanced_surface_continuity.py
Version: advanced-surface-continuity-fixed-2026-07-30-v2

Important v2 correction:
- CATIA Measurable.Area is returned in square metres.
- area_m2 preserves the original CATIA value.
- area_mm2 is converted using area_m2 * 1_000_000.
"""

from __future__ import annotations

import math
from typing import Any, Optional

from catia_mcp.connection import CATIAError


IMPLEMENTATION_VERSION = "advanced-surface-continuity-fixed-2026-07-30-v2"


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


def _get_connection(ctx: Any) -> Any:
    return getattr(ctx, "conn", ctx)


def _ensure_connected(conn: Any) -> None:
    method = getattr(conn, "ensure_connected", None)
    if callable(method):
        method()


def _active_document(conn: Any) -> Any:
    app = getattr(conn, "app", None)
    if app is None:
        app = getattr(conn, "_app", None)

    if app is not None:
        try:
            return app.ActiveDocument
        except Exception:
            pass

    getter = getattr(conn, "get_active_document", None)
    if callable(getter):
        try:
            return getter()
        except Exception:
            pass

    raise CATIAError("Cannot access the active CATIA document.")


def _get_active_part(conn: Any) -> Any:
    getter = getattr(conn, "get_active_part", None)
    if callable(getter):
        return getter()

    document = _active_document(conn)
    try:
        return document.Part
    except Exception as exc:
        raise CATIAError(
            "The active CATIA document is not a CATPart document."
        ) from exc


def _normalise_name(value: Any, default: str, parameter_name: str) -> str:
    text = str(value).strip()
    if text:
        return text

    fallback = str(default).strip()
    if fallback:
        return fallback

    raise CATIAError(f"{parameter_name} cannot be empty.")


def _finite_number(value: Any, name: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(f"{name} must be a number.") from exc

    if not math.isfinite(number):
        raise CATIAError(f"{name} must be finite.")

    return number


def _positive_number(value: Any, name: str) -> float:
    number = _finite_number(value, name)
    if number <= 0.0:
        raise CATIAError(f"{name} must be greater than 0.")
    return number


def _bounded_number(
    value: Any,
    name: str,
    minimum: float,
    maximum: float,
) -> float:
    number = _finite_number(value, name)
    if number < minimum or number > maximum:
        raise CATIAError(
            f"{name} must be between {minimum:g} and {maximum:g}."
        )
    return number


def _signed_orientation(value: Any, name: str) -> int:
    if isinstance(value, bool):
        raise CATIAError(f"{name} must be 1 or -1.")

    try:
        integer = int(value)
        numeric = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(f"{name} must be 1 or -1.") from exc

    if numeric != float(integer) or integer not in (-1, 1):
        raise CATIAError(f"{name} must be 1 or -1.")

    return integer


def _describe_com_object(value: Any) -> dict[str, Any]:
    return {
        "python_type": type(value).__name__,
        "python_module": type(value).__module__,
        "has_oleobj": bool(hasattr(value, "_oleobj_")),
    }


def _has_callable(value: Any, method_name: str) -> bool:
    try:
        return callable(getattr(value, method_name, None))
    except Exception:
        return False


def _dispatch_com(value: Any) -> Any:
    import win32com.client  # type: ignore

    return win32com.client.Dispatch(value)


def _resolve_com_interface(
    raw_object: Any,
    required_methods: list[str],
    interface_name: str,
) -> tuple[Any, dict[str, Any]]:
    details: dict[str, Any] = {
        "interface": interface_name,
        "required_methods": list(required_methods),
        "raw": _describe_com_object(raw_object),
        "dispatch_used": False,
    }

    if all(_has_callable(raw_object, method) for method in required_methods):
        details["resolved"] = _describe_com_object(raw_object)
        details["required_methods_available"] = True
        return raw_object, details

    try:
        resolved = _dispatch_com(raw_object)
    except Exception as exc:
        details["dispatch_error"] = str(exc)
        raise CATIAError(
            f"{interface_name} was returned as a generic COM object and "
            f"could not be dynamically dispatched: {exc}"
        ) from exc

    missing_methods = [
        method
        for method in required_methods
        if not _has_callable(resolved, method)
    ]

    details["dispatch_used"] = True
    details["resolved"] = _describe_com_object(resolved)
    details["required_methods_available"] = not missing_methods

    if missing_methods:
        raise CATIAError(
            f"{interface_name} dispatch completed, but the resolved COM "
            f"object does not expose: {', '.join(missing_methods)}."
        )

    return resolved, details


def _get_hybrid_shape_factory(
    part: Any,
    required_methods: list[str],
) -> tuple[Any, dict[str, Any]]:
    try:
        raw_factory = part.HybridShapeFactory
    except Exception as exc:
        raise CATIAError(
            f"Cannot access Part.HybridShapeFactory: {exc}"
        ) from exc

    return _resolve_com_interface(
        raw_factory,
        required_methods,
        "HybridShapeFactory",
    )


def _get_hybrid_bodies_collection(
    part: Any,
) -> tuple[Any, dict[str, Any]]:
    try:
        raw_collection = part.HybridBodies
    except Exception as exc:
        raise CATIAError(f"Cannot access Part.HybridBodies: {exc}") from exc

    return _resolve_com_interface(
        raw_collection,
        ["Item", "Add"],
        "HybridBodies",
    )


def _object_name(obj: Any, fallback: str = "") -> str:
    try:
        name = str(obj.Name).strip()
        return name or fallback
    except Exception:
        return fallback


def _set_object_name(obj: Any, requested_name: str) -> list[str]:
    warnings: list[str] = []
    try:
        obj.Name = requested_name
    except Exception as exc:
        warnings.append(
            f"Object was created but could not be renamed to "
            f"'{requested_name}': {exc}"
        )
    return warnings


def _get_or_create_hybrid_body(
    part: Any,
    requested_name: str,
) -> tuple[Any, bool, dict[str, Any], list[str]]:
    body_name = _normalise_name(
        requested_name,
        "BridgeSurfaces",
        "geometrical_set",
    )
    warnings: list[str] = []

    collection, collection_details = _get_hybrid_bodies_collection(part)

    existing_body = None
    try:
        existing_body = collection.Item(body_name)
    except Exception:
        existing_body = None

    if existing_body is not None:
        body, body_interface = _resolve_com_interface(
            existing_body,
            ["AppendHybridShape"],
            "HybridBody",
        )
        return (
            body,
            False,
            {
                "name": _object_name(body, body_name),
                "created": False,
                "collection": collection_details,
                "body_interface": body_interface,
            },
            warnings,
        )

    try:
        raw_body = collection.Add()
    except Exception as exc:
        raise CATIAError(
            f"Cannot create geometrical set '{body_name}': {exc}"
        ) from exc

    body, body_interface = _resolve_com_interface(
        raw_body,
        ["AppendHybridShape"],
        "HybridBody",
    )
    warnings.extend(_set_object_name(body, body_name))

    return (
        body,
        True,
        {
            "name": _object_name(body, body_name),
            "requested_name": body_name,
            "created": True,
            "collection": collection_details,
            "body_interface": body_interface,
        },
        warnings,
    )


def _set_in_work_object(part: Any, obj: Any) -> None:
    try:
        part.InWorkObject = obj
    except Exception as exc:
        raise CATIAError(
            f"Cannot set Part.InWorkObject to the target container: {exc}"
        ) from exc


def _find_named_object(part: Any, object_name: str) -> Any:
    name = _normalise_name(object_name, "", "object name")

    try:
        obj = part.FindObjectByName(name)
    except Exception as exc:
        raise CATIAError(
            f"CATIA could not search for object '{name}': {exc}"
        ) from exc

    if obj is None:
        raise CATIAError(f"Object '{name}' was not found in the active CATPart.")

    return obj


def _create_reference(
    part: Any,
    obj: Any,
    requested_name: str,
) -> tuple[Any, dict[str, Any]]:
    try:
        reference = part.CreateReferenceFromObject(obj)
    except Exception as exc:
        raise CATIAError(
            f"Cannot create a CATIA Reference from '{requested_name}': {exc}"
        ) from exc

    if reference is None:
        raise CATIAError(
            f"CATIA returned an empty Reference for '{requested_name}'."
        )

    return reference, {
        "requested_name": requested_name,
        "resolved_name": _object_name(obj, requested_name),
        "object": _describe_com_object(obj),
        "reference": _describe_com_object(reference),
    }


def _resolve_named_reference(
    part: Any,
    object_name: str,
) -> tuple[Any, Any, dict[str, Any]]:
    obj = _find_named_object(part, object_name)
    reference, details = _create_reference(part, obj, object_name)
    return obj, reference, details


def _clear_selection(conn: Any) -> None:
    try:
        selection = _active_document(conn).Selection
        selection.Clear()
    except Exception:
        pass


def _delete_object(conn: Any, obj: Any) -> bool:
    if obj is None:
        return True

    try:
        document = _active_document(conn)
        selection = document.Selection
        selection.Clear()
        selection.Add(obj)
        selection.Delete()
        selection.Clear()
        return True
    except Exception:
        _clear_selection(conn)
        return False


def _update_feature(
    part: Any,
    feature: Any,
) -> tuple[bool, str, list[str]]:
    warnings: list[str] = []

    try:
        part.UpdateObject(feature)
        return True, "UpdateObject", warnings
    except Exception as exc:
        warnings.append(f"UpdateObject failed: {exc}")

    try:
        part.Update()
        warnings.append("Part.Update fallback succeeded.")
        return True, "Part.Update", warnings
    except Exception as exc:
        warnings.append(f"Part.Update fallback failed: {exc}")
        return False, "failed", warnings


def _refresh_display(conn: Any) -> list[str]:
    warnings: list[str] = []
    method = getattr(conn, "refresh_display", None)

    if callable(method):
        try:
            method()
        except Exception as exc:
            warnings.append(f"Display refresh failed: {exc}")

    return warnings


def _rollback_created_objects(
    conn: Any,
    feature: Any,
    feature_appended: bool,
    container: Any,
    container_created: bool,
) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []

    if feature is None:
        feature_rollback_succeeded = True
    elif feature_appended:
        feature_rollback_succeeded = _delete_object(conn, feature)
    else:
        feature_rollback_succeeded = True

    feature_persisted = bool(
        feature is not None
        and feature_appended
        and not feature_rollback_succeeded
    )

    container_rollback_attempted = bool(
        container_created and not feature_persisted
    )
    container_rollback_succeeded: Optional[bool] = None

    if container_rollback_attempted:
        container_rollback_succeeded = _delete_object(conn, container)
        if not container_rollback_succeeded:
            warnings.append(
                "The newly created geometrical set could not be removed "
                "after the feature failure."
            )

    rollback_succeeded = bool(
        feature_rollback_succeeded
        and container_rollback_succeeded is not False
    )

    return (
        {
            "feature_created": feature is not None,
            "feature_appended": feature_appended,
            "feature_persisted": feature_persisted,
            "feature_rollback_succeeded": feature_rollback_succeeded,
            "container_created": container_created,
            "container_rollback_attempted": container_rollback_attempted,
            "container_rollback_succeeded": container_rollback_succeeded,
            "rollback_succeeded": rollback_succeeded,
        },
        warnings,
    )


def _operation_failure(
    conn: Any,
    message: str,
    feature: Any,
    feature_appended: bool,
    container: Any,
    container_created: bool,
    feature_type: str,
    warnings: Optional[list[str]] = None,
    extra_data: Optional[dict[str, Any]] = None,
) -> dict[str, Any]:
    warning_list = list(warnings or [])

    rollback_data, rollback_warnings = _rollback_created_objects(
        conn,
        feature,
        feature_appended,
        container,
        container_created,
    )
    warning_list.extend(rollback_warnings)

    data: dict[str, Any] = {
        "type": feature_type,
        "update_succeeded": False,
        **rollback_data,
    }
    if extra_data:
        data.update(extra_data)

    status = "partial_success" if data["feature_persisted"] else "error"
    if data["feature_persisted"]:
        warning_list.append(
            "The operation failed and the created feature could not be "
            "removed; the active CATPart may have been modified."
        )

    return _error(
        message,
        data=data,
        warnings=warning_list,
        status=status,
    )


def _continuity_request(
    continuity_level: Any,
    allow_g3_degrade: bool,
) -> tuple[str, str, int, list[str]]:
    requested = str(continuity_level).strip().upper()
    warnings: list[str] = []

    if requested not in {"G1", "G2", "G3"}:
        raise CATIAError("continuity_level must be one of: G1, G2, G3.")

    if requested == "G1":
        return requested, "G1", 1, warnings

    if requested == "G2":
        return requested, "G2", 2, warnings

    if not bool(allow_g3_degrade):
        raise CATIAError(
            "G3 is not exposed by the standard HybridShapeBlend Automation "
            "continuity values. Set allow_g3_degrade=true to create G2 instead."
        )

    warnings.append(
        "G3 was requested, but standard HybridShapeBlend Automation exposes "
        "only G0/G1/G2 continuity. The request was explicitly degraded to G2."
    )
    return requested, "G2", 2, warnings


def _coupling_value(coupling_mode: Any) -> tuple[str, int]:
    text = str(coupling_mode).strip().lower().replace("_", "").replace("-", "")

    mapping = {
        "ratio": ("Ratio", 1),
        "tangency": ("Tangency", 2),
        "curvature": ("TangencyThenCurvature", 3),
        "tangencythencurvature": ("TangencyThenCurvature", 3),
        "vertices": ("Vertices", 4),
        "vertex": ("Vertices", 4),
    }

    if text not in mapping:
        raise CATIAError(
            "coupling_mode must be one of: Ratio, Tangency, "
            "TangencyThenCurvature, Curvature, Vertices."
        )

    return mapping[text]


def _parameter_value(value: Any) -> Any:
    if value is None:
        return None

    for attribute in ("Value", "ValueAsString"):
        try:
            result = getattr(value, attribute)
            if callable(result):
                result = result()
            if result is not None:
                try:
                    return float(result)
                except (TypeError, ValueError):
                    return result
        except Exception:
            continue

    try:
        return float(value)
    except Exception:
        return str(value)


def _read_blend_definition(blend: Any) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    definition: dict[str, Any] = {}

    try:
        definition["coupling"] = int(blend.Coupling)
    except Exception as exc:
        warnings.append(f"Could not read Blend.Coupling: {exc}")

    limits: list[dict[str, Any]] = []
    for limit in (1, 2):
        limit_data: dict[str, Any] = {"limit": limit}

        getters = {
            "continuity": "GetContinuity",
            "orientation": "GetOrientation",
            "transition": "GetTransition",
            "tension_type": "GetTensionType",
            "trim_support": "GetTrimSupport",
        }
        for key, method_name in getters.items():
            try:
                limit_data[key] = int(getattr(blend, method_name)(limit))
            except Exception as exc:
                warnings.append(
                    f"Could not read {method_name} for limit {limit}: {exc}"
                )

        try:
            tension_param = blend.GetTensionInDouble(limit, 1)
            limit_data["tension"] = _parameter_value(tension_param)
        except Exception as exc:
            warnings.append(
                f"Could not read GetTensionInDouble for limit {limit}: {exc}"
            )

        limits.append(limit_data)

    definition["limits"] = limits
    return definition, warnings


def _try_numeric_property(
    obj: Any,
    property_name: str,
) -> tuple[Optional[float], Optional[str]]:
    try:
        value = getattr(obj, property_name)
        if callable(value):
            value = value()
        number = float(value)
        if math.isfinite(number):
            return number, None
        return None, f"{property_name} returned a non-finite value."
    except Exception as exc:
        return None, str(exc)


def _try_cog(measurable: Any) -> tuple[Optional[list[float]], Optional[str]]:
    try:
        values = measurable.GetCOG()
        if values is None:
            return None, "GetCOG returned no values."
        result = [float(values[index]) for index in range(3)]
        return result, None
    except Exception as exc:
        return None, str(exc)


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = _get_connection(ctx)
    names: list[str] = []

    @mcp.tool()
    def create_g2_bridge_surface(
        edge1_name: str,
        support1_name: str,
        edge2_name: str,
        support2_name: str,
        continuity_level: str = "G2",
        tension1: float = 1.0,
        tension2: float = 1.0,
        coupling_mode: str = "Ratio",
        orientation1: int = 1,
        orientation2: int = 1,
        transition1: int = 1,
        transition2: int = 1,
        trim_supports: bool = False,
        feature_name: str = "MCP_G2_Bridge",
        geometrical_set: str = "BridgeSurfaces",
        allow_g3_degrade: bool = True,
    ) -> dict[str, Any]:
        """Create a CATIA HybridShapeBlend between two curves on support surfaces.

        G1 and G2 are native HybridShapeBlend continuity values. G3 can only be
        handled by an explicit, reported degradation to G2.
        """

        bridge = None
        feature_appended = False
        hybrid_body = None
        hybrid_body_created = False
        warnings: list[str] = []

        try:
            edge1 = _normalise_name(edge1_name, "", "edge1_name")
            support1 = _normalise_name(support1_name, "", "support1_name")
            edge2 = _normalise_name(edge2_name, "", "edge2_name")
            support2 = _normalise_name(support2_name, "", "support2_name")
            output_name = _normalise_name(
                feature_name,
                "MCP_G2_Bridge",
                "feature_name",
            )
            body_name = _normalise_name(
                geometrical_set,
                "BridgeSurfaces",
                "geometrical_set",
            )

            requested_continuity, achieved_continuity, continuity_value, continuity_warnings = (
                _continuity_request(
                    continuity_level,
                    bool(allow_g3_degrade),
                )
            )
            warnings.extend(continuity_warnings)

            tension_1 = _bounded_number(tension1, "tension1", 0.1, 10.0)
            tension_2 = _bounded_number(tension2, "tension2", 0.1, 10.0)
            coupling_name, coupling_value = _coupling_value(coupling_mode)
            orientation_1 = _signed_orientation(orientation1, "orientation1")
            orientation_2 = _signed_orientation(orientation2, "orientation2")
            transition_1 = _signed_orientation(transition1, "transition1")
            transition_2 = _signed_orientation(transition2, "transition2")

            _ensure_connected(conn)
            _clear_selection(conn)
            part = _get_active_part(conn)

            factory, factory_details = _get_hybrid_shape_factory(
                part,
                ["AddNewBlend"],
            )
            (
                hybrid_body,
                hybrid_body_created,
                body_details,
                body_warnings,
            ) = _get_or_create_hybrid_body(part, body_name)
            warnings.extend(body_warnings)
            _set_in_work_object(part, hybrid_body)

            _, edge1_reference, edge1_details = _resolve_named_reference(
                part,
                edge1,
            )
            _, support1_reference, support1_details = _resolve_named_reference(
                part,
                support1,
            )
            _, edge2_reference, edge2_details = _resolve_named_reference(
                part,
                edge2,
            )
            _, support2_reference, support2_details = _resolve_named_reference(
                part,
                support2,
            )

            raw_bridge = factory.AddNewBlend()
            bridge, bridge_interface = _resolve_com_interface(
                raw_bridge,
                [
                    "SetCurve",
                    "SetSupport",
                    "SetContinuity",
                    "SetOrientation",
                    "SetTransition",
                    "SetTensionInDouble",
                    "SetTrimSupport",
                    "GetContinuity",
                    "GetOrientation",
                    "GetTransition",
                    "GetTensionInDouble",
                    "GetTensionType",
                    "GetTrimSupport",
                ],
                "HybridShapeBlend",
            )

            bridge.SetCurve(1, edge1_reference)
            bridge.SetSupport(1, support1_reference)
            bridge.SetContinuity(1, continuity_value)
            bridge.SetOrientation(1, orientation_1)
            bridge.SetTransition(1, transition_1)
            bridge.SetTensionInDouble(1, 2, tension_1, tension_1)
            bridge.SetTrimSupport(1, 2 if bool(trim_supports) else 1)

            bridge.SetCurve(2, edge2_reference)
            bridge.SetSupport(2, support2_reference)
            bridge.SetContinuity(2, continuity_value)
            bridge.SetOrientation(2, orientation_2)
            bridge.SetTransition(2, transition_2)
            bridge.SetTensionInDouble(2, 2, tension_2, tension_2)
            bridge.SetTrimSupport(2, 2 if bool(trim_supports) else 1)

            try:
                bridge.Coupling = coupling_value
            except Exception as exc:
                raise CATIAError(
                    f"Cannot set HybridShapeBlend.Coupling to "
                    f"{coupling_value}: {exc}"
                ) from exc

            warnings.extend(_set_object_name(bridge, output_name))
            hybrid_body.AppendHybridShape(bridge)
            feature_appended = True

            updated, update_strategy, update_warnings = _update_feature(
                part,
                bridge,
            )
            warnings.extend(update_warnings)

            common_data = {
                "name": output_name,
                "type": "HybridShapeBlend",
                "edge1": edge1_details,
                "support1": support1_details,
                "edge2": edge2_details,
                "support2": support2_details,
                "continuity_requested": requested_continuity,
                "continuity_achieved": achieved_continuity,
                "continuity_value": continuity_value,
                "tension1": tension_1,
                "tension2": tension_2,
                "coupling_mode": coupling_name,
                "coupling_value": coupling_value,
                "orientation1": orientation_1,
                "orientation2": orientation_2,
                "transition1": transition_1,
                "transition2": transition_2,
                "trim_supports": bool(trim_supports),
                "geometrical_set": body_details,
                "factory": factory_details,
                "feature_interface": bridge_interface,
                "api_methods": [
                    "AddNewBlend",
                    "SetCurve",
                    "SetSupport",
                    "SetContinuity",
                    "SetOrientation",
                    "SetTransition",
                    "SetTensionInDouble",
                    "SetTrimSupport",
                ],
                "update_strategy": update_strategy,
            }

            if not updated:
                return _operation_failure(
                    conn,
                    "Blend was created and appended, but CATIA could not "
                    "update the feature. Verify that each curve lies on its "
                    "corresponding support surface and that both limits are "
                    "geometrically compatible.",
                    bridge,
                    feature_appended,
                    hybrid_body,
                    hybrid_body_created,
                    "HybridShapeBlend",
                    warnings,
                    common_data,
                )

            definition, definition_warnings = _read_blend_definition(bridge)
            warnings.extend(definition_warnings)
            warnings.extend(_refresh_display(conn))
            _clear_selection(conn)

            return _success(
                {
                    **common_data,
                    "name": _object_name(bridge, output_name),
                    "blend_definition": definition,
                    "feature_created": True,
                    "feature_appended": True,
                    "feature_persisted": True,
                    "update_succeeded": True,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            _clear_selection(conn)
            return _operation_failure(
                conn,
                str(exc),
                bridge,
                feature_appended,
                hybrid_body,
                hybrid_body_created,
                "HybridShapeBlend",
                warnings,
            )

    names.append("create_g2_bridge_surface")

    @mcp.tool()
    def analyze_surface_quality(
        surface_name: str,
        analysis_type: str = "both",
        geometrical_set: Optional[str] = None,
    ) -> dict[str, Any]:
        """Return honest Automation-accessible surface measurements.

        Standard CATIA V5 Automation does not expose a quantitative zebra-stripe
        or curvature-comb evaluator through SPAWorkbench. This tool therefore
        reports measurable area/perimeter/COG and, for HybridShapeBlend objects,
        reads the configured blend continuity, tension, orientation and coupling.
        CATIA Measurable.Area is preserved as area_m2 and converted to area_mm2.
        It never fabricates zebra scores, curvature values or A-Class grades.
        """

        warnings: list[str] = []

        try:
            requested_name = _normalise_name(
                surface_name,
                "",
                "surface_name",
            )
            requested_analysis = (
                str(analysis_type).strip().lower().replace("-", "_")
            )

            allowed_analysis = {
                "measurable",
                "continuity",
                "both",
                "zebra",
                "curvature_comb",
            }
            if requested_analysis not in allowed_analysis:
                raise CATIAError(
                    "analysis_type must be one of: measurable, continuity, "
                    "both, zebra, curvature_comb."
                )

            _ensure_connected(conn)
            _clear_selection(conn)
            document = _active_document(conn)
            part = _get_active_part(conn)

            surface_obj = None
            resolved_container: Optional[str] = None

            if geometrical_set is not None and str(geometrical_set).strip():
                body_name = _normalise_name(
                    geometrical_set,
                    "",
                    "geometrical_set",
                )
                try:
                    body = part.HybridBodies.Item(body_name)
                    surface_obj = body.HybridShapes.Item(requested_name)
                    resolved_container = body_name
                except Exception as exc:
                    raise CATIAError(
                        f"Surface '{requested_name}' was not found in "
                        f"geometrical set '{body_name}': {exc}"
                    ) from exc
            else:
                surface_obj = _find_named_object(part, requested_name)

            surface_reference, reference_details = _create_reference(
                part,
                surface_obj,
                requested_name,
            )

            measurable_data: dict[str, Any] = {
                "available": False,
                "area_m2": None,
                "area_mm2": None,
                "perimeter_mm": None,
                "center_of_gravity_mm": None,
                "units": {
                    "area_m2": "m^2",
                    "area_mm2": "mm^2",
                    "perimeter_mm": "mm",
                    "center_of_gravity_mm": "mm",
                },
            }
            measurable_interface: Optional[dict[str, Any]] = None

            try:
                raw_spa = document.GetWorkbench("SPAWorkbench")
                spa, spa_interface = _resolve_com_interface(
                    raw_spa,
                    ["GetMeasurable"],
                    "SPAWorkbench",
                )
                raw_measurable = spa.GetMeasurable(surface_reference)
                measurable, measurable_interface = _resolve_com_interface(
                    raw_measurable,
                    [],
                    "Measurable",
                )

                measurable_data["available"] = True
                measurable_data["spa_interface"] = spa_interface

                area_m2, area_error = _try_numeric_property(
                    measurable,
                    "Area",
                )
                measurable_data["area_m2"] = area_m2
                measurable_data["area_mm2"] = (
                    area_m2 * 1_000_000.0
                    if area_m2 is not None
                    else None
                )
                if area_error:
                    warnings.append(
                        f"Surface area could not be read: {area_error}"
                    )

                perimeter, perimeter_error = _try_numeric_property(
                    measurable,
                    "Perimeter",
                )
                measurable_data["perimeter_mm"] = perimeter
                if perimeter_error:
                    warnings.append(
                        f"Surface perimeter could not be read: "
                        f"{perimeter_error}"
                    )

                cog, cog_error = _try_cog(measurable)
                measurable_data["center_of_gravity_mm"] = cog
                if cog_error:
                    warnings.append(
                        f"Surface center of gravity could not be read: "
                        f"{cog_error}"
                    )
            except Exception as exc:
                warnings.append(
                    f"SPAWorkbench measurable summary was unavailable: {exc}"
                )

            blend_definition: Optional[dict[str, Any]] = None
            blend_interface: Optional[dict[str, Any]] = None

            try:
                blend, blend_interface = _resolve_com_interface(
                    surface_obj,
                    [
                        "GetContinuity",
                        "GetOrientation",
                        "GetTransition",
                        "GetTensionInDouble",
                        "GetTensionType",
                        "GetTrimSupport",
                    ],
                    "HybridShapeBlend",
                )
                blend_definition, blend_warnings = _read_blend_definition(
                    blend
                )
                warnings.extend(blend_warnings)
            except Exception:
                if requested_analysis in {
                    "continuity",
                    "both",
                    "zebra",
                    "curvature_comb",
                }:
                    warnings.append(
                        "The selected object does not expose the "
                        "HybridShapeBlend definition interface; configured "
                        "blend continuity could not be read."
                    )

            if requested_analysis in {"zebra", "curvature_comb", "both"}:
                warnings.append(
                    "Quantitative zebra-stripe and curvature-comb evaluation "
                    "is not exposed by the standard SPAWorkbench Measurable "
                    "Automation interface. No synthetic score, curvature radius "
                    "or A-Class grade was generated."
                )

            continuity_summary: Optional[str] = None
            if blend_definition is not None:
                continuity_values = [
                    item.get("continuity")
                    for item in blend_definition.get("limits", [])
                    if item.get("continuity") is not None
                ]
                if continuity_values:
                    minimum_continuity = min(continuity_values)
                    continuity_summary = {
                        0: "G0",
                        1: "G1",
                        2: "G2",
                    }.get(minimum_continuity, f"Unknown({minimum_continuity})")

            _clear_selection(conn)

            return _success(
                {
                    "surface_name": requested_name,
                    "resolved_name": _object_name(
                        surface_obj,
                        requested_name,
                    ),
                    "geometrical_set": resolved_container,
                    "analysis_type_requested": requested_analysis,
                    "reference": reference_details,
                    "measurable": measurable_data,
                    "measurable_interface": measurable_interface,
                    "blend_definition_available": (
                        blend_definition is not None
                    ),
                    "blend_definition": blend_definition,
                    "blend_interface": blend_interface,
                    "continuity_summary": continuity_summary,
                    "quality_grade": "Not evaluated",
                    "quantitative_quality_score": None,
                    "zebra_stripe_smoothness": None,
                    "curvature_comb_uniformity": None,
                    "min_curvature_radius": None,
                    "max_curvature_radius": None,
                    "analysis_capabilities": {
                        "surface_area": measurable_data["area_mm2"] is not None,
                        "surface_area_source_unit": "m^2",
                        "surface_area_output_unit": "mm^2",
                        "surface_perimeter": (
                            measurable_data["perimeter_mm"] is not None
                        ),
                        "center_of_gravity": (
                            measurable_data["center_of_gravity_mm"] is not None
                        ),
                        "blend_definition": blend_definition is not None,
                        "quantitative_zebra": False,
                        "quantitative_curvature_comb": False,
                        "a_class_grading": False,
                    },
                    "feature_created": False,
                    "model_modified": False,
                },
                warnings,
            )
        except Exception as exc:
            _clear_selection(conn)
            return _error(
                str(exc),
                data={
                    "surface_name": str(surface_name),
                    "feature_created": False,
                    "model_modified": False,
                    "rollback_succeeded": True,
                },
                warnings=warnings,
            )

    names.append("analyze_surface_quality")

    return names


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
