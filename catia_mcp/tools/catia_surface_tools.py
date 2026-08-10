from __future__ import annotations

import math
from typing import Any, List, Optional

from catia_mcp.connection import CATIAError


IMPLEMENTATION_VERSION = "surface-tools-fixed-2026-07-30-v2"


def _success(data: Any, warnings: Optional[List[str]] = None) -> dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "ok": True,
        "status": "success_with_warnings" if warning_list else "success",
        "implementation_version": IMPLEMENTATION_VERSION,
        "data": data,
        "warnings": warning_list,
    }


def _error(
    message: str,
    *,
    data: Any = None,
    warnings: Optional[List[str]] = None,
    status: str = "error",
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "ok": False,
        "status": status,
        "implementation_version": IMPLEMENTATION_VERSION,
        "error": str(message),
        "warnings": list(warnings or []),
    }
    if data is not None:
        result["data"] = data
    return result


def _get_connection(ctx: Any) -> Any:
    """Support both the project context object and a direct connection object."""
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


def _integer_in_range(
    value: Any,
    name: str,
    allowed_values: set[int],
) -> int:
    if isinstance(value, bool):
        raise CATIAError(f"{name} must be an integer.")

    try:
        integer = int(value)
        numeric = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(f"{name} must be an integer.") from exc

    if not math.isfinite(numeric) or numeric != float(integer):
        raise CATIAError(f"{name} must be an integer.")

    if integer not in allowed_values:
        allowed = ", ".join(str(item) for item in sorted(allowed_values))
        raise CATIAError(f"{name} must be one of: {allowed}.")

    return integer


def _normalise_name(value: Any, default: str, parameter_name: str) -> str:
    text = str(value).strip()
    if text:
        return text

    fallback = str(default).strip()
    if fallback:
        return fallback

    raise CATIAError(f"{parameter_name} cannot be empty.")


def _normalise_name_list(
    values: Any,
    parameter_name: str,
    minimum_count: int,
) -> List[str]:
    if values is None:
        values = []

    if isinstance(values, (str, bytes)) or not isinstance(values, (list, tuple)):
        raise CATIAError(f"{parameter_name} must be a list of object names.")

    names: List[str] = []
    for index, value in enumerate(values, start=1):
        name = str(value).strip()
        if not name:
            raise CATIAError(
                f"{parameter_name}[{index}] cannot be empty."
            )
        names.append(name)

    if len(names) < minimum_count:
        raise CATIAError(
            f"{parameter_name} must contain at least "
            f"{minimum_count} object names."
        )

    duplicate_names = sorted(
        {name for name in names if names.count(name) > 1}
    )
    if duplicate_names:
        raise CATIAError(
            f"{parameter_name} contains duplicate names: "
            f"{', '.join(duplicate_names)}."
        )

    return names


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
    required_methods: List[str],
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
    required_methods: List[str],
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


def _get_shape_factory(
    part: Any,
    required_methods: List[str],
) -> tuple[Any, dict[str, Any]]:
    try:
        raw_factory = part.ShapeFactory
    except Exception as exc:
        raise CATIAError(
            f"Cannot access Part.ShapeFactory: {exc}"
        ) from exc

    return _resolve_com_interface(
        raw_factory,
        required_methods,
        "ShapeFactory",
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


def _set_object_name(obj: Any, requested_name: str) -> List[str]:
    warnings: List[str] = []
    try:
        obj.Name = requested_name
    except Exception as exc:
        warnings.append(
            f"Object was created but could not be renamed to "
            f"'{requested_name}': {exc}"
        )
    return warnings


def _object_name(obj: Any, fallback: str = "") -> str:
    try:
        name = str(obj.Name).strip()
        return name or fallback
    except Exception:
        return fallback


def _get_or_create_hybrid_body(
    part: Any,
    requested_name: str,
) -> tuple[Any, bool, dict[str, Any], List[str]]:
    body_name = _normalise_name(
        requested_name,
        "MCP_Surfaces",
        "geometrical_set",
    )
    warnings: List[str] = []

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


def _get_or_create_solid_body(
    part: Any,
) -> tuple[Any, bool, dict[str, Any], List[str]]:
    warnings: List[str] = []

    try:
        main_body = part.MainBody
        if main_body is not None:
            return (
                main_body,
                False,
                {
                    "name": _object_name(main_body, "PartBody"),
                    "created": False,
                    "source": "Part.MainBody",
                },
                warnings,
            )
    except Exception:
        pass

    try:
        bodies = part.Bodies
    except Exception as exc:
        raise CATIAError(f"Cannot access Part.Bodies: {exc}") from exc

    try:
        count = int(bodies.Count)
    except Exception:
        count = 0

    if count > 0:
        try:
            body = bodies.Item(1)
            return (
                body,
                False,
                {
                    "name": _object_name(body, "PartBody"),
                    "created": False,
                    "source": "Part.Bodies.Item(1)",
                },
                warnings,
            )
        except Exception as exc:
            raise CATIAError(
                f"Cannot access the existing solid body: {exc}"
            ) from exc

    try:
        body = bodies.Add()
    except Exception as exc:
        raise CATIAError(f"Cannot create a solid body: {exc}") from exc

    warnings.extend(_set_object_name(body, "PartBody"))

    return (
        body,
        True,
        {
            "name": _object_name(body, "PartBody"),
            "created": True,
            "source": "Part.Bodies.Add",
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
    object_name: str,
) -> tuple[Any, dict[str, Any]]:
    try:
        reference = part.CreateReferenceFromObject(obj)
    except Exception as exc:
        raise CATIAError(
            f"Cannot create a CATIA Reference from '{object_name}': {exc}"
        ) from exc

    if reference is None:
        raise CATIAError(
            f"CATIA returned an empty Reference for '{object_name}'."
        )

    return reference, {
        "requested_name": object_name,
        "resolved_name": _object_name(obj, object_name),
        "object": _describe_com_object(obj),
        "reference": _describe_com_object(reference),
    }


def _resolve_named_reference(
    part: Any,
    object_name: str,
) -> tuple[Any, dict[str, Any]]:
    obj = _find_named_object(part, object_name)
    return _create_reference(part, obj, object_name)


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
        try:
            document = _active_document(conn)
            document.Selection.Clear()
        except Exception:
            pass
        return False


def _update_feature(
    part: Any,
    feature: Any,
) -> tuple[bool, str, List[str]]:
    warnings: List[str] = []

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


def _refresh_display(conn: Any) -> List[str]:
    warnings: List[str] = []
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
) -> tuple[dict[str, Any], List[str]]:
    warnings: List[str] = []

    if feature is None:
        feature_rollback_succeeded = True
    elif feature_appended:
        feature_rollback_succeeded = _delete_object(conn, feature)
    else:
        # A non-appended HybridShape is not persistent in the specification tree.
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
                "The newly created target container could not be removed "
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
    warnings: Optional[List[str]] = None,
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


def _normalise_orientations(
    orientations: Optional[List[int]],
    section_count: int,
) -> List[int]:
    if orientations is None:
        return [1] * section_count

    if isinstance(orientations, (str, bytes)) or not isinstance(
        orientations,
        (list, tuple),
    ):
        raise CATIAError(
            "section_orientations must be a list containing 1 or -1."
        )

    if len(orientations) != section_count:
        raise CATIAError(
            "section_orientations must contain exactly one value for "
            "each section."
        )

    return [
        _integer_in_range(
            value,
            f"section_orientations[{index}]",
            {-1, 1},
        )
        for index, value in enumerate(orientations, start=1)
    ]


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = _get_connection(ctx)
    names: List[str] = []

    @mcp.tool()
    def catia_create_loft(
        section_names: List[str],
        guide_names: Optional[List[str]] = None,
        feature_name: str = "MCP_Loft",
        geometrical_set: str = "MCP_Surfaces",
        section_orientations: Optional[List[int]] = None,
        relimitation: int = 1,
        section_coupling: int = 1,
    ) -> dict[str, Any]:
        """Create a GSD multi-section Loft surface from named curve references.

        guide_names may be omitted, None, or an empty list for an unguided
        Loft. section_orientations uses 1 for the natural curve direction
        and -1 for the reversed direction.
        """

        loft = None
        feature_appended = False
        hybrid_body = None
        hybrid_body_created = False
        warnings: List[str] = []

        try:
            sections = _normalise_name_list(
                section_names,
                "section_names",
                minimum_count=2,
            )
            # None and [] both mean that the Loft has no guide curves.
            # Explicit empty arrays are common in JSON/MCP calls and must not
            # be rejected as an invalid one-element list.
            guides = (
                []
                if not guide_names
                else _normalise_name_list(
                    guide_names,
                    "guide_names",
                    minimum_count=1,
                )
            )
            orientations = _normalise_orientations(
                section_orientations,
                len(sections),
            )
            loft_name = _normalise_name(
                feature_name, "MCP_Loft", "feature_name"
            )
            body_name = _normalise_name(
                geometrical_set,
                "MCP_Surfaces",
                "geometrical_set",
            )
            relimitation_value = _integer_in_range(
                relimitation,
                "relimitation",
                {1, 2, 3, 4},
            )
            coupling_value = _integer_in_range(
                section_coupling,
                "section_coupling",
                {1, 2, 3, 4},
            )

            _ensure_connected(conn)
            part = _get_active_part(conn)

            factory, factory_details = _get_hybrid_shape_factory(
                part,
                ["AddNewLoft"],
            )
            (
                hybrid_body,
                hybrid_body_created,
                body_details,
                body_warnings,
            ) = _get_or_create_hybrid_body(part, body_name)
            warnings.extend(body_warnings)
            _set_in_work_object(part, hybrid_body)

            section_references: List[dict[str, Any]] = []
            resolved_section_refs: List[Any] = []
            for section_name in sections:
                section_ref, details = _resolve_named_reference(
                    part,
                    section_name,
                )
                resolved_section_refs.append(section_ref)
                section_references.append(details)

            guide_references: List[dict[str, Any]] = []
            resolved_guide_refs: List[Any] = []
            for guide_name in guides:
                guide_ref, details = _resolve_named_reference(
                    part,
                    guide_name,
                )
                resolved_guide_refs.append(guide_ref)
                guide_references.append(details)

            raw_loft = factory.AddNewLoft()
            loft_required_methods = ["AddSectionToLoft"]
            if resolved_guide_refs:
                loft_required_methods.append("AddGuide")

            loft, loft_interface = _resolve_com_interface(
                raw_loft,
                loft_required_methods,
                "HybridShapeLoft",
            )

            for section_ref, orientation in zip(
                resolved_section_refs,
                orientations,
            ):
                loft.AddSectionToLoft(
                    section_ref,
                    orientation,
                    None,
                )

            for guide_ref in resolved_guide_refs:
                loft.AddGuide(guide_ref)

            try:
                loft.Relimitation = relimitation_value
            except Exception as exc:
                raise CATIAError(
                    f"Cannot set Loft.Relimitation to "
                    f"{relimitation_value}: {exc}"
                ) from exc

            try:
                loft.SectionCoupling = coupling_value
            except Exception as exc:
                raise CATIAError(
                    f"Cannot set Loft.SectionCoupling to "
                    f"{coupling_value}: {exc}"
                ) from exc

            warnings.extend(_set_object_name(loft, loft_name))

            hybrid_body.AppendHybridShape(loft)
            feature_appended = True

            updated, update_strategy, update_warnings = _update_feature(
                part,
                loft,
            )
            warnings.extend(update_warnings)

            if not updated:
                return _operation_failure(
                    conn,
                    "Loft was created and appended, but CATIA could not "
                    "update the feature.",
                    loft,
                    feature_appended,
                    hybrid_body,
                    hybrid_body_created,
                    "HybridShapeLoft",
                    warnings,
                    {
                        "name": loft_name,
                        "sections": section_references,
                        "section_count": len(sections),
                        "section_orientations": orientations,
                        "guides": guide_references,
                        "guide_count": len(guides),
                        "relimitation": relimitation_value,
                        "section_coupling": coupling_value,
                        "geometrical_set": body_details,
                        "factory": factory_details,
                        "feature_interface": loft_interface,
                        "api_methods": [
                            "AddNewLoft",
                            "AddSectionToLoft",
                            *(["AddGuide"] if resolved_guide_refs else []),
                        ],
                        "update_strategy": update_strategy,
                    },
                )

            warnings.extend(_refresh_display(conn))

            return _success(
                {
                    "name": _object_name(loft, loft_name),
                    "surface": _object_name(loft, loft_name),
                    "type": "HybridShapeLoft",
                    "sections": section_references,
                    "section_count": len(sections),
                    "section_orientations": orientations,
                    "guides": guide_references,
                    "guide_count": len(guides),
                    "relimitation": relimitation_value,
                    "section_coupling": coupling_value,
                    "geometrical_set": body_details,
                    "factory": factory_details,
                    "feature_interface": loft_interface,
                    "api_methods": [
                        "AddNewLoft",
                        "AddSectionToLoft",
                        *(["AddGuide"] if resolved_guide_refs else []),
                    ],
                    "feature_created": True,
                    "feature_appended": True,
                    "feature_persisted": True,
                    "update_succeeded": True,
                    "update_strategy": update_strategy,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _operation_failure(
                conn,
                str(exc),
                loft,
                feature_appended,
                hybrid_body,
                hybrid_body_created,
                "HybridShapeLoft",
                warnings,
            )

    names.append("catia_create_loft")

    @mcp.tool()
    def catia_close_surface(
        surface_name: str,
        feature_name: str = "MCP_CloseSurface",
    ) -> dict[str, Any]:
        """Close a watertight surface skin and add the resulting solid to PartBody."""

        close_feature = None
        feature_appended = False
        target_body = None
        target_body_created = False
        warnings: List[str] = []

        try:
            source_name = _normalise_name(
                surface_name,
                "",
                "surface_name",
            )
            resolved_feature_name = _normalise_name(
                feature_name,
                "MCP_CloseSurface",
                "feature_name",
            )

            _ensure_connected(conn)
            part = _get_active_part(conn)

            source_reference, source_details = _resolve_named_reference(
                part,
                source_name,
            )

            factory, factory_details = _get_shape_factory(
                part,
                ["AddNewCloseSurface"],
            )
            (
                target_body,
                target_body_created,
                body_details,
                body_warnings,
            ) = _get_or_create_solid_body(part)
            warnings.extend(body_warnings)
            _set_in_work_object(part, target_body)

            close_feature = factory.AddNewCloseSurface(source_reference)
            # ShapeFactory features are inserted into the current Body by CATIA.
            feature_appended = True
            warnings.extend(
                _set_object_name(close_feature, resolved_feature_name)
            )

            updated, update_strategy, update_warnings = _update_feature(
                part,
                close_feature,
            )
            warnings.extend(update_warnings)

            if not updated:
                return _operation_failure(
                    conn,
                    "CloseSurface was created, but CATIA could not update it. "
                    "Verify that the source surface is a watertight closed skin.",
                    close_feature,
                    feature_appended,
                    target_body,
                    target_body_created,
                    "CloseSurface",
                    warnings,
                    {
                        "name": resolved_feature_name,
                        "source_surface": source_details,
                        "target_body": body_details,
                        "factory": factory_details,
                        "api_method": "AddNewCloseSurface",
                        "update_strategy": update_strategy,
                    },
                )

            warnings.extend(_refresh_display(conn))

            return _success(
                {
                    "name": _object_name(close_feature, resolved_feature_name),
                    "solid_body": _object_name(
                        close_feature,
                        resolved_feature_name,
                    ),
                    "type": "CloseSurface",
                    "source_surface": source_details,
                    "target_body": body_details,
                    "factory": factory_details,
                    "api_method": "AddNewCloseSurface",
                    "feature_created": True,
                    "feature_appended": True,
                    "feature_persisted": True,
                    "update_succeeded": True,
                    "update_strategy": update_strategy,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _operation_failure(
                conn,
                str(exc),
                close_feature,
                feature_appended,
                target_body,
                target_body_created,
                "CloseSurface",
                warnings,
            )

    names.append("catia_close_surface")

    @mcp.tool()
    def catia_join_surfaces(
        surface_names: List[str],
        feature_name: str = "MCP_Join",
        geometrical_set: str = "MCP_Surfaces",
        deviation: float = 0.001,
        connex_check: bool = False,
        manifold_check: bool = False,
        simplify: bool = False,
    ) -> dict[str, Any]:
        """Join two or more named curves or surfaces into one GSD Join feature."""

        join = None
        feature_appended = False
        hybrid_body = None
        hybrid_body_created = False
        warnings: List[str] = []

        try:
            surfaces = _normalise_name_list(
                surface_names,
                "surface_names",
                minimum_count=2,
            )
            join_name = _normalise_name(
                feature_name, "MCP_Join", "feature_name"
            )
            body_name = _normalise_name(
                geometrical_set,
                "MCP_Surfaces",
                "geometrical_set",
            )
            deviation_value = _positive_number(
                deviation,
                "deviation",
            )

            _ensure_connected(conn)
            part = _get_active_part(conn)

            factory, factory_details = _get_hybrid_shape_factory(
                part,
                ["AddNewJoin"],
            )
            (
                hybrid_body,
                hybrid_body_created,
                body_details,
                body_warnings,
            ) = _get_or_create_hybrid_body(part, body_name)
            warnings.extend(body_warnings)
            _set_in_work_object(part, hybrid_body)

            surface_references: List[dict[str, Any]] = []
            resolved_refs: List[Any] = []
            for surface_name in surfaces:
                surface_ref, details = _resolve_named_reference(
                    part,
                    surface_name,
                )
                resolved_refs.append(surface_ref)
                surface_references.append(details)

            raw_join = factory.AddNewJoin(
                resolved_refs[0],
                resolved_refs[1],
            )
            join, join_interface = _resolve_com_interface(
                raw_join,
                [
                    "AddElement",
                    "SetDeviation",
                    "SetConnex",
                    "SetManifold",
                    "SetSimplify",
                ],
                "HybridShapeAssemble",
            )

            for surface_ref in resolved_refs[2:]:
                join.AddElement(surface_ref)

            join.SetDeviation(deviation_value)
            join.SetConnex(bool(connex_check))
            join.SetManifold(bool(manifold_check))
            join.SetSimplify(bool(simplify))

            warnings.extend(_set_object_name(join, join_name))

            hybrid_body.AppendHybridShape(join)
            feature_appended = True

            updated, update_strategy, update_warnings = _update_feature(
                part,
                join,
            )
            warnings.extend(update_warnings)

            if not updated:
                return _operation_failure(
                    conn,
                    "Join was created and appended, but CATIA could not "
                    "update the feature.",
                    join,
                    feature_appended,
                    hybrid_body,
                    hybrid_body_created,
                    "HybridShapeAssemble",
                    warnings,
                    {
                        "name": join_name,
                        "elements": surface_references,
                        "element_count": len(surfaces),
                        "deviation": deviation_value,
                        "merging_distance": deviation_value,
                        "connex_check": bool(connex_check),
                        "manifold_check": bool(manifold_check),
                        "simplify": bool(simplify),
                        "geometrical_set": body_details,
                        "factory": factory_details,
                        "feature_interface": join_interface,
                        "api_methods": [
                            "AddNewJoin",
                            "AddElement",
                            "SetDeviation",
                        ],
                        "update_strategy": update_strategy,
                    },
                )

            actual_deviation: Optional[float] = None
            try:
                actual_deviation = float(join.GetDeviation())
            except Exception as exc:
                warnings.append(
                    f"Join succeeded, but GetDeviation could not verify the "
                    f"requested tolerance: {exc}"
                )

            warnings.extend(_refresh_display(conn))

            return _success(
                {
                    "name": _object_name(join, join_name),
                    "joined_surface": _object_name(join, join_name),
                    "type": "HybridShapeAssemble",
                    "elements": surface_references,
                    "element_count": len(surfaces),
                    "deviation": deviation_value,
                    "merging_distance": deviation_value,
                    "verified_deviation": actual_deviation,
                    "verified_merging_distance": actual_deviation,
                    "connex_check": bool(connex_check),
                    "manifold_check": bool(manifold_check),
                    "simplify": bool(simplify),
                    "geometrical_set": body_details,
                    "factory": factory_details,
                    "feature_interface": join_interface,
                    "api_methods": [
                        "AddNewJoin",
                        "AddElement",
                        "SetDeviation",
                    ],
                    "feature_created": True,
                    "feature_appended": True,
                    "feature_persisted": True,
                    "update_succeeded": True,
                    "update_strategy": update_strategy,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _operation_failure(
                conn,
                str(exc),
                join,
                feature_appended,
                hybrid_body,
                hybrid_body_created,
                "HybridShapeAssemble",
                warnings,
            )

    names.append("catia_join_surfaces")

    return names


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
