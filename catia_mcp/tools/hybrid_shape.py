from __future__ import annotations

import math
from typing import Any

from catia_mcp.connection import CATIAError


def _success(data: Any, warnings: list[str] | None = None) -> dict[str, Any]:
    warning_list = list(warnings or [])
    return {
        "ok": True,
        "status": "success_with_warnings" if warning_list else "success",
        "data": data,
        "warnings": warning_list,
    }


def _error(
    message: str,
    *,
    data: Any | None = None,
    warnings: list[str] | None = None,
    status: str = "error",
) -> dict[str, Any]:
    result: dict[str, Any] = {
        "ok": False,
        "status": status,
        "error": str(message),
        "warnings": list(warnings or []),
    }
    if data is not None:
        result["data"] = data
    return result


def _finite_number(value: Any, name: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(f"{name} must be a number.") from exc

    if not math.isfinite(number):
        raise CATIAError(f"{name} must be finite.")

    return number


def _normalise_name(value: Any, default: str, parameter_name: str) -> str:
    text = str(value).strip()
    if text:
        return text

    fallback = str(default).strip()
    if fallback:
        return fallback

    raise CATIAError(f"{parameter_name} cannot be empty.")


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
    """Resolve a typed CATIA COM interface from a generic pywin32 wrapper."""

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

    details["dispatch_used"] = True
    details["resolved"] = _describe_com_object(resolved)
    details["required_methods_available"] = all(
        _has_callable(resolved, method) for method in required_methods
    )

    if not details["required_methods_available"]:
        missing = [
            method
            for method in required_methods
            if not _has_callable(resolved, method)
        ]
        raise CATIAError(
            f"{interface_name} dispatch completed, but the resolved COM "
            f"object does not expose: {', '.join(missing)}."
        )

    return resolved, details


def _active_document(conn: Any) -> Any:
    try:
        return conn.app.ActiveDocument
    except Exception as exc:
        raise CATIAError("Cannot access the active CATIA document.") from exc


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
            _active_document(conn).Selection.Clear()
        except Exception:
            pass
        return False


def _set_name(obj: Any, requested_name: str) -> list[str]:
    warnings: list[str] = []
    try:
        obj.Name = requested_name
    except Exception as exc:
        warnings.append(
            f"Object was created but could not be renamed to "
            f"'{requested_name}': {exc}"
        )
    return warnings


def _text_attribute(obj: Any, attribute_name: str) -> str:
    try:
        return str(getattr(obj, attribute_name)).strip()
    except Exception:
        return ""


def _refresh_display(conn: Any) -> list[str]:
    warnings: list[str] = []
    try:
        conn.refresh_display()
    except Exception as exc:
        warnings.append(f"Feature created, but display refresh failed: {exc}")
    return warnings


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


def _get_hybrid_shape_factory(
    part: Any,
    required_method: str,
) -> tuple[Any, dict[str, Any]]:
    try:
        raw_factory = part.HybridShapeFactory
    except Exception as exc:
        raise CATIAError(
            f"Cannot access Part.HybridShapeFactory: {exc}"
        ) from exc

    return _resolve_com_interface(
        raw_factory,
        [required_method],
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


def _get_or_create_hybrid_body(
    part: Any,
    requested_name: str,
) -> tuple[Any, bool, dict[str, Any]]:
    body_name = _normalise_name(
        requested_name,
        "MCP_Construction",
        "geometrical_set",
    )

    collection, collection_details = _get_hybrid_bodies_collection(part)

    raw_body = None
    body_found = False

    try:
        raw_body = collection.Item(body_name)
        body_found = True
    except Exception:
        body_found = False

    if body_found:
        body, body_interface_details = _resolve_com_interface(
            raw_body,
            ["AppendHybridShape"],
            "HybridBody",
        )
        return body, False, {
            "name": _text_attribute(body, "Name") or body_name,
            "created": False,
            "collection": collection_details,
            "body_interface": body_interface_details,
        }

    try:
        raw_body = collection.Add()
    except Exception as exc:
        raise CATIAError(
            f"Cannot create geometrical set '{body_name}': {exc}"
        ) from exc

    body, body_interface_details = _resolve_com_interface(
        raw_body,
        ["AppendHybridShape"],
        "HybridBody",
    )

    warnings = _set_name(body, body_name)
    actual_name = _text_attribute(body, "Name") or body_name

    return body, True, {
        "name": actual_name,
        "requested_name": body_name,
        "created": True,
        "rename_warnings": warnings,
        "collection": collection_details,
        "body_interface": body_interface_details,
    }


def _set_in_work_object(part: Any, obj: Any) -> None:
    try:
        part.InWorkObject = obj
    except Exception as exc:
        raise CATIAError(
            f"Cannot set the active geometrical set as Part.InWorkObject: {exc}"
        ) from exc


def _normalise_base_plane(base_plane: Any) -> tuple[str, str]:
    key = str(base_plane).strip().lower().replace(" ", "")

    aliases = {
        "xy": ("xy", "PlaneXY"),
        "planexy": ("xy", "PlaneXY"),
        "yz": ("yz", "PlaneYZ"),
        "planeyz": ("yz", "PlaneYZ"),
        "zx": ("zx", "PlaneZX"),
        "xz": ("zx", "PlaneZX"),
        "planezx": ("zx", "PlaneZX"),
        "planexz": ("zx", "PlaneZX"),
    }

    if key not in aliases:
        raise CATIAError(
            "base_plane must be one of: xy, yz, zx, or xz."
        )

    return aliases[key]


def _get_origin_plane_reference(
    part: Any,
    base_plane: Any,
) -> tuple[Any, dict[str, Any]]:
    normalised, attribute_name = _normalise_base_plane(base_plane)

    try:
        origin_elements = part.OriginElements
        origin_plane = getattr(origin_elements, attribute_name)
    except Exception as exc:
        raise CATIAError(
            f"Cannot access origin plane {attribute_name}: {exc}"
        ) from exc

    try:
        reference = part.CreateReferenceFromObject(origin_plane)
    except Exception as exc:
        raise CATIAError(
            f"Cannot create a CATIA Reference from {attribute_name}: {exc}"
        ) from exc

    return reference, {
        "requested": str(base_plane),
        "normalised": normalised,
        "origin_attribute": attribute_name,
        "origin_plane": _describe_com_object(origin_plane),
        "reference": _describe_com_object(reference),
    }


def _hybrid_shape_failure(
    conn: Any,
    feature: Any,
    *,
    feature_appended: bool,
    hybrid_body: Any,
    hybrid_body_created: bool,
    message: str,
    feature_type: str,
    warnings: list[str] | None = None,
    extra_data: dict[str, Any] | None = None,
) -> dict[str, Any]:
    warning_list = list(warnings or [])
    feature_created = feature is not None

    if feature_appended:
        feature_rollback_succeeded = _delete_object(conn, feature)
    else:
        # A HybridShape not appended to a HybridBody is not persisted in the tree.
        feature_rollback_succeeded = True

    feature_persisted = bool(
        feature_appended and not feature_rollback_succeeded
    )

    body_rollback_attempted = bool(
        hybrid_body_created and not feature_persisted
    )
    body_rollback_succeeded: bool | None = None

    if body_rollback_attempted:
        body_rollback_succeeded = _delete_object(conn, hybrid_body)
        if not body_rollback_succeeded:
            warning_list.append(
                "The newly created geometrical set could not be removed "
                "after the feature failure."
            )

    data: dict[str, Any] = {
        "type": feature_type,
        "feature_created": feature_created,
        "feature_appended": feature_appended,
        "feature_persisted": feature_persisted,
        "update_succeeded": False,
        "feature_rollback_succeeded": feature_rollback_succeeded,
        "geometrical_set_created": hybrid_body_created,
        "geometrical_set_rollback_attempted": body_rollback_attempted,
        "geometrical_set_rollback_succeeded": body_rollback_succeeded,
        "rollback_succeeded": bool(
            feature_rollback_succeeded
            and (
                body_rollback_succeeded is not False
            )
        ),
    }

    if extra_data:
        data.update(extra_data)

    if feature_persisted:
        warning_list.append(
            "The operation failed and the HybridShape could not be removed; "
            "the active CATPart may have been modified."
        )
        return _error(
            message,
            data=data,
            warnings=warning_list,
            status="partial_success",
        )

    return _error(
        message,
        data=data,
        warnings=warning_list,
    )


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = ctx.conn
    names: list[str] = []

    @mcp.tool()
    def catia_create_point(
        x: float,
        y: float,
        z: float,
        name: str = "MCP_Point",
        geometrical_set: str = "MCP_Construction",
    ) -> dict[str, Any]:
        """Create a cartesian 3D point in a top-level Geometrical Set."""

        feature = None
        feature_appended = False
        hybrid_body = None
        hybrid_body_created = False
        warnings: list[str] = []

        try:
            x_value = _finite_number(x, "x")
            y_value = _finite_number(y, "y")
            z_value = _finite_number(z, "z")
            feature_name = _normalise_name(name, "MCP_Point", "name")
            body_name = _normalise_name(
                geometrical_set,
                "MCP_Construction",
                "geometrical_set",
            )

            conn.ensure_connected()
            part = conn.get_active_part()

            factory, factory_details = _get_hybrid_shape_factory(
                part,
                "AddNewPointCoord",
            )
            (
                hybrid_body,
                hybrid_body_created,
                body_details,
            ) = _get_or_create_hybrid_body(part, body_name)

            warnings.extend(body_details.get("rename_warnings", []))
            _set_in_work_object(part, hybrid_body)

            feature = factory.AddNewPointCoord(
                x_value,
                y_value,
                z_value,
            )

            hybrid_body.AppendHybridShape(feature)
            feature_appended = True

            warnings.extend(_set_name(feature, feature_name))

            updated, update_strategy, update_warnings = _update_feature(
                part,
                feature,
            )
            warnings.extend(update_warnings)

            if not updated:
                return _hybrid_shape_failure(
                    conn,
                    feature,
                    feature_appended=feature_appended,
                    hybrid_body=hybrid_body,
                    hybrid_body_created=hybrid_body_created,
                    message=(
                        "Point was appended to the geometrical set, but "
                        "CATIA could not update it."
                    ),
                    feature_type="HybridShapePointCoord",
                    warnings=warnings,
                    extra_data={
                        "name": feature_name,
                        "coordinates": {
                            "x": x_value,
                            "y": y_value,
                            "z": z_value,
                        },
                        "geometrical_set": body_details,
                        "factory": factory_details,
                        "api_method": "AddNewPointCoord",
                        "update_strategy": update_strategy,
                    },
                )

            warnings.extend(_refresh_display(conn))

            return _success(
                {
                    "name": _text_attribute(feature, "Name") or feature_name,
                    "type": "HybridShapePointCoord",
                    "x": x_value,
                    "y": y_value,
                    "z": z_value,
                    "coordinates": {
                        "x": x_value,
                        "y": y_value,
                        "z": z_value,
                    },
                    "geometrical_set": body_details,
                    "api_method": "AddNewPointCoord",
                    "factory": factory_details,
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
            return _hybrid_shape_failure(
                conn,
                feature,
                feature_appended=feature_appended,
                hybrid_body=hybrid_body,
                hybrid_body_created=hybrid_body_created,
                message=str(exc),
                feature_type="HybridShapePointCoord",
                warnings=warnings,
            )

    names.append("catia_create_point")

    @mcp.tool()
    def catia_create_offset_plane(
        base_plane: str = "xy",
        offset: float = 0.0,
        name: str = "",
        geometrical_set: str = "MCP_Construction",
    ) -> dict[str, Any]:
        """Create an actual offset plane from an origin XY/YZ/ZX plane.

        A zero offset still creates a real HybridShapePlaneOffset feature;
        it does not merely return the origin-plane reference.
        """

        feature = None
        feature_appended = False
        hybrid_body = None
        hybrid_body_created = False
        warnings: list[str] = []

        try:
            offset_value = _finite_number(offset, "offset")
            normalised_plane, _ = _normalise_base_plane(base_plane)
            default_name = (
                f"MCP_{normalised_plane.upper()}_Offset_"
                f"{offset_value:g}"
            )
            feature_name = _normalise_name(name, default_name, "name")
            body_name = _normalise_name(
                geometrical_set,
                "MCP_Construction",
                "geometrical_set",
            )

            conn.ensure_connected()
            part = conn.get_active_part()

            base_reference, base_details = _get_origin_plane_reference(
                part,
                normalised_plane,
            )
            factory, factory_details = _get_hybrid_shape_factory(
                part,
                "AddNewPlaneOffset",
            )
            (
                hybrid_body,
                hybrid_body_created,
                body_details,
            ) = _get_or_create_hybrid_body(part, body_name)

            warnings.extend(body_details.get("rename_warnings", []))
            _set_in_work_object(part, hybrid_body)

            feature = factory.AddNewPlaneOffset(
                base_reference,
                offset_value,
                False,
            )

            hybrid_body.AppendHybridShape(feature)
            feature_appended = True

            warnings.extend(_set_name(feature, feature_name))

            updated, update_strategy, update_warnings = _update_feature(
                part,
                feature,
            )
            warnings.extend(update_warnings)

            if not updated:
                return _hybrid_shape_failure(
                    conn,
                    feature,
                    feature_appended=feature_appended,
                    hybrid_body=hybrid_body,
                    hybrid_body_created=hybrid_body_created,
                    message=(
                        "Offset plane was appended to the geometrical set, "
                        "but CATIA could not update it."
                    ),
                    feature_type="HybridShapePlaneOffset",
                    warnings=warnings,
                    extra_data={
                        "name": feature_name,
                        "base_plane": normalised_plane,
                        "offset": offset_value,
                        "orientation_reversed": False,
                        "base_reference": base_details,
                        "geometrical_set": body_details,
                        "factory": factory_details,
                        "api_method": "AddNewPlaneOffset",
                        "update_strategy": update_strategy,
                    },
                )

            warnings.extend(_refresh_display(conn))

            return _success(
                {
                    "name": _text_attribute(feature, "Name") or feature_name,
                    "type": "HybridShapePlaneOffset",
                    "base_plane": normalised_plane,
                    "offset": offset_value,
                    "orientation_reversed": False,
                    "base_reference": base_details,
                    "geometrical_set": body_details,
                    "api_method": "AddNewPlaneOffset",
                    "factory": factory_details,
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
            return _hybrid_shape_failure(
                conn,
                feature,
                feature_appended=feature_appended,
                hybrid_body=hybrid_body,
                hybrid_body_created=hybrid_body_created,
                message=str(exc),
                feature_type="HybridShapePlaneOffset",
                warnings=warnings,
            )

    names.append("catia_create_offset_plane")

    return names


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
