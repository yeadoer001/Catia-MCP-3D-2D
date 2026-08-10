from __future__ import annotations

import math
from typing import Any, Optional

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


def _finite_positive(value: Any, name: str) -> float:
    number = _finite_number(value, name)
    if number <= 0.0:
        raise CATIAError(f"{name} must be greater than 0.")
    return number


def _positive_integer(value: Any, name: str, minimum: int = 1) -> int:
    if isinstance(value, bool):
        raise CATIAError(f"{name} must be an integer.")

    try:
        integer = int(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(f"{name} must be an integer.") from exc

    try:
        numeric = float(value)
    except (TypeError, ValueError):
        numeric = float(integer)

    if not math.isfinite(numeric) or numeric != float(integer):
        raise CATIAError(f"{name} must be an integer.")

    if integer < minimum:
        raise CATIAError(f"{name} must be at least {minimum}.")

    return integer


def _normalise_name(value: Any, default: str) -> str:
    text = str(value).strip()
    return text or default


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


def _get_shape_factory(
    part: Any,
    required_methods: list[str],
) -> tuple[Any, str, dict[str, Any]]:
    """Resolve ShapeFactory and select the first available method."""

    try:
        raw_factory = part.ShapeFactory
    except Exception as exc:
        raise CATIAError(f"Cannot access Part.ShapeFactory: {exc}") from exc

    details: dict[str, Any] = {
        "required_methods": list(required_methods),
        "raw": _describe_com_object(raw_factory),
        "dispatch_used": False,
    }

    for method_name in required_methods:
        if _has_callable(raw_factory, method_name):
            details["selected_method"] = method_name
            details["resolved"] = _describe_com_object(raw_factory)
            details["required_method_available"] = True
            return raw_factory, method_name, details

    try:
        shape_factory = _dispatch_com(raw_factory)
    except Exception as exc:
        details["dispatch_error"] = str(exc)
        raise CATIAError(
            "Part.ShapeFactory was returned as a generic COM Factory and "
            f"could not be dynamically dispatched: {exc}"
        ) from exc

    details["dispatch_used"] = True
    details["resolved"] = _describe_com_object(shape_factory)

    for method_name in required_methods:
        if _has_callable(shape_factory, method_name):
            details["selected_method"] = method_name
            details["required_method_available"] = True
            return shape_factory, method_name, details

    details["required_method_available"] = False
    raise CATIAError(
        "ShapeFactory does not expose any required method: "
        + ", ".join(required_methods)
    )


def _active_document(conn: Any) -> Any:
    try:
        return conn.app.ActiveDocument
    except Exception as exc:
        raise CATIAError("Cannot access the active CATIA document.") from exc


def _update_feature(part: Any, feature: Any) -> tuple[bool, str, list[str]]:
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


def _delete_object(conn: Any, obj: Any) -> bool:
    if obj is None:
        return True

    try:
        selection = _active_document(conn).Selection
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


def _feature_failure(
    conn: Any,
    feature: Any,
    message: str,
    *,
    feature_type: str,
    warnings: list[str] | None = None,
    extra_data: dict[str, Any] | None = None,
) -> dict[str, Any]:
    warning_list = list(warnings or [])
    feature_created = feature is not None
    rollback_succeeded = _delete_object(conn, feature) if feature_created else True
    feature_persisted = feature_created and not rollback_succeeded

    data: dict[str, Any] = {
        "type": feature_type,
        "feature_created": feature_created,
        "feature_persisted": feature_persisted,
        "update_succeeded": False,
        "rollback_succeeded": rollback_succeeded,
    }
    if extra_data:
        data.update(extra_data)

    if feature_persisted:
        warning_list.append(
            "The operation failed and rollback also failed; "
            "the active CATPart may have been modified."
        )
        return _error(
            message,
            data=data,
            warnings=warning_list,
            status="partial_success",
        )

    return _error(message, data=data, warnings=warning_list)


def _set_feature_name(feature: Any, requested_name: str) -> list[str]:
    warnings: list[str] = []
    name = str(requested_name).strip()
    if not name:
        return warnings

    try:
        feature.Name = name
    except Exception as exc:
        warnings.append(
            f"Feature was created but could not be renamed to '{name}': {exc}"
        )

    return warnings


def _refresh_display(conn: Any) -> list[str]:
    warnings: list[str] = []
    try:
        conn.refresh_display()
    except Exception as exc:
        warnings.append(f"Feature created, but display refresh failed: {exc}")
    return warnings


def _selected_item(selection: Any, index: int) -> Any:
    try:
        return selection.Item2(index)
    except Exception:
        return selection.Item(index)


def _text_attribute(value: Any, attribute_name: str) -> str:
    try:
        text = str(getattr(value, attribute_name)).strip()
        return text
    except Exception:
        return ""


def _topology_record(item: Any, index: int, kind: str) -> dict[str, Any]:
    try:
        value = item.Value
    except Exception:
        value = item

    try:
        reference = item.Reference
    except Exception:
        reference = value

    try:
        parent = value.Parent
    except Exception:
        parent = None

    value_name = _text_attribute(value, "Name")
    reference_name = (
        _text_attribute(reference, "DisplayName")
        or _text_attribute(reference, "Name")
    )
    parent_name = _text_attribute(parent, "Name")

    return {
        "index": index,
        "selector": f"{kind}:{index}",
        "name": value_name,
        "reference_name": reference_name,
        "parent_name": parent_name,
        "python_type": type(value).__name__,
        "python_module": type(value).__module__,
        "_value": value,
    }


def _selection_count(selection: Any) -> int:
    """Read Selection count without eagerly evaluating an unavailable property."""

    try:
        return int(selection.Count2)
    except Exception:
        return int(selection.Count)


def _topology_search_queries(kind: str) -> list[str]:
    """Return locale-independent topology query candidates.

    CATIA's Search syntax separates the criterion and scope with a comma.
    The standard Topology.Edge/Face form is attempted first. Some CATIA
    installations expose the equivalent CGM type name, so that form is kept
    as a compatibility fallback.
    """

    kind_key = str(kind).strip().lower()
    if kind_key == "edge":
        return [
            "Topology.Edge,all",
            "Topology.CGMEdge,all",
        ]
    if kind_key == "face":
        return [
            "Topology.Face,all",
            "Topology.CGMFace,all",
        ]
    raise CATIAError("kind must be either 'edge' or 'face'.")


def _search_topology(
    document: Any,
    kind: str,
    *,
    feature_name: str = "",
) -> tuple[list[dict[str, Any]], str, list[str]]:
    """Search edges/faces and report which CATIA query actually succeeded."""

    kind_key = str(kind).strip().lower()
    queries = _topology_search_queries(kind_key)
    selection = document.Selection
    feature_filter = str(feature_name).strip().casefold()
    attempt_errors: list[str] = []

    try:
        for query in queries:
            try:
                selection.Clear()
                selection.Search(query)

                records: list[dict[str, Any]] = []
                count = _selection_count(selection)

                for index in range(1, count + 1):
                    record = _topology_record(
                        _selected_item(selection, index),
                        index,
                        kind_key,
                    )

                    if feature_filter:
                        searchable = " ".join(
                            [
                                record["parent_name"],
                                record["name"],
                                record["reference_name"],
                            ]
                        ).casefold()
                        if feature_filter not in searchable:
                            continue

                    records.append(record)

                return records, query, attempt_errors

            except Exception as exc:
                attempt_errors.append(f"{query}: {exc}")

        raise CATIAError(
            "CATIA topology search failed for all supported queries. "
            + " | ".join(attempt_errors)
        )
    finally:
        try:
            selection.Clear()
        except Exception:
            pass


def _public_topology_record(record: dict[str, Any]) -> dict[str, Any]:
    return {
        key: value
        for key, value in record.items()
        if not key.startswith("_")
    }


def _parse_selector_index(selector: Any, kind: str) -> int | None:
    if isinstance(selector, bool):
        return None

    if isinstance(selector, int):
        return selector

    text = str(selector).strip()
    if not text:
        return None

    lowered = text.casefold()
    prefix = f"{kind}:"
    if lowered.startswith(prefix):
        suffix = text[len(prefix):].strip()
        try:
            return int(suffix)
        except ValueError:
            return None

    if text.isdigit():
        return int(text)

    return None


def _resolve_topology_selector(
    part: Any,
    document: Any,
    selector: Any,
    kind: str,
) -> tuple[Any, dict[str, Any]]:
    kind_key = str(kind).strip().lower()
    selector_text = str(selector).strip()
    if not selector_text:
        raise CATIAError(f"{kind_key} selector cannot be empty.")

    records, search_query, search_attempt_errors = _search_topology(
        document,
        kind_key,
    )
    selector_index = _parse_selector_index(selector, kind_key)

    if selector_index is not None:
        if selector_index < 1:
            raise CATIAError(f"{kind_key} index must be at least 1.")

        for record in records:
            if int(record["index"]) == selector_index:
                details = _public_topology_record(record)
                details["search_query"] = search_query
                details["search_attempt_errors"] = search_attempt_errors
                return record["_value"], details

        raise CATIAError(
            f"{kind_key} selector '{selector_text}' is outside the current "
            f"topology range (1..{len(records)}). Run catia_list_topology "
            "immediately before the advanced feature operation."
        )

    wanted = selector_text.casefold()
    matches: list[dict[str, Any]] = []

    for record in records:
        candidates = {
            str(record["name"]).casefold(),
            str(record["reference_name"]).casefold(),
            str(record["selector"]).casefold(),
        }
        if wanted in candidates:
            matches.append(record)

    if len(matches) == 1:
        record = matches[0]
        details = _public_topology_record(record)
        details["search_query"] = search_query
        details["search_attempt_errors"] = search_attempt_errors
        return record["_value"], details

    if len(matches) > 1:
        raise CATIAError(
            f"{kind_key} selector '{selector_text}' is ambiguous. "
            "Use the indexed selector returned by catia_list_topology."
        )

    # Named tree objects are retained as a final compatibility fallback.
    try:
        found = part.FindObjectByName(selector_text)
        if found is not None:
            return found, {
                "selector": selector_text,
                "resolution": "Part.FindObjectByName",
                "name": _text_attribute(found, "Name"),
                "python_type": type(found).__name__,
                "python_module": type(found).__module__,
            }
    except Exception:
        pass

    raise CATIAError(
        f"Cannot resolve {kind_key} selector '{selector_text}'. "
        "Call catia_list_topology and use an indexed selector such as "
        f"'{kind_key}:1'."
    )


def _find_shape(body: Any, feature_name: str) -> Any:
    name = str(feature_name).strip()
    if not name:
        raise CATIAError("feature_name cannot be empty.")

    shapes = body.Shapes

    try:
        return shapes.Item(name)
    except Exception as direct_error:
        try:
            count = int(shapes.Count)
            for index in range(1, count + 1):
                shape = shapes.Item(index)
                candidate_name = _text_attribute(shape, "Name")
                if candidate_name.casefold() == name.casefold():
                    return shape
        except Exception:
            pass

        raise CATIAError(f"Feature not found in active PartBody: {name}") from direct_error


def _create_reference(part: Any, value: Any) -> Any:
    # Selection topology values can already be Boundary/Reference objects.
    try:
        return part.CreateReferenceFromObject(value)
    except Exception:
        return value


def _resolve_pattern_axis(
    part: Any,
    document: Any,
    axis_name: Optional[str],
) -> tuple[Any, dict[str, Any]]:
    requested = str(axis_name or "").strip()
    key = requested.casefold()

    if not requested or key in {"z", "z_axis", "axis_z", "xy", "plane_xy"}:
        axis_object = part.OriginElements.PlaneXY
        return _create_reference(part, axis_object), {
            "requested": requested,
            "strategy": "origin_plane_normal",
            "resolved_axis": "Z",
            "support": "PlaneXY",
        }

    if key in {"x", "x_axis", "axis_x", "yz", "plane_yz"}:
        axis_object = part.OriginElements.PlaneYZ
        return _create_reference(part, axis_object), {
            "requested": requested,
            "strategy": "origin_plane_normal",
            "resolved_axis": "X",
            "support": "PlaneYZ",
        }

    if key in {"y", "y_axis", "axis_y", "zx", "plane_zx", "xz", "plane_xz"}:
        axis_object = part.OriginElements.PlaneZX
        return _create_reference(part, axis_object), {
            "requested": requested,
            "strategy": "origin_plane_normal",
            "resolved_axis": "Y",
            "support": "PlaneZX",
        }

    if key.startswith("edge:") or requested.isdigit():
        edge, details = _resolve_topology_selector(
            part,
            document,
            requested,
            "edge",
        )
        return _create_reference(part, edge), {
            "requested": requested,
            "strategy": "topology_edge",
            "resolved": details,
        }

    try:
        axis_object = part.FindObjectByName(requested)
    except Exception as exc:
        raise CATIAError(
            f"Cannot find rotation axis object: {requested}"
        ) from exc

    return _create_reference(part, axis_object), {
        "requested": requested,
        "strategy": "named_object",
        "resolved_name": _text_attribute(axis_object, "Name"),
        "python_type": type(axis_object).__name__,
    }


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = ctx.conn
    names: list[str] = []

    @mcp.tool()
    def catia_list_topology(
        kind: str,
        feature_name: str = "",
        limit: int = 200,
    ) -> dict[str, Any]:
        """List usable edge or face selectors in the active CATPart.

        kind:
        - "edge": return selectors such as edge:1
        - "face": return selectors such as face:1

        Call this tool immediately before catia_add_fillet or catia_shell,
        because CATIA topological Boundary objects may become invalid after
        Part.Update.
        """
        try:
            conn.ensure_connected()
            document = _active_document(conn)
            part = conn.get_active_part()

            kind_key = str(kind).strip().lower()
            limit_value = _positive_integer(limit, "limit", 1)

            (
                records,
                search_query,
                search_attempt_errors,
            ) = _search_topology(
                document,
                kind_key,
                feature_name=feature_name,
            )
            public_records = [
                _public_topology_record(record)
                for record in records[:limit_value]
            ]

            return _success(
                {
                    "kind": kind_key,
                    "feature_filter": str(feature_name).strip(),
                    "count": len(records),
                    "returned_count": len(public_records),
                    "items": public_records,
                    "part": _text_attribute(part, "Name"),
                    "search_query": search_query,
                    "search_attempt_errors": search_attempt_errors,
                    "search_query_candidates": _topology_search_queries(kind_key),
                }
            )
        except Exception as exc:
            return _error(str(exc))

    names.append("catia_list_topology")

    @mcp.tool()
    def catia_add_fillet(
        edge_names: list[str],
        radius_mm: float,
        propagation_mode: int = 1,
        feature_name: str = "",
    ) -> dict[str, Any]:
        """Add a constant-radius solid edge fillet.

        edge_names accepts selectors returned by catia_list_topology, for
        example ["edge:1", "edge:4"]. Plain topology names are also accepted
        when they are unique.

        propagation_mode is passed to CATIA's fillet API. The commonly used
        value is 1.
        """
        feature = None
        warnings: list[str] = []

        try:
            if not isinstance(edge_names, list) or not edge_names:
                raise CATIAError("edge_names must contain at least one edge selector.")

            radius_value = _finite_positive(radius_mm, "radius_mm")
            propagation_value = _positive_integer(
                propagation_mode,
                "propagation_mode",
                0,
            )

            conn.ensure_connected()
            document = _active_document(conn)
            part = conn.get_active_part()
            body = conn.get_active_part_body()
            part.InWorkObject = body

            resolved_edges: list[Any] = []
            edge_details: list[dict[str, Any]] = []

            for selector in edge_names:
                edge, details = _resolve_topology_selector(
                    part,
                    document,
                    selector,
                    "edge",
                )
                resolved_edges.append(edge)
                edge_details.append(details)

            shape_factory, method_name, factory_details = _get_shape_factory(
                part,
                [
                    "AddNewSolidEdgeFilletWithConstantRadius",
                    "AddNewEdgeFilletWithConstantRadius",
                ],
            )

            feature = getattr(shape_factory, method_name)(
                resolved_edges[0],
                propagation_value,
                radius_value,
            )

            for edge in resolved_edges[1:]:
                feature.AddObjectToFillet(edge)

            try:
                feature.EdgePropagation = propagation_value
            except Exception as exc:
                warnings.append(f"EdgePropagation could not be set: {exc}")

            try:
                feature.Radius.Value = radius_value
            except Exception as exc:
                warnings.append(f"Radius read/write confirmation failed: {exc}")

            warnings.extend(
                _set_feature_name(
                    feature,
                    _normalise_name(feature_name, "MCP_EdgeFillet"),
                )
            )

            updated, update_strategy, update_warnings = _update_feature(
                part,
                feature,
            )
            warnings.extend(update_warnings)

            if not updated:
                return _feature_failure(
                    conn,
                    feature,
                    "Fillet was created but CATIA could not update the feature.",
                    feature_type="EdgeFillet",
                    warnings=warnings,
                    extra_data={
                        "radius_mm": radius_value,
                        "propagation_mode": propagation_value,
                        "edges": edge_details,
                        "factory": factory_details,
                        "api_method": method_name,
                        "update_strategy": update_strategy,
                    },
                )

            warnings.extend(_refresh_display(conn))

            return _success(
                {
                    "feature": _text_attribute(feature, "Name"),
                    "type": "EdgeFillet",
                    "radius_mm": radius_value,
                    "propagation_mode": propagation_value,
                    "edge_count": len(resolved_edges),
                    "edges": edge_details,
                    "api_method": method_name,
                    "factory": factory_details,
                    "feature_created": True,
                    "feature_persisted": True,
                    "update_succeeded": True,
                    "update_strategy": update_strategy,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _feature_failure(
                conn,
                feature,
                str(exc),
                feature_type="EdgeFillet",
                warnings=warnings,
            )

    names.append("catia_add_fillet")

    @mcp.tool()
    def catia_shell(
        thickness_mm: float,
        removed_face_names: list[str],
        feature_name: str = "",
    ) -> dict[str, Any]:
        """Create an inward shell and remove selected faces.

        removed_face_names accepts selectors returned by catia_list_topology,
        for example ["face:2"]. At least one face is required.

        The input thickness is used as CATIA's internal thickness; external
        thickness is set to zero.
        """
        feature = None
        warnings: list[str] = []

        try:
            thickness_value = _finite_positive(thickness_mm, "thickness_mm")
            if not isinstance(removed_face_names, list) or not removed_face_names:
                raise CATIAError(
                    "removed_face_names must contain at least one face selector."
                )

            conn.ensure_connected()
            document = _active_document(conn)
            part = conn.get_active_part()
            body = conn.get_active_part_body()
            part.InWorkObject = body

            resolved_faces: list[Any] = []
            face_details: list[dict[str, Any]] = []

            for selector in removed_face_names:
                face, details = _resolve_topology_selector(
                    part,
                    document,
                    selector,
                    "face",
                )
                resolved_faces.append(face)
                face_details.append(details)

            shape_factory, method_name, factory_details = _get_shape_factory(
                part,
                ["AddNewShell"],
            )

            feature = getattr(shape_factory, method_name)(
                resolved_faces[0],
                thickness_value,
                0.0,
            )

            for face in resolved_faces[1:]:
                feature.AddFaceToRemove(face)

            warnings.extend(
                _set_feature_name(
                    feature,
                    _normalise_name(feature_name, "MCP_Shell"),
                )
            )

            updated, update_strategy, update_warnings = _update_feature(
                part,
                feature,
            )
            warnings.extend(update_warnings)

            if not updated:
                return _feature_failure(
                    conn,
                    feature,
                    "Shell was created but CATIA could not update the feature.",
                    feature_type="Shell",
                    warnings=warnings,
                    extra_data={
                        "internal_thickness_mm": thickness_value,
                        "external_thickness_mm": 0.0,
                        "removed_faces": face_details,
                        "factory": factory_details,
                        "api_method": method_name,
                        "update_strategy": update_strategy,
                    },
                )

            warnings.extend(_refresh_display(conn))

            return _success(
                {
                    "feature": _text_attribute(feature, "Name"),
                    "type": "Shell",
                    "internal_thickness_mm": thickness_value,
                    "external_thickness_mm": 0.0,
                    "removed_face_count": len(resolved_faces),
                    "removed_faces": face_details,
                    "api_method": method_name,
                    "factory": factory_details,
                    "feature_created": True,
                    "feature_persisted": True,
                    "update_succeeded": True,
                    "update_strategy": update_strategy,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _feature_failure(
                conn,
                feature,
                str(exc),
                feature_type="Shell",
                warnings=warnings,
            )

    names.append("catia_shell")

    @mcp.tool()
    def catia_circular_pattern(
        feature_name: str,
        count: int,
        total_angle_deg: float = 360.0,
        axis_name: Optional[str] = None,
        pattern_name: str = "",
    ) -> dict[str, Any]:
        """Create an angular circular pattern of a PartBody feature.

        count is the total angular instance count, including the original
        feature.

        axis_name may be:
        - omitted, "z", or "xy": use the Z axis (PlaneXY normal)
        - "x"/"yz": use the X axis
        - "y"/"zx": use the Y axis
        - an edge selector returned by catia_list_topology
        - the name of a CATIA axis/line object

        For a complete 360-degree crown, angular spacing is 360/count.
        For a partial crown, spacing is total_angle/(count-1).
        """
        feature = None
        warnings: list[str] = []

        try:
            count_value = _positive_integer(count, "count", 2)
            total_angle_value = _finite_positive(
                total_angle_deg,
                "total_angle_deg",
            )
            if total_angle_value > 360.0:
                raise CATIAError(
                    "total_angle_deg must be less than or equal to 360."
                )

            conn.ensure_connected()
            document = _active_document(conn)
            part = conn.get_active_part()
            body = conn.get_active_part_body()
            part.InWorkObject = body

            source_feature = _find_shape(body, feature_name)
            axis_reference, axis_details = _resolve_pattern_axis(
                part,
                document,
                axis_name,
            )

            try:
                rotation_center = part.CreateReferenceFromName("")
            except Exception as exc:
                raise CATIAError(
                    f"Cannot create the empty rotation-center reference: {exc}"
                ) from exc

            if math.isclose(total_angle_value, 360.0, abs_tol=1e-9):
                angular_spacing = 360.0 / count_value
                crown_mode = "complete"
            else:
                angular_spacing = total_angle_value / (count_value - 1)
                crown_mode = "partial"

            shape_factory, method_name, factory_details = _get_shape_factory(
                part,
                ["AddNewCircPattern"],
            )

            # Signature:
            # shape, radial_count, angular_count, radial_step, angular_step,
            # radial_position, angular_position, rotation_center,
            # rotation_axis, reverse_axis, rotation_angle, radius_aligned.
            feature = getattr(shape_factory, method_name)(
                source_feature,
                1,
                count_value,
                0.0,
                angular_spacing,
                1,
                1,
                rotation_center,
                axis_reference,
                False,
                0.0,
                True,
            )

            warnings.extend(
                _set_feature_name(
                    feature,
                    _normalise_name(pattern_name, "MCP_CircularPattern"),
                )
            )

            updated, update_strategy, update_warnings = _update_feature(
                part,
                feature,
            )
            warnings.extend(update_warnings)

            if not updated:
                return _feature_failure(
                    conn,
                    feature,
                    "Circular pattern was created but CATIA could not update the feature.",
                    feature_type="CircularPattern",
                    warnings=warnings,
                    extra_data={
                        "source_feature": _text_attribute(source_feature, "Name"),
                        "count": count_value,
                        "total_angle_deg": total_angle_value,
                        "angular_spacing_deg": angular_spacing,
                        "crown_mode": crown_mode,
                        "axis": axis_details,
                        "factory": factory_details,
                        "api_method": method_name,
                        "update_strategy": update_strategy,
                    },
                )

            warnings.extend(_refresh_display(conn))

            return _success(
                {
                    "feature": _text_attribute(feature, "Name"),
                    "type": "CircularPattern",
                    "source_feature": _text_attribute(source_feature, "Name"),
                    "count": count_value,
                    "total_angle_deg": total_angle_value,
                    "angular_spacing_deg": angular_spacing,
                    "crown_mode": crown_mode,
                    "axis": axis_details,
                    "api_method": method_name,
                    "factory": factory_details,
                    "feature_created": True,
                    "feature_persisted": True,
                    "update_succeeded": True,
                    "update_strategy": update_strategy,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _feature_failure(
                conn,
                feature,
                str(exc),
                feature_type="CircularPattern",
                warnings=warnings,
            )

    names.append("catia_circular_pattern")

    return names
