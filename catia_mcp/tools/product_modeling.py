from __future__ import annotations

import math
from typing import Any, Callable

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


def _finite_nonnegative(value: Any, name: str) -> float:
    number = _finite_number(value, name)
    if number < 0.0:
        raise CATIAError(f"{name} must be greater than or equal to 0.")
    return number


def _normalise_name(value: Any, default: str) -> str:
    name = str(value).strip()
    return name or default


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


def _get_shape_factory(part: Any, required_method: str) -> tuple[Any, dict[str, Any]]:
    """Resolve ShapeFactory when pywin32 exposes it as a generic COM Factory."""

    try:
        raw_factory = part.ShapeFactory
    except Exception as exc:
        raise CATIAError(f"Cannot access Part.ShapeFactory: {exc}") from exc

    details: dict[str, Any] = {
        "required_method": required_method,
        "raw": _describe_com_object(raw_factory),
        "dispatch_used": False,
    }

    if _has_callable(raw_factory, required_method):
        details["resolved"] = _describe_com_object(raw_factory)
        details["required_method_available"] = True
        return raw_factory, details

    try:
        import win32com.client  # type: ignore

        shape_factory = win32com.client.Dispatch(raw_factory)
    except Exception as exc:
        details["dispatch_error"] = str(exc)
        raise CATIAError(
            "Part.ShapeFactory was returned as a generic COM Factory and "
            f"could not be dynamically dispatched: {exc}"
        ) from exc

    details["dispatch_used"] = True
    details["resolved"] = _describe_com_object(shape_factory)
    details["required_method_available"] = _has_callable(
        shape_factory,
        required_method,
    )

    if not details["required_method_available"]:
        raise CATIAError(
            "ShapeFactory dispatch completed, but the resolved COM object "
            f"does not expose {required_method}."
        )

    return shape_factory, details


def _update_object(part: Any, obj: Any) -> tuple[str, list[str]]:
    """Update one object, falling back to a full Part.Update when necessary."""

    warnings: list[str] = []

    try:
        part.UpdateObject(obj)
        return "UpdateObject", warnings
    except Exception as exc:
        warnings.append(f"UpdateObject failed: {exc}")

    try:
        part.Update()
        warnings.append("Part.Update fallback succeeded.")
        return "Part.Update", warnings
    except Exception as exc:
        warnings.append(f"Part.Update fallback failed: {exc}")
        raise CATIAError("; ".join(warnings)) from exc


def _set_name(obj: Any, requested_name: str, object_label: str) -> list[str]:
    warnings: list[str] = []
    try:
        obj.Name = requested_name
    except Exception as exc:
        warnings.append(
            f"{object_label} was created but could not be renamed to "
            f"'{requested_name}': {exc}"
        )
    return warnings


def _set_part_number(part: Any, requested_name: str) -> list[str]:
    warnings: list[str] = []
    try:
        part.PartNumber = requested_name
    except Exception as exc:
        warnings.append(
            f"The CATPart was created but PartNumber could not be set to "
            f"'{requested_name}': {exc}"
        )
    return warnings


def _refresh_display(conn: Any) -> list[str]:
    warnings: list[str] = []
    try:
        conn.refresh_display()
    except Exception as exc:
        warnings.append(f"Model created, but display refresh failed: {exc}")
    return warnings


def _describe_document(conn: Any, doc: Any) -> tuple[dict[str, Any], list[str]]:
    warnings: list[str] = []
    try:
        return conn.describe_document(doc), warnings
    except Exception as exc:
        warnings.append(f"Document description could not be generated: {exc}")
        return {
            "name": str(getattr(doc, "Name", "")),
            "type": type(doc).__name__,
        }, warnings


def _close_failed_document(doc: Any) -> tuple[bool, list[str]]:
    """Close an unsaved CATPart created by a failed high-level operation."""

    if doc is None:
        return True, []

    warnings: list[str] = []
    try:
        doc.Close()
        return True, warnings
    except Exception as exc:
        warnings.append(
            "The operation failed and the temporary CATPart could not be closed: "
            f"{exc}"
        )
        return False, warnings


def _failure_with_document_rollback(
    exc: Exception,
    doc: Any,
    *,
    warnings: list[str] | None = None,
) -> dict[str, Any]:
    warning_list = list(warnings or [])
    document_created = doc is not None
    rollback_succeeded, rollback_warnings = _close_failed_document(doc)
    warning_list.extend(rollback_warnings)

    document_persisted = document_created and not rollback_succeeded
    return _error(
        str(exc),
        status="partial_success" if document_persisted else "error",
        warnings=warning_list,
        data={
            "document_created": document_created,
            "document_persisted": document_persisted,
            "rollback_succeeded": rollback_succeeded,
        },
    )


def _create_sketch_geometry(
    sketch: Any,
    builder: Callable[[Any], None],
) -> None:
    """Open a sketch, build geometry, and always leave sketch edit mode."""

    opened = False
    try:
        factory = sketch.OpenEdition()
        opened = True
        builder(factory)
    finally:
        if opened:
            sketch.CloseEdition()


def _validate_holes(
    holes: list[dict[str, float]] | None,
) -> list[dict[str, float]]:
    validated: list[dict[str, float]] = []

    for index, hole in enumerate(holes or [], start=1):
        if not isinstance(hole, dict):
            raise CATIAError(
                f"holes[{index}] must be an object containing x, y, and diameter."
            )

        if "diameter" not in hole:
            raise CATIAError(f"holes[{index}].diameter is required.")

        validated.append(
            {
                "x": _finite_number(hole.get("x", 0.0), f"holes[{index}].x"),
                "y": _finite_number(hole.get("y", 0.0), f"holes[{index}].y"),
                "diameter": _finite_positive(
                    hole.get("diameter"),
                    f"holes[{index}].diameter",
                ),
            }
        )

    return validated


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = ctx.conn
    names: list[str] = []

    @mcp.tool()
    def catia_create_rounded_rect_slab(
        name: str = "RoundedRectSlab",
        width: float = 100.0,
        height: float = 100.0,
        thickness: float = 10.0,
        corner_radius: float = 5.0,
    ) -> dict[str, Any]:
        """Create a rounded-rectangle slab in a new CATPart.

        The outline is created with four lines and four Factory2D.CreateCircle
        arc segments. A corner radius of 0 creates a normal rectangular slab.
        """

        doc = None
        warnings: list[str] = []

        try:
            width_value = _finite_positive(width, "width")
            height_value = _finite_positive(height, "height")
            thickness_value = _finite_positive(thickness, "thickness")
            radius_value = _finite_nonnegative(corner_radius, "corner_radius")
            part_name = _normalise_name(name, "RoundedRectSlab")

            maximum_radius = min(width_value, height_value) / 2.0
            if radius_value > 0.0 and radius_value >= maximum_radius:
                raise CATIAError(
                    "corner_radius must be smaller than half of the smaller "
                    f"slab dimension ({maximum_radius})."
                )

            app = conn.connect(visible=True)
            doc = app.Documents.Add("Part")
            part = doc.Part
            body = part.MainBody
            part.InWorkObject = body

            warnings.extend(_set_part_number(part, part_name))

            plane_ref = conn.get_origin_plane_reference("xy")
            sketch = body.Sketches.Add(plane_ref)
            warnings.extend(
                _set_name(sketch, "BaseRoundedRectSketch", "Base sketch")
            )

            left = -width_value / 2.0
            right = width_value / 2.0
            bottom = -height_value / 2.0
            top = height_value / 2.0
            r = radius_value

            def build_outline(factory: Any) -> None:
                if r <= 0.0:
                    factory.CreateLine(left, bottom, right, bottom)
                    factory.CreateLine(right, bottom, right, top)
                    factory.CreateLine(right, top, left, top)
                    factory.CreateLine(left, top, left, bottom)
                    return

                factory.CreateLine(left + r, bottom, right - r, bottom)
                factory.CreateLine(right, bottom + r, right, top - r)
                factory.CreateLine(right - r, top, left + r, top)
                factory.CreateLine(left, top - r, left, bottom + r)

                # Factory2D has no CreateArc method. The five-argument
                # CreateCircle method creates an arc segment.
                factory.CreateCircle(
                    right - r,
                    bottom + r,
                    r,
                    1.5 * math.pi,
                    2.0 * math.pi,
                )
                factory.CreateCircle(
                    right - r,
                    top - r,
                    r,
                    0.0,
                    0.5 * math.pi,
                )
                factory.CreateCircle(
                    left + r,
                    top - r,
                    r,
                    0.5 * math.pi,
                    math.pi,
                )
                factory.CreateCircle(
                    left + r,
                    bottom + r,
                    r,
                    math.pi,
                    1.5 * math.pi,
                )

            _create_sketch_geometry(sketch, build_outline)

            sketch_update_strategy, sketch_update_warnings = _update_object(
                part,
                sketch,
            )
            warnings.extend(sketch_update_warnings)

            shape_factory, factory_details = _get_shape_factory(part, "AddNewPad")
            part.InWorkObject = body
            pad = shape_factory.AddNewPad(sketch, thickness_value)
            warnings.extend(_set_name(pad, "BaseSlab", "Pad"))

            pad_update_strategy, pad_update_warnings = _update_object(part, pad)
            warnings.extend(pad_update_warnings)
            warnings.extend(_refresh_display(conn))

            document_details, document_warnings = _describe_document(conn, doc)
            warnings.extend(document_warnings)

            return _success(
                {
                    "document": document_details,
                    "base_feature": str(getattr(pad, "Name", "")),
                    "width": width_value,
                    "height": height_value,
                    "thickness": thickness_value,
                    "corner_radius": radius_value,
                    "sketch_update_strategy": sketch_update_strategy,
                    "feature_update_strategy": pad_update_strategy,
                    "factory": factory_details,
                    "document_created": True,
                    "document_persisted": True,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _failure_with_document_rollback(
                exc,
                doc,
                warnings=warnings,
            )

    names.append("catia_create_rounded_rect_slab")

    @mcp.tool()
    def catia_create_rect_plate_with_holes(
        name: str = "PlateWithHoles",
        width: float = 100.0,
        height: float = 100.0,
        thickness: float = 10.0,
        holes: list[dict[str, float]] | None = None,
    ) -> dict[str, Any]:
        """Create a rectangular plate and cut all circular holes in one Pocket.

        Hole format:
        [
            {"x": 0, "y": 0, "diameter": 10},
            {"x": 30, "y": 30, "diameter": 6}
        ]

        Hole circles are created on the same XY origin plane as the base
        profile. One Pocket cuts all circles through the plate.
        """

        doc = None
        warnings: list[str] = []

        try:
            width_value = _finite_positive(width, "width")
            height_value = _finite_positive(height, "height")
            thickness_value = _finite_positive(thickness, "thickness")
            validated_holes = _validate_holes(holes)
            part_name = _normalise_name(name, "PlateWithHoles")

            # Validate all user data before creating a CATPart.
            app = conn.connect(visible=True)
            doc = app.Documents.Add("Part")
            part = doc.Part
            body = part.MainBody
            part.InWorkObject = body

            warnings.extend(_set_part_number(part, part_name))

            plane_ref = conn.get_origin_plane_reference("xy")

            base_sketch = body.Sketches.Add(plane_ref)
            warnings.extend(
                _set_name(base_sketch, "PlateBaseSketch", "Base sketch")
            )

            x1 = -width_value / 2.0
            y1 = -height_value / 2.0
            x2 = width_value / 2.0
            y2 = height_value / 2.0

            def build_base(factory: Any) -> None:
                factory.CreateLine(x1, y1, x2, y1)
                factory.CreateLine(x2, y1, x2, y2)
                factory.CreateLine(x2, y2, x1, y2)
                factory.CreateLine(x1, y2, x1, y1)

            _create_sketch_geometry(base_sketch, build_base)

            base_sketch_update_strategy, base_sketch_warnings = _update_object(
                part,
                base_sketch,
            )
            warnings.extend(base_sketch_warnings)

            pad_factory, pad_factory_details = _get_shape_factory(
                part,
                "AddNewPad",
            )
            part.InWorkObject = body
            pad = pad_factory.AddNewPad(base_sketch, thickness_value)
            warnings.extend(_set_name(pad, "PlateBody", "Base Pad"))

            pad_update_strategy, pad_update_warnings = _update_object(part, pad)
            warnings.extend(pad_update_warnings)

            hole_feature_names: list[str] = []
            hole_sketch_name = ""
            hole_sketch_update_strategy: str | None = None
            pocket_update_strategy: str | None = None
            pocket_factory_details: dict[str, Any] | None = None

            if validated_holes:
                # Use the same XY origin plane rather than an offset plane.
                # The Pocket starts at the base plane and cuts in the Pad's
                # default direction through the complete thickness.
                hole_sketch = body.Sketches.Add(plane_ref)
                warnings.extend(
                    _set_name(hole_sketch, "HolePatternSketch", "Hole sketch")
                )
                hole_sketch_name = str(getattr(hole_sketch, "Name", ""))

                def build_holes(factory: Any) -> None:
                    for hole in validated_holes:
                        factory.CreateClosedCircle(
                            hole["x"],
                            hole["y"],
                            hole["diameter"] / 2.0,
                        )

                _create_sketch_geometry(hole_sketch, build_holes)

                (
                    hole_sketch_update_strategy,
                    hole_sketch_warnings,
                ) = _update_object(part, hole_sketch)
                warnings.extend(hole_sketch_warnings)

                pocket_factory, pocket_factory_details = _get_shape_factory(
                    part,
                    "AddNewPocket",
                )
                part.InWorkObject = body
                pocket = pocket_factory.AddNewPocket(
                    hole_sketch,
                    thickness_value + 1.0,
                )
                warnings.extend(_set_name(pocket, "HoleCuts", "Pocket"))

                pocket_update_strategy, pocket_update_warnings = _update_object(
                    part,
                    pocket,
                )
                warnings.extend(pocket_update_warnings)
                hole_feature_names.append(str(getattr(pocket, "Name", "")))

            warnings.extend(_refresh_display(conn))
            document_details, document_warnings = _describe_document(conn, doc)
            warnings.extend(document_warnings)

            return _success(
                {
                    "document": document_details,
                    "base_feature": str(getattr(pad, "Name", "")),
                    "hole_features": hole_feature_names,
                    "hole_count": len(validated_holes),
                    "hole_sketch": hole_sketch_name,
                    "width": width_value,
                    "height": height_value,
                    "thickness": thickness_value,
                    "holes": validated_holes,
                    "base_sketch_update_strategy": base_sketch_update_strategy,
                    "base_feature_update_strategy": pad_update_strategy,
                    "hole_sketch_update_strategy": hole_sketch_update_strategy,
                    "hole_feature_update_strategy": pocket_update_strategy,
                    "pad_factory": pad_factory_details,
                    "pocket_factory": pocket_factory_details,
                    "document_created": True,
                    "document_persisted": True,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _failure_with_document_rollback(
                exc,
                doc,
                warnings=warnings,
            )

    names.append("catia_create_rect_plate_with_holes")

    @mcp.tool()
    def catia_create_cylinder(
        name: str = "Cylinder",
        radius: float = 10.0,
        height: float = 20.0,
    ) -> dict[str, Any]:
        """Create a simple cylinder in a new CATPart."""

        doc = None
        warnings: list[str] = []

        try:
            radius_value = _finite_positive(radius, "radius")
            height_value = _finite_positive(height, "height")
            part_name = _normalise_name(name, "Cylinder")

            app = conn.connect(visible=True)
            doc = app.Documents.Add("Part")
            part = doc.Part
            body = part.MainBody
            part.InWorkObject = body

            warnings.extend(_set_part_number(part, part_name))

            plane_ref = conn.get_origin_plane_reference("xy")
            sketch = body.Sketches.Add(plane_ref)
            warnings.extend(_set_name(sketch, "CylinderSketch", "Cylinder sketch"))

            def build_circle(factory: Any) -> None:
                factory.CreateClosedCircle(0.0, 0.0, radius_value)

            _create_sketch_geometry(sketch, build_circle)

            sketch_update_strategy, sketch_update_warnings = _update_object(
                part,
                sketch,
            )
            warnings.extend(sketch_update_warnings)

            shape_factory, factory_details = _get_shape_factory(part, "AddNewPad")
            part.InWorkObject = body
            pad = shape_factory.AddNewPad(sketch, height_value)
            warnings.extend(_set_name(pad, "CylinderBody", "Cylinder Pad"))

            pad_update_strategy, pad_update_warnings = _update_object(part, pad)
            warnings.extend(pad_update_warnings)
            warnings.extend(_refresh_display(conn))

            document_details, document_warnings = _describe_document(conn, doc)
            warnings.extend(document_warnings)

            return _success(
                {
                    "document": document_details,
                    "feature": str(getattr(pad, "Name", "")),
                    "radius": radius_value,
                    "height": height_value,
                    "sketch_update_strategy": sketch_update_strategy,
                    "feature_update_strategy": pad_update_strategy,
                    "factory": factory_details,
                    "document_created": True,
                    "document_persisted": True,
                    "rollback_succeeded": None,
                },
                warnings,
            )
        except Exception as exc:
            return _failure_with_document_rollback(
                exc,
                doc,
                warnings=warnings,
            )

    names.append("catia_create_cylinder")

    return names
