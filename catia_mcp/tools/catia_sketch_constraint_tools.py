from __future__ import annotations

import math
from typing import Any

from catia_mcp.connection import CATIAError


# CatConstraintType values from the CATIA V5 Automation API.
_CAT_CST_TYPE_ON = 2
_CAT_CST_TYPE_CONCENTRICITY = 3
_CAT_CST_TYPE_TANGENCY = 4
_CAT_CST_TYPE_LENGTH = 5
_CAT_CST_TYPE_PARALLELISM = 8
_CAT_CST_TYPE_HORIZONTALITY = 10
_CAT_CST_TYPE_PERPENDICULARITY = 11
_CAT_CST_TYPE_VERTICALITY = 13
_CAT_CST_TYPE_RADIUS = 14


_GEOMETRIC_CONSTRAINTS: dict[str, tuple[int, int]] = {
    "coincident": (_CAT_CST_TYPE_ON, 2),
    "concentric": (_CAT_CST_TYPE_CONCENTRICITY, 2),
    "tangent": (_CAT_CST_TYPE_TANGENCY, 2),
    "horizontal": (_CAT_CST_TYPE_HORIZONTALITY, 1),
    "vertical": (_CAT_CST_TYPE_VERTICALITY, 1),
    "parallel": (_CAT_CST_TYPE_PARALLELISM, 2),
    "perpendicular": (_CAT_CST_TYPE_PERPENDICULARITY, 2),
}

_GEOMETRIC_ALIASES: dict[str, str] = {
    "coincidence": "coincident",
    "on": "coincident",
    "concentricity": "concentric",
    "tangency": "tangent",
    "horizontality": "horizontal",
    "verticality": "vertical",
    "parallelism": "parallel",
    "perpendicularity": "perpendicular",
}

_SUPPORTED_DIMENSION_TYPES = {"length", "radius", "diameter"}


def _normalize_name(value: str) -> str:
    return value.strip().lower().replace("-", "_").replace(" ", "_")


def _get_active_sketch(ctx: Any) -> tuple[Any, Any]:
    part = ctx.conn.get_active_part()
    sketch = part.InWorkObject

    if not hasattr(sketch, "Constraints") or not hasattr(sketch, "GeometricElements"):
        raise CATIAError(
            "Active object is not a Sketch. Open or activate a sketch before "
            "calling a sketch constraint tool."
        )

    return part, sketch


def _find_geometric_element(sketch: Any, element_name: str) -> Any:
    name = str(element_name).strip()
    if not name:
        raise CATIAError("Geometry element name cannot be empty.")

    geometric_elements = sketch.GeometricElements

    try:
        return geometric_elements.Item(name)
    except Exception as direct_error:
        # Some CATIA installations/locales are more reliable with an explicit
        # case-insensitive scan than with Item(name).
        wanted = name.casefold()
        for index in range(1, geometric_elements.Count + 1):
            item = geometric_elements.Item(index)
            if str(getattr(item, "Name", "")).casefold() == wanted:
                return item

        raise CATIAError(
            f"Geometry element not found in active sketch: {name}"
        ) from direct_error


def _safe_constraint_property(constraint: Any, property_name: str) -> Any | None:
    try:
        return getattr(constraint, property_name)
    except Exception:
        return None


def _safe_dimension_value(constraint: Any) -> float | None:
    try:
        return float(constraint.Dimension.Value)
    except Exception:
        return None


def _remove_constraint_safely(constraints: Any, constraint: Any) -> None:
    try:
        name = getattr(constraint, "Name", "")
        if name:
            constraints.Remove(name)
            return
    except Exception:
        pass

    try:
        constraints.Remove(constraints.Count)
    except Exception:
        pass


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    """Register CATIA sketch constraint tools."""

    @mcp.tool()
    def catia_add_geometric_constraint(
        constraint_type: str,
        element_names: list[str],
    ) -> dict[str, Any]:
        """Add a geometric constraint to geometry in the active sketch.

        Supported types:
        coincident, concentric, tangent, horizontal, vertical, parallel,
        perpendicular.

        Horizontal and vertical require one line. All other supported types
        require exactly two compatible sketch elements.
        """
        try:
            _, sketch = _get_active_sketch(ctx)

            normalized_type = _normalize_name(constraint_type)
            normalized_type = _GEOMETRIC_ALIASES.get(
                normalized_type,
                normalized_type,
            )

            definition = _GEOMETRIC_CONSTRAINTS.get(normalized_type)
            if definition is None:
                supported = ", ".join(_GEOMETRIC_CONSTRAINTS)
                raise CATIAError(
                    f"Unsupported geometric constraint: {constraint_type}. "
                    f"Supported types: {supported}."
                )

            if not isinstance(element_names, list):
                raise CATIAError("element_names must be a list of geometry names.")

            constraint_enum, required_count = definition
            if len(element_names) != required_count:
                raise CATIAError(
                    f"{normalized_type} constraint requires exactly "
                    f"{required_count} element name(s); received "
                    f"{len(element_names)}."
                )

            elements = [
                _find_geometric_element(sketch, name)
                for name in element_names
            ]

            constraints = sketch.Constraints
            if required_count == 1:
                constraint = constraints.AddMonoEltCst(
                    constraint_enum,
                    elements[0],
                )
            else:
                constraint = constraints.AddBiEltCst(
                    constraint_enum,
                    elements[0],
                    elements[1],
                )

            # Do not call Part.Update while the sketch is in edition. CATIA may
            # reject a part-level update in this state. The sketcher workflow
            # should close the sketch and update the part afterwards.
            return {
                "ok": True,
                "data": {
                    "constraint_name": getattr(constraint, "Name", ""),
                    "type": normalized_type,
                    "catia_constraint_type": constraint_enum,
                    "element_names": list(element_names),
                    "element_count": required_count,
                    "update_deferred": True,
                    "message": (
                        "Constraint created. Close the sketch and update the "
                        "part to solve the complete sketch."
                    ),
                },
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    @mcp.tool()
    def catia_add_dimension_constraint(
        element_name: str,
        value_mm: float,
        dim_type: str = "length",
    ) -> dict[str, Any]:
        """Add a length, radius, or diameter-driving constraint.

        CATIA V5 has native sketch constraint types for length and radius, but
        no separate CatConstraintType value for diameter. A diameter request is
        therefore enforced with a native radius constraint whose value is half
        the requested diameter.
        """
        constraint = None
        constraints = None

        try:
            _, sketch = _get_active_sketch(ctx)

            normalized_type = _normalize_name(dim_type)
            if normalized_type not in _SUPPORTED_DIMENSION_TYPES:
                supported = ", ".join(sorted(_SUPPORTED_DIMENSION_TYPES))
                raise CATIAError(
                    f"Unsupported dimension type: {dim_type}. "
                    f"Supported types: {supported}."
                )

            requested_value = float(value_mm)
            if not math.isfinite(requested_value):
                raise CATIAError("Dimension value must be a finite number.")
            if requested_value <= 0:
                raise CATIAError("Dimension value must be positive.")

            element = _find_geometric_element(sketch, element_name)
            constraints = sketch.Constraints

            if normalized_type == "length":
                constraint_enum = _CAT_CST_TYPE_LENGTH
                native_type = "length"
                native_value = requested_value
            elif normalized_type == "radius":
                constraint_enum = _CAT_CST_TYPE_RADIUS
                native_type = "radius"
                native_value = requested_value
            else:
                constraint_enum = _CAT_CST_TYPE_RADIUS
                native_type = "radius"
                native_value = requested_value / 2.0

            constraint = constraints.AddMonoEltCst(
                constraint_enum,
                element,
            )
            constraint.Dimension.Value = native_value

            data: dict[str, Any] = {
                "constraint_name": getattr(constraint, "Name", ""),
                "element": str(element_name),
                "requested_dimension_type": normalized_type,
                "requested_value_mm": requested_value,
                "native_constraint_type": native_type,
                "native_value_mm": native_value,
                "catia_constraint_type": constraint_enum,
                "update_deferred": True,
                "message": (
                    "Dimension constraint created. Close the sketch and update "
                    "the part to solve the complete sketch."
                ),
            }

            if normalized_type == "diameter":
                data["note"] = (
                    "CATIA V5 exposes no separate diameter value in "
                    "CatConstraintType. The requested diameter is enforced as "
                    "a radius constraint with half the requested value."
                )

            return {"ok": True, "data": data}
        except Exception as exc:
            # If CATIA created the constraint but rejected the value assignment,
            # remove the partial constraint to avoid leaving a residual constraint.
            if constraint is not None and constraints is not None:
                _remove_constraint_safely(constraints, constraint)
            return {"ok": False, "error": str(exc)}

    @mcp.tool()
    def catia_list_constraints() -> dict[str, Any]:
        """List constraints and readable status data in the active sketch."""
        try:
            _, sketch = _get_active_sketch(ctx)
            constraint_collection = sketch.Constraints
            result: list[dict[str, Any]] = []

            for index in range(1, constraint_collection.Count + 1):
                constraint = constraint_collection.Item(index)
                result.append(
                    {
                        "index": index,
                        "name": getattr(constraint, "Name", ""),
                        "type": _safe_constraint_property(constraint, "Type"),
                        "status": str(
                            _safe_constraint_property(constraint, "Status")
                        ),
                        "mode": _safe_constraint_property(constraint, "Mode"),
                        "value": _safe_dimension_value(constraint),
                    }
                )

            return {"ok": True, "data": result}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    return [
        "catia_add_geometric_constraint",
        "catia_add_dimension_constraint",
        "catia_list_constraints",
    ]
