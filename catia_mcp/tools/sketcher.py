"""
sketcher.py
Version: sketcher-fixed-2026-08-04-v3

CATIA V5 MCP sketch creation and 2D geometry tools.

v3 corrections:
- Every result carries an implementation version and consistent status.
- Sketch creation validates plane/offset and returns a local-to-global plane
  coordinate contract; callers must not assume sketch x/y are global X/Y.
- Sketch creation and composite geometry operations clean up partial objects on
  failure and report partial_success when rollback cannot be verified.
- Closing a sketch clears the in-memory session only after CloseEdition succeeds;
  close/update failures are reported truthfully without a stale factory wrapper.
- Lines, rectangles, points, circles and arcs validate finite/nondegenerate
  inputs and verify GeometricElements collection deltas.
"""

from __future__ import annotations

import math
from typing import Any, Callable

from catia_mcp.connection import CATIAError


IMPLEMENTATION_VERSION = "sketcher-fixed-2026-08-04-v3"
_ANGLE_EPSILON_DEG = 1e-9

_PLANE_CONTRACTS = {
    "xy": {
        "origin_for_offset": lambda d: [0.0, 0.0, d],
        "local_u_global": [1.0, 0.0, 0.0],
        "local_v_global": [0.0, 1.0, 0.0],
        "nominal_normal_global": [0.0, 0.0, 1.0],
    },
    "yz": {
        "origin_for_offset": lambda d: [d, 0.0, 0.0],
        "local_u_global": [0.0, 1.0, 0.0],
        "local_v_global": [0.0, 0.0, 1.0],
        "nominal_normal_global": [1.0, 0.0, 0.0],
    },
    "zx": {
        "origin_for_offset": lambda d: [0.0, d, 0.0],
        "local_u_global": [0.0, 0.0, 1.0],
        "local_v_global": [1.0, 0.0, 0.0],
        "nominal_normal_global": [0.0, 1.0, 0.0],
    },
}


def _success(data: Any, warnings: list[str] | None = None) -> dict[str, Any]:
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
    data: Any | None = None,
    warnings: list[str] | None = None,
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


def _require_finite(value: Any, label: str) -> float:
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise CATIAError(f"{label} must be a number.") from exc
    if not math.isfinite(number):
        raise CATIAError(f"{label} must be a finite number.")
    return number


def _require_positive(value: Any, label: str) -> float:
    number = _require_finite(value, label)
    if number <= 0.0:
        raise CATIAError(f"{label} must be greater than 0.")
    return number


def _normalise_plane(value: Any) -> str:
    plane = str(value).strip().lower().replace("xz", "zx")
    if plane not in _PLANE_CONTRACTS:
        raise CATIAError("plane must be one of: xy, yz, zx.")
    return plane


def _plane_coordinate_contract(plane: Any, offset: Any) -> dict[str, Any]:
    plane_key = _normalise_plane(plane)
    offset_value = _require_finite(offset, "offset")
    base = _PLANE_CONTRACTS[plane_key]
    return {
        "plane": plane_key,
        "offset_mm": offset_value,
        "sketch_origin_global_mm": base["origin_for_offset"](offset_value),
        "local_u_global": list(base["local_u_global"]),
        "local_v_global": list(base["local_v_global"]),
        "nominal_normal_global": list(base["nominal_normal_global"]),
        "normal_sign_verified_from_catia": False,
        "coordinate_warning": (
            "Sketch x/y are local u/v coordinates. For offset-plane Pad/Pocket "
            "operations, use a geometry-effect-verified direction policy rather "
            "than assuming the nominal plane-normal sign."
        ),
    }


def _to_catia_arc_parameters(
    start_angle_deg: float,
    end_angle_deg: float,
) -> tuple[float, float, float]:
    start_deg = _require_finite(start_angle_deg, "Arc start angle")
    end_deg = _require_finite(end_angle_deg, "Arc end angle")
    raw_sweep_deg = end_deg - start_deg

    if math.isclose(raw_sweep_deg, 0.0, abs_tol=_ANGLE_EPSILON_DEG):
        raise CATIAError(
            "Arc start and end angles must be different. Use catia_sketch_circle for a full circle."
        )
    if abs(raw_sweep_deg) >= 360.0 - _ANGLE_EPSILON_DEG:
        raise CATIAError(
            "Arc sweep must be less than 360 degrees. Use catia_sketch_circle for a full circle."
        )

    sweep_deg = raw_sweep_deg if raw_sweep_deg > 0.0 else raw_sweep_deg + 360.0
    normalized_start_deg = start_deg % 360.0
    start_param = math.radians(normalized_start_deg)
    end_param = start_param + math.radians(sweep_deg)
    return start_param, end_param, sweep_deg


def _active_document(conn: Any) -> Any:
    try:
        return conn.app.ActiveDocument
    except Exception as exc:
        raise CATIAError("Cannot access the active CATIA document.") from exc


def _delete_objects(conn: Any, objects: list[Any]) -> dict[str, Any]:
    unique = [obj for obj in objects if obj is not None]
    result = {
        "attempted_count": len(unique),
        "deleted_count": 0,
        "failed_count": 0,
        "errors": [],
        "verified": True,
    }
    if not unique:
        return result
    try:
        selection = _active_document(conn).Selection
        for obj in reversed(unique):
            try:
                selection.Clear()
                selection.Add(obj)
                selection.Delete()
                result["deleted_count"] += 1
            except Exception as exc:
                result["failed_count"] += 1
                result["errors"].append(str(exc))
            finally:
                try:
                    selection.Clear()
                except Exception:
                    pass
    except Exception as exc:
        result["failed_count"] = len(unique)
        result["errors"].append(str(exc))
    result["verified"] = result["failed_count"] == 0
    return result


class SketchSession:
    def __init__(self) -> None:
        self.active_sketch: Any | None = None
        self.active_factory: Any | None = None
        self.document_name: str = ""

    def ensure_open(self, conn: Any | None = None) -> None:
        if self.active_sketch is None or self.active_factory is None:
            raise CATIAError("No active sketch. Call catia_create_sketch first.")
        if conn is not None and self.document_name:
            try:
                current_name = str(getattr(_active_document(conn), "Name", ""))
            except Exception as exc:
                raise CATIAError(
                    f"Cannot verify the document that owns the active sketch: {exc}"
                ) from exc
            if current_name and current_name.casefold() != self.document_name.casefold():
                raise CATIAError(
                    "The active CATIA document changed while a sketch edition was open. "
                    f"Expected '{self.document_name}', current '{current_name}'. "
                    "Reactivate the owning CATPart before continuing or closing the sketch."
                )

    def reset(self) -> None:
        self.active_sketch = None
        self.active_factory = None
        self.document_name = ""


def _geometry_count(sketch: Any) -> int:
    return int(sketch.GeometricElements.Count)


def _create_geometry_transaction(
    conn: Any,
    session: SketchSession,
    *,
    label: str,
    expected_count: int | None,
    creator: Callable[[list[Any]], dict[str, Any]],
) -> dict[str, Any]:
    sketch: Any | None = None
    before: int | None = None
    created: list[Any] = []
    try:
        session.ensure_open(conn)
        sketch = session.active_sketch
        before = _geometry_count(sketch)
        payload = creator(created)
        after = _geometry_count(sketch)
        delta = after - before
        verified = (
            delta == expected_count
            if expected_count is not None
            else delta == len(created)
        )
        if not verified:
            rollback = _delete_objects(conn, created)
            after_rollback = _geometry_count(sketch)
            rollback["count_after_rollback"] = after_rollback
            rollback["count_restored"] = after_rollback == before
            rollback["verified"] = bool(
                rollback["verified"] and rollback["count_restored"]
            )
            return _error(
                f"{label} collection-delta verification failed.",
                status="error" if rollback["verified"] else "partial_success",
                data={
                    **payload,
                    "geometry_count_before": before,
                    "geometry_count_after": after,
                    "geometry_count_delta": delta,
                    "expected_delta": expected_count,
                    "creation_verified": False,
                    "rollback": rollback,
                    "model_modified": not rollback["verified"],
                    "document_save_required": not rollback["verified"],
                },
            )
        return _success(
            {
                **payload,
                "geometry_count_before": before,
                "geometry_count_after": after,
                "geometry_count_delta": delta,
                "expected_delta": expected_count,
                "creation_verified": True,
                "model_modified": True,
                "document_save_required": True,
            }
        )
    except Exception as exc:
        rollback = _delete_objects(conn, created)
        after_rollback: int | None = None
        if sketch is not None:
            try:
                after_rollback = _geometry_count(sketch)
            except Exception:
                pass
        count_restored = (
            before is None
            or (after_rollback is not None and after_rollback == before)
        )
        rollback["count_after_rollback"] = after_rollback
        rollback["count_restored"] = count_restored
        rollback["verified"] = bool(rollback["verified"] and count_restored)
        return _error(
            str(exc),
            status="error" if rollback["verified"] else "partial_success",
            data={
                "type": label,
                "geometry_count_before": before,
                "geometry_count_after_rollback": after_rollback,
                "rollback": rollback,
                "model_modified": not rollback["verified"],
                "document_save_required": not rollback["verified"],
            },
        )


def _create_rectangle_geometry(factory: Any, x1: float, y1: float, x2: float, y2: float, created: list[Any]) -> None:
    created.append(factory.CreateLine(x1, y1, x2, y1))
    created.append(factory.CreateLine(x2, y1, x2, y2))
    created.append(factory.CreateLine(x2, y2, x1, y2))
    created.append(factory.CreateLine(x1, y2, x1, y1))


def _create_circle_arc(factory: Any, center_x: float, center_y: float, radius: float, start_angle_deg: float, end_angle_deg: float) -> Any:
    start_param, end_param, _ = _to_catia_arc_parameters(start_angle_deg, end_angle_deg)
    return factory.CreateCircle(center_x, center_y, radius, start_param, end_param)


def _create_line_if_nonzero(factory: Any, x1: float, y1: float, x2: float, y2: float) -> Any | None:
    if math.isclose(x1, x2, abs_tol=1e-12) and math.isclose(y1, y2, abs_tol=1e-12):
        return None
    return factory.CreateLine(x1, y1, x2, y2)


def register_tools(mcp: Any, ctx: Any) -> list[str]:
    conn = ctx.conn
    session = SketchSession()
    names: list[str] = []

    @mcp.tool()
    def catia_create_sketch(
        plane: str = "xy",
        sketch_name: str = "",
        offset: float = 0.0,
    ) -> dict[str, Any]:
        """Create and open a sketch on an origin or offset plane."""
        sketch = None
        opened = False
        contract: dict[str, Any] | None = None
        try:
            if session.active_sketch is not None:
                raise CATIAError(
                    "A sketch is already open in this MCP session. Close it before creating another sketch."
                )
            conn.ensure_connected()
            part = conn.get_active_part()
            body = conn.get_active_part_body()
            plane_key = _normalise_plane(plane)
            offset_value = _require_finite(offset, "offset")
            contract = _plane_coordinate_contract(plane_key, offset_value)
            plane_ref = conn.create_offset_plane_reference(plane_key, offset_value)

            before = int(body.Sketches.Count)
            part.InWorkObject = body
            sketch = body.Sketches.Add(plane_ref)
            if str(sketch_name).strip():
                sketch.Name = str(sketch_name).strip()
            part.InWorkObject = sketch
            factory = sketch.OpenEdition()
            opened = True
            after = int(body.Sketches.Count)
            if after != before + 1:
                raise CATIAError(
                    f"Sketches.Count verification failed: before={before}, after={after}."
                )

            session.active_sketch = sketch
            session.active_factory = factory
            session.document_name = str(getattr(_active_document(conn), "Name", ""))
            return _success(
                {
                    "message": "Sketch created and opened.",
                    "sketch_name": getattr(sketch, "Name", ""),
                    "plane": plane_key,
                    "offset": offset_value,
                    "coordinate_contract": contract,
                    "sketch_count_before": before,
                    "sketch_count_after": after,
                    "creation_verified": True,
                    "edition_open": True,
                    "model_modified": True,
                    "document_save_required": True,
                }
            )
        except Exception as exc:
            cleanup_errors: list[str] = []
            if opened and sketch is not None:
                try:
                    sketch.CloseEdition()
                except Exception as close_exc:
                    cleanup_errors.append(f"CloseEdition: {close_exc}")
            cleanup = _delete_objects(conn, [sketch] if sketch is not None else [])
            session.reset()
            cleanup["errors"].extend(cleanup_errors)
            cleanup["verified"] = bool(cleanup["verified"] and not cleanup_errors)
            return _error(
                str(exc),
                status="error" if cleanup["verified"] else "partial_success",
                data={
                    "coordinate_contract": contract,
                    "cleanup": cleanup,
                    "model_modified": not cleanup["verified"],
                    "document_save_required": not cleanup["verified"],
                },
            )

    names.append("catia_create_sketch")

    @mcp.tool()
    def catia_close_sketch(update: bool = True) -> dict[str, Any]:
        """Close the active sketch and optionally update it."""
        try:
            session.ensure_open(conn)
            sketch = session.active_sketch
            sketch_name = str(getattr(sketch, "Name", ""))
            geometry_count = _geometry_count(sketch)
        except Exception as exc:
            return _error(str(exc))

        try:
            sketch.CloseEdition()
        except Exception as exc:
            return _error(
                f"CATIA could not close the sketch edition: {exc}",
                data={
                    "sketch_name": sketch_name,
                    "geometry_count": geometry_count,
                    "edition_open": True,
                    "update_requested": bool(update),
                    "update_succeeded": None,
                    "session_cleared": False,
                    "model_modified": True,
                    "document_save_required": True,
                },
                status="partial_success",
            )

        # CloseEdition succeeded.  Clear the process-local factory immediately
        # so a later update failure cannot leave a stale OpenEdition wrapper.
        session.reset()

        update_succeeded = True
        update_error = None
        if update:
            try:
                part = conn.get_active_part()
                part.UpdateObject(sketch)
            except Exception as exc:
                update_succeeded = False
                update_error = str(exc)

        try:
            conn.refresh_display()
        except Exception:
            pass

        data = {
            "message": "Sketch closed.",
            "sketch_name": sketch_name,
            "geometry_count": geometry_count,
            "edition_open": False,
            "update_requested": bool(update),
            "update_succeeded": update_succeeded,
            "update_error": update_error,
            "session_cleared": True,
            "model_modified": True,
            "document_save_required": True,
        }
        if update and not update_succeeded:
            return _error(
                "Sketch edition was closed, but CATIA could not update the sketch.",
                status="partial_success",
                data=data,
            )
        return _success(data)

    names.append("catia_close_sketch")

    @mcp.tool()
    def catia_sketch_line(x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """Draw a nonzero 2D line in the active sketch."""
        try:
            a = _require_finite(x1, "x1")
            b = _require_finite(y1, "y1")
            c = _require_finite(x2, "x2")
            d = _require_finite(y2, "y2")
            if math.isclose(a, c, abs_tol=1e-12) and math.isclose(b, d, abs_tol=1e-12):
                raise CATIAError("Line start and end points must differ.")
        except Exception as exc:
            return _error(str(exc))

        def create(created: list[Any]) -> dict[str, Any]:
            line = session.active_factory.CreateLine(a, b, c, d)
            created.append(line)
            return {"type": "line", "x1": a, "y1": b, "x2": c, "y2": d, "name": getattr(line, "Name", "")}

        return _create_geometry_transaction(conn, session, label="line", expected_count=1, creator=create)

    names.append("catia_sketch_line")

    @mcp.tool()
    def catia_sketch_rectangle(x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """Draw a closed rectangle by two distinct opposite corners."""
        try:
            a = _require_finite(x1, "x1")
            b = _require_finite(y1, "y1")
            c = _require_finite(x2, "x2")
            d = _require_finite(y2, "y2")
            if math.isclose(a, c, abs_tol=1e-12) or math.isclose(b, d, abs_tol=1e-12):
                raise CATIAError("Rectangle width and height must both be greater than zero.")
        except Exception as exc:
            return _error(str(exc))

        def create(created: list[Any]) -> dict[str, Any]:
            _create_rectangle_geometry(session.active_factory, a, b, c, d, created)
            return {"type": "rectangle", "x1": a, "y1": b, "x2": c, "y2": d, "width": abs(c-a), "height": abs(d-b)}

        return _create_geometry_transaction(conn, session, label="rectangle", expected_count=4, creator=create)

    names.append("catia_sketch_rectangle")

    @mcp.tool()
    def catia_sketch_centered_rectangle(width: float, height: float, center_x: float = 0.0, center_y: float = 0.0) -> dict[str, Any]:
        """Draw a closed rectangle centered at a local sketch point."""
        try:
            w = _require_positive(width, "width")
            h = _require_positive(height, "height")
            cx = _require_finite(center_x, "center_x")
            cy = _require_finite(center_y, "center_y")
        except Exception as exc:
            return _error(str(exc))
        x1, y1, x2, y2 = cx-w/2.0, cy-h/2.0, cx+w/2.0, cy+h/2.0

        def create(created: list[Any]) -> dict[str, Any]:
            _create_rectangle_geometry(session.active_factory, x1, y1, x2, y2, created)
            return {"type": "centered_rectangle", "width": w, "height": h, "center_x": cx, "center_y": cy, "x1": x1, "y1": y1, "x2": x2, "y2": y2}

        return _create_geometry_transaction(conn, session, label="centered_rectangle", expected_count=4, creator=create)

    names.append("catia_sketch_centered_rectangle")

    @mcp.tool()
    def catia_sketch_circle(radius: float, center_x: float = 0.0, center_y: float = 0.0) -> dict[str, Any]:
        """Draw a closed circle in active sketch."""
        try:
            r = _require_positive(radius, "Circle radius")
            cx = _require_finite(center_x, "Circle center_x")
            cy = _require_finite(center_y, "Circle center_y")
        except Exception as exc:
            return _error(str(exc))

        def create(created: list[Any]) -> dict[str, Any]:
            circle = session.active_factory.CreateClosedCircle(cx, cy, r)
            created.append(circle)
            return {"type": "circle", "center_x": cx, "center_y": cy, "radius": r, "name": getattr(circle, "Name", "")}

        return _create_geometry_transaction(conn, session, label="circle", expected_count=1, creator=create)

    names.append("catia_sketch_circle")

    @mcp.tool()
    def catia_sketch_point(x: float, y: float) -> dict[str, Any]:
        """Create a point in active sketch."""
        try:
            px = _require_finite(x, "x")
            py = _require_finite(y, "y")
        except Exception as exc:
            return _error(str(exc))

        def create(created: list[Any]) -> dict[str, Any]:
            point = session.active_factory.CreatePoint(px, py)
            created.append(point)
            return {"type": "point", "x": px, "y": py, "name": getattr(point, "Name", "")}

        return _create_geometry_transaction(conn, session, label="point", expected_count=1, creator=create)

    names.append("catia_sketch_point")

    @mcp.tool()
    def catia_sketch_rounded_rectangle(width: float, height: float, radius: float, center_x: float = 0.0, center_y: float = 0.0) -> dict[str, Any]:
        """Draw a rounded rectangle profile using verified lines and arcs."""
        try:
            w = _require_positive(width, "width")
            h = _require_positive(height, "height")
            r = _require_finite(radius, "radius")
            cx = _require_finite(center_x, "center_x")
            cy = _require_finite(center_y, "center_y")
            if r < 0.0:
                raise CATIAError("Radius cannot be negative.")
            r = min(r, min(w, h)/2.0)
        except Exception as exc:
            return _error(str(exc))

        if math.isclose(r, 0.0, abs_tol=1e-12):
            x1, y1, x2, y2 = cx-w/2.0, cy-h/2.0, cx+w/2.0, cy+h/2.0
            def create_rect(created: list[Any]) -> dict[str, Any]:
                _create_rectangle_geometry(session.active_factory, x1, y1, x2, y2, created)
                return {"type": "rounded_rectangle", "width": w, "height": h, "radius": 0.0, "center_x": cx, "center_y": cy}
            return _create_geometry_transaction(conn, session, label="rounded_rectangle", expected_count=4, creator=create_rect)

        left, right = cx-w/2.0, cx+w/2.0
        bottom, top = cy-h/2.0, cy+h/2.0

        def create(created: list[Any]) -> dict[str, Any]:
            f = session.active_factory
            created.append(_create_circle_arc(f, right-r, bottom+r, r, 270.0, 360.0))
            created.append(_create_circle_arc(f, right-r, top-r, r, 0.0, 90.0))
            created.append(_create_circle_arc(f, left+r, top-r, r, 90.0, 180.0))
            created.append(_create_circle_arc(f, left+r, bottom+r, r, 180.0, 270.0))
            for args in [
                (left+r, bottom, right-r, bottom),
                (right, bottom+r, right, top-r),
                (right-r, top, left+r, top),
                (left, top-r, left, bottom+r),
            ]:
                line = _create_line_if_nonzero(f, *args)
                if line is not None:
                    created.append(line)
            return {"type": "rounded_rectangle", "width": w, "height": h, "radius": r, "center_x": cx, "center_y": cy, "created_element_count": len(created)}

        expected = 4 + (2 if math.isclose(2*r, min(w,h), abs_tol=1e-12) and not math.isclose(w,h,abs_tol=1e-12) else 4)
        if math.isclose(w, h, abs_tol=1e-12) and math.isclose(r, w/2.0, abs_tol=1e-12):
            expected = 4
        return _create_geometry_transaction(conn, session, label="rounded_rectangle", expected_count=expected, creator=create)

    names.append("catia_sketch_rounded_rectangle")

    @mcp.tool()
    def catia_sketch_arc(center_x: float, center_y: float, radius: float, start_angle_deg: float, end_angle_deg: float) -> dict[str, Any]:
        """Draw a circular arc in active sketch."""
        try:
            r = _require_positive(radius, "Arc radius")
            cx = _require_finite(center_x, "Arc center_x")
            cy = _require_finite(center_y, "Arc center_y")
            start_param, end_param, sweep = _to_catia_arc_parameters(start_angle_deg, end_angle_deg)
        except Exception as exc:
            return _error(str(exc))

        def create(created: list[Any]) -> dict[str, Any]:
            arc = session.active_factory.CreateCircle(cx, cy, r, start_param, end_param)
            created.append(arc)
            return {"type": "arc", "center_x": cx, "center_y": cy, "radius": r, "start_angle_deg": float(start_angle_deg), "end_angle_deg": float(end_angle_deg), "sweep_angle_deg": sweep, "name": getattr(arc, "Name", "")}

        return _create_geometry_transaction(conn, session, label="arc", expected_count=1, creator=create)

    names.append("catia_sketch_arc")

    @mcp.tool()
    def catia_sketch_get_geometry() -> dict[str, Any]:
        """List geometry elements in the currently open sketch."""
        try:
            session.ensure_open(conn)
            geom = session.active_sketch.GeometricElements
            result: list[dict[str, Any]] = []
            for i in range(1, int(geom.Count)+1):
                item = geom.Item(i)
                result.append({"index": i, "name": getattr(item, "Name", ""), "type": getattr(item, "GeometricType", "")})
            return _success({"sketch_name": getattr(session.active_sketch, "Name", ""), "geometry_count": int(geom.Count), "elements": result, "model_modified": False, "document_save_required": False})
        except Exception as exc:
            return _error(str(exc))

    names.append("catia_sketch_get_geometry")

    return names


# ---------------------------------------------------------------------------
# Registry Entry Point (Required by registry.py)
# ---------------------------------------------------------------------------
