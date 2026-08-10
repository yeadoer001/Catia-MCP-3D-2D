from __future__ import annotations

from typing import Any


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
        """Create a 3D point in active CATPart."""
        try:
            part = conn.get_active_part()
            hsf = part.HybridShapeFactory
            hb = conn.get_or_create_hybrid_body(geometrical_set)

            point = hsf.AddNewPointCoord(float(x), float(y), float(z))
            point.Name = name
            hb.AppendHybridShape(point)

            part.UpdateObject(point)
            conn.refresh_display()

            return {
                "ok": True,
                "data": {
                    "name": getattr(point, "Name", ""),
                    "x": x,
                    "y": y,
                    "z": z,
                    "geometrical_set": geometrical_set,
                },
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    names.append("catia_create_point")

    @mcp.tool()
    def catia_create_offset_plane(
        base_plane: str = "xy",
        offset: float = 0.0,
        name: str = "",
        geometrical_set: str = "MCP_Construction",
    ) -> dict[str, Any]:
        """Create an offset plane from xy/yz/zx origin plane."""
        try:
            part = conn.get_active_part()

            if abs(float(offset)) < 1e-9:
                ref = conn.get_origin_plane_reference(base_plane)
                return {
                    "ok": True,
                    "data": {
                        "name": base_plane.upper(),
                        "base_plane": base_plane,
                        "offset": offset,
                        "note": "Offset is zero; origin plane reference used.",
                    },
                }

            base_ref = conn.get_origin_plane_reference(base_plane)
            hsf = part.HybridShapeFactory
            hb = conn.get_or_create_hybrid_body(geometrical_set)

            plane = hsf.AddNewPlaneOffset(base_ref, float(offset), False)
            plane.Name = name or f"MCP_{base_plane.upper()}_Offset_{float(offset):g}"
            hb.AppendHybridShape(plane)

            part.UpdateObject(plane)
            conn.refresh_display()

            return {
                "ok": True,
                "data": {
                    "name": getattr(plane, "Name", ""),
                    "base_plane": base_plane,
                    "offset": offset,
                    "geometrical_set": geometrical_set,
                },
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    names.append("catia_create_offset_plane")

    return names