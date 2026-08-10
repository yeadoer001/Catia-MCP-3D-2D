from __future__ import annotations

from typing import Any

from mcp.server.fastmcp import FastMCP

from catia_mcp.connection import CATIAConnection
from catia_mcp.tools.document import DocumentTools
from catia_mcp.tools.sketcher import SketcherTools
from catia_mcp.tools.part_design import PartDesignTools
from catia_mcp.tools.product_modeling import ProductModelingTools
from catia_mcp.tools.drafting import DraftingTools
from catia_mcp.tools.export import ExportTools


INSTRUCTIONS = """
CATIA V5 MCP Server.

This server controls CATIA V5 through Windows COM automation.

Main capabilities:
- Start/connect CATIA V5.
- Create/open/save Part, Product, Drawing documents.
- Create sketches on origin or offset planes.
- Draw lines, rectangles, rounded rectangles, and circles.
- Create Pads, Pockets, circular bosses, and circular cuts.
- Create a parametric phone/iPhone-style concept model.
- Generate 3-view engineering drawings from CATPart/CATProduct files.
- Add title blocks, dimension notes, tolerance notes, and GD&T notes.
- Export CAD data, drawings, and screenshots.

Requirements:
- Windows.
- CATIA V5 installed and licensed.
- pywin32 installed.
"""

mcp = FastMCP("CATIA V5 MCP", instructions=INSTRUCTIONS)

conn = CATIAConnection()
doc_tools = DocumentTools(conn)
sketch_tools = SketcherTools(conn)
part_tools = PartDesignTools(conn)
product_tools = ProductModelingTools(conn)
drafting_tools = DraftingTools(conn)
export_tools = ExportTools(conn)


def _ok(data: Any) -> dict[str, Any]:
    return {"ok": True, "data": data}


def _fail(exc: Exception) -> dict[str, Any]:
    return {"ok": False, "error": str(exc)}


@mcp.tool()
def catia_start(visible: bool = True) -> dict[str, Any]:
    """Start CATIA if needed and return connection status."""
    try:
        return _ok(doc_tools.connect(visible=visible))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_status() -> dict[str, Any]:
    """Return CATIA connection status and active document information."""
    try:
        return _ok(conn.get_connection_info())
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_disconnect() -> dict[str, Any]:
    """Disconnect from CATIA COM object. Does not close CATIA."""
    try:
        return _ok(doc_tools.disconnect())
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_new_part(name: str = "MCP_Part") -> dict[str, Any]:
    """Create a new CATPart document."""
    try:
        return _ok(doc_tools.new_part(name))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_new_product(name: str = "MCP_Product") -> dict[str, Any]:
    """Create a new CATProduct document."""
    try:
        return _ok(doc_tools.new_product(name))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_open_document(file_path: str) -> dict[str, Any]:
    """Open a CATIA document from a file path."""
    try:
        return _ok(doc_tools.open_document(file_path))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_save_document(file_path: str = "") -> dict[str, Any]:
    """Save active document. If file_path is provided, Save As."""
    try:
        return _ok(doc_tools.save_document(file_path or None))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_close_document(save: bool = False) -> dict[str, Any]:
    """Close active document."""
    try:
        return _ok(doc_tools.close_document(save=save))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_list_documents() -> dict[str, Any]:
    """List open CATIA documents."""
    try:
        return _ok(doc_tools.list_documents())
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_get_active_document_info() -> dict[str, Any]:
    """Get active CATIA document info."""
    try:
        return _ok(doc_tools.active_document_info())
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_create_sketch(
    plane: str = "xy",
    sketch_name: str = "",
    offset: float = 0.0,
) -> dict[str, Any]:
    """Create and open a sketch on xy/yz/zx plane, optionally offset in mm."""
    try:
        return _ok(
            sketch_tools.create_sketch(
                plane=plane,
                sketch_name=sketch_name or None,
                offset=offset,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_close_sketch(update: bool = True) -> dict[str, Any]:
    """Close active sketch."""
    try:
        return _ok(sketch_tools.close_sketch(update=update))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_sketch_line(x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
    """Draw a 2D line in active sketch."""
    try:
        return _ok(sketch_tools.line(x1, y1, x2, y2))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_sketch_rectangle(
    x1: float,
    y1: float,
    x2: float,
    y2: float,
) -> dict[str, Any]:
    """Draw a rectangle in active sketch by two opposite corners."""
    try:
        return _ok(sketch_tools.rectangle(x1, y1, x2, y2))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_sketch_centered_rectangle(
    width: float,
    height: float,
    center_x: float = 0.0,
    center_y: float = 0.0,
) -> dict[str, Any]:
    """Draw a centered rectangle in active sketch."""
    try:
        return _ok(
            sketch_tools.centered_rectangle(
                width=width,
                height=height,
                center_x=center_x,
                center_y=center_y,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_sketch_rounded_rectangle(
    width: float,
    height: float,
    radius: float,
    center_x: float = 0.0,
    center_y: float = 0.0,
) -> dict[str, Any]:
    """Draw a rounded rectangle in active sketch."""
    try:
        return _ok(
            sketch_tools.rounded_rectangle(
                width=width,
                height=height,
                radius=radius,
                center_x=center_x,
                center_y=center_y,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_sketch_circle(
    radius: float,
    center_x: float = 0.0,
    center_y: float = 0.0,
) -> dict[str, Any]:
    """Draw a circle in active sketch."""
    try:
        return _ok(
            sketch_tools.circle(
                radius=radius,
                center_x=center_x,
                center_y=center_y,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_pad(
    height: float,
    sketch_name: str = "",
    direction: str = "normal",
    symmetric: bool = False,
    feature_name: str = "",
) -> dict[str, Any]:
    """Create a Pad from a sketch."""
    try:
        return _ok(
            part_tools.pad(
                height=height,
                sketch_name=sketch_name or None,
                direction=direction,
                symmetric=symmetric,
                feature_name=feature_name or None,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_pocket(
    depth: float,
    sketch_name: str = "",
    direction: str = "normal",
    feature_name: str = "",
) -> dict[str, Any]:
    """Create a Pocket from a sketch."""
    try:
        return _ok(
            part_tools.pocket(
                depth=depth,
                sketch_name=sketch_name or None,
                direction=direction,
                feature_name=feature_name or None,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_circular_pad(
    radius: float,
    height: float,
    center_x: float = 0.0,
    center_y: float = 0.0,
    plane: str = "xy",
    offset: float = 0.0,
    sketch_name: str = "",
    feature_name: str = "",
) -> dict[str, Any]:
    """Create a circular boss by circle sketch + Pad."""
    try:
        return _ok(
            part_tools.circular_pad(
                radius=radius,
                height=height,
                center_x=center_x,
                center_y=center_y,
                plane=plane,
                offset=offset,
                sketch_name=sketch_name or None,
                feature_name=feature_name or None,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_cut_circular_hole(
    diameter: float,
    depth: float,
    center_x: float = 0.0,
    center_y: float = 0.0,
    plane: str = "xy",
    offset: float = 0.0,
    direction: str = "normal",
    sketch_name: str = "",
    feature_name: str = "",
) -> dict[str, Any]:
    """Cut a circular hole using circle sketch + Pocket."""
    try:
        return _ok(
            part_tools.cut_circular_hole(
                diameter=diameter,
                depth=depth,
                center_x=center_x,
                center_y=center_y,
                plane=plane,
                offset=offset,
                direction=direction,
                sketch_name=sketch_name or None,
                feature_name=feature_name or None,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_list_features() -> dict[str, Any]:
    """List features in active PartBody."""
    try:
        return _ok(part_tools.list_features())
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_create_phone_mockup(
    model_name: str = "Phone_Concept",
    width: float = 77.6,
    height: float = 163.0,
    thickness: float = 8.3,
    corner_radius: float = 18.0,
    camera_island_width: float = 38.0,
    camera_island_height: float = 38.0,
    camera_island_radius: float = 8.0,
    camera_island_offset_x: float = -18.0,
    camera_island_offset_y: float = 50.0,
    camera_island_height_z: float = 1.8,
    lens_diameter: float = 13.5,
    lens_height: float = 1.2,
) -> dict[str, Any]:
    """Create a simplified iPhone-style 3D concept model."""
    try:
        return _ok(
            product_tools.create_phone_mockup(
                model_name=model_name,
                width=width,
                height=height,
                thickness=thickness,
                corner_radius=corner_radius,
                camera_island_width=camera_island_width,
                camera_island_height=camera_island_height,
                camera_island_radius=camera_island_radius,
                camera_island_offset_x=camera_island_offset_x,
                camera_island_offset_y=camera_island_offset_y,
                camera_island_height_z=camera_island_height_z,
                lens_diameter=lens_diameter,
                lens_height=lens_height,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_create_3view_drawing_from_file(
    model_path: str,
    paper_size: str = "A3",
    orientation: str = "landscape",
    scale: float = 1.0,
    drawing_title: str = "",
    part_number: str = "",
    material: str = "",
    general_tolerance: str = "ISO 2768-mK unless otherwise specified",
    generate_dimensions: bool = False,
    dimension_notes: list[str] | None = None,
    tolerance_notes: list[str] | None = None,
    gdt_notes: list[str] | None = None,
    output_path: str = "",
    export_pdf_path: str = "",
) -> dict[str, Any]:
    """Create a 3-view engineering drawing from a CATPart/CATProduct file."""
    try:
        return _ok(
            drafting_tools.create_three_view_drawing_from_file(
                model_path=model_path,
                paper_size=paper_size,
                orientation=orientation,
                scale=scale,
                drawing_title=drawing_title,
                part_number=part_number,
                material=material,
                general_tolerance=general_tolerance,
                generate_dimensions=generate_dimensions,
                dimension_notes=dimension_notes or [],
                tolerance_notes=tolerance_notes or [],
                gdt_notes=gdt_notes or [],
                output_path=output_path,
                export_pdf_path=export_pdf_path,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_generate_drawing_dimensions() -> dict[str, Any]:
    """Ask CATIA to auto-generate dimensions on active drawing sheet."""
    try:
        return _ok(drafting_tools.generate_dimensions_on_active_sheet())
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_add_drawing_text(
    text: str,
    x: float = 20.0,
    y: float = 20.0,
    font_size: float = 3.5,
    use_background_view: bool = False,
) -> dict[str, Any]:
    """Add drawing text to active CATDrawing."""
    try:
        return _ok(
            drafting_tools.add_drawing_text(
                text=text,
                x=x,
                y=y,
                font_size=font_size,
                use_background_view=use_background_view,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_add_gdt_note(
    feature: str,
    control: str,
    tolerance: str,
    datum: str = "",
    x: float = 20.0,
    y: float = 70.0,
) -> dict[str, Any]:
    """Add a GD&T note, such as flatness, perpendicularity, or position."""
    try:
        return _ok(
            drafting_tools.add_gdt_note(
                feature=feature,
                control=control,
                tolerance=tolerance,
                datum=datum,
                x=x,
                y=y,
            )
        )
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_export_active_drawing(path: str, format_name: str = "pdf") -> dict[str, Any]:
    """Export active CATDrawing to PDF/DWG/DXF."""
    try:
        return _ok(drafting_tools.export_active_drawing(path, format_name))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_save_active_drawing_as(path: str) -> dict[str, Any]:
    """Save active CATDrawing as a CATDrawing file."""
    try:
        return _ok(drafting_tools.save_active_drawing_as(path))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_export(file_path: str, format_name: str = "") -> dict[str, Any]:
    """Export active CATIA document."""
    try:
        return _ok(export_tools.export(file_path, format_name or None))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_screenshot(file_path: str) -> dict[str, Any]:
    """Save screenshot of active CATIA viewer."""
    try:
        return _ok(export_tools.screenshot(file_path))
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_fit_all() -> dict[str, Any]:
    """Fit all geometry in active CATIA viewer."""
    try:
        return _ok(export_tools.fit_all())
    except Exception as exc:
        return _fail(exc)


@mcp.tool()
def catia_set_view(view: str) -> dict[str, Any]:
    """Set active 3D view: front, back, top, bottom, left, right, isometric."""
    try:
        return _ok(export_tools.set_view(view))
    except Exception as exc:
        return _fail(exc)


TOOL_NAMES = [
    "catia_start",
    "catia_status",
    "catia_disconnect",
    "catia_new_part",
    "catia_new_product",
    "catia_open_document",
    "catia_save_document",
    "catia_close_document",
    "catia_list_documents",
    "catia_get_active_document_info",
    "catia_create_sketch",
    "catia_close_sketch",
    "catia_sketch_line",
    "catia_sketch_rectangle",
    "catia_sketch_centered_rectangle",
    "catia_sketch_rounded_rectangle",
    "catia_sketch_circle",
    "catia_pad",
    "catia_pocket",
    "catia_circular_pad",
    "catia_cut_circular_hole",
    "catia_list_features",
    "catia_create_phone_mockup",
    "catia_create_3view_drawing_from_file",
    "catia_generate_drawing_dimensions",
    "catia_add_drawing_text",
    "catia_add_gdt_note",
    "catia_export_active_drawing",
    "catia_save_active_drawing_as",
    "catia_export",
    "catia_screenshot",
    "catia_fit_all",
    "catia_set_view",
]


class CATIAMCPServer:
    """Small wrapper for tests and CLI startup."""

    def __init__(self) -> None:
        self.mcp = mcp
        self._tool_router = {name: True for name in TOOL_NAMES}

    def run(self) -> None:
        self.mcp.run()


def main() -> None:
    CATIAMCPServer().run()


if __name__ == "__main__":
    main()