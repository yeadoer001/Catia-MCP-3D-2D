from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

logger = logging.getLogger("catia_mcp")

try:
    import pythoncom
    import win32com.client

    HAS_COM = True
except ImportError:
    pythoncom = None
    win32com = None
    HAS_COM = False


class CATIAError(RuntimeError):
    """Raised when CATIA COM automation fails."""


class CATIAConnection:
    """Lazy CATIA V5 COM connection manager.

    Importing this module never starts CATIA.
    CATIA starts only when connect() or ensure_connected() is called.
    """

    CATIA_PROGID = "CATIA.Application"

    def __init__(self) -> None:
        self.app: Any | None = None
        self._com_initialized = False

    @property
    def is_connected(self) -> bool:
        if self.app is None:
            return False

        try:
            _ = self.app.Caption
            return True
        except Exception:
            self.app = None
            return False

    def connect(self, visible: bool = True) -> Any:
        if not HAS_COM:
            raise CATIAError(
                "pywin32 is not installed or this is not a Windows Python environment. "
                "Install pywin32 with: pip install pywin32"
            )

        if self.is_connected:
            if visible:
                try:
                    self.app.Visible = True
                except Exception:
                    pass
            return self.app

        if not self._com_initialized:
            pythoncom.CoInitialize()
            self._com_initialized = True

        try:
            self.app = win32com.client.GetActiveObject(self.CATIA_PROGID)
            logger.info("Attached to running CATIA V5.")
        except Exception:
            logger.info("No running CATIA V5 found. Starting a new instance.")
            try:
                self.app = win32com.client.Dispatch(self.CATIA_PROGID)
            except Exception as exc:
                self.app = None
                raise CATIAError(
                    "Failed to launch CATIA V5. Make sure CATIA V5 is installed, "
                    "licensed, and registered as a COM server."
                ) from exc

        if visible:
            try:
                self.app.Visible = True
            except Exception:
                pass

        return self.app

    def disconnect(self) -> None:
        self.app = None

        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

            self._com_initialized = False

    def ensure_connected(self) -> Any:
        if not self.is_connected:
            return self.connect(visible=True)
        return self.app

    @property
    def catia(self) -> Any:
        return self.ensure_connected()

    @property
    def documents(self) -> Any:
        return self.catia.Documents

    @property
    def active_document(self) -> Any:
        app = self.ensure_connected()

        try:
            return app.ActiveDocument
        except Exception as exc:
            raise CATIAError("No active CATIA document. Create or open a document first.") from exc

    @property
    def active_editor(self) -> Any:
        app = self.ensure_connected()

        try:
            return app.ActiveEditor
        except Exception as exc:
            raise CATIAError("No active CATIA editor/viewer is available.") from exc

    @property
    def selection(self) -> Any:
        return self.active_document.Selection

    def get_active_part_document(self) -> Any:
        doc = self.active_document

        try:
            _ = doc.Part
            return doc
        except Exception as exc:
            raise CATIAError("Active document is not a CATPart document.") from exc

    def get_active_part(self) -> Any:
        return self.get_active_part_document().Part

    def get_active_product_document(self) -> Any:
        doc = self.active_document

        try:
            _ = doc.Product
            return doc
        except Exception as exc:
            raise CATIAError("Active document is not a CATProduct document.") from exc

    def get_active_product(self) -> Any:
        return self.get_active_product_document().Product

    def get_active_drawing_document(self) -> Any:
        doc = self.active_document
        name = safe_str(getattr(doc, "Name", "")).lower()
        doc_type = safe_str(getattr(doc, "Type", "")).lower()

        if "drawing" in doc_type or name.endswith(".catdrawing"):
            return doc

        raise CATIAError("Active document is not a CATDrawing document.")

    def get_active_part_body(self) -> Any:
        part = self.get_active_part()

        try:
            return part.MainBody
        except Exception as exc:
            raise CATIAError("Active CATPart has no usable MainBody.") from exc

    def get_origin_plane_object(self, plane: str) -> Any:
        part = self.get_active_part()
        origin = part.OriginElements
        key = plane.lower()

        if key == "xy":
            return origin.PlaneXY
        if key == "yz":
            return origin.PlaneYZ
        if key in ("zx", "xz"):
            return origin.PlaneZX

        raise CATIAError(f"Unsupported origin plane: {plane}. Use xy, yz, or zx.")

    def get_origin_plane_reference(self, plane: str) -> Any:
        part = self.get_active_part()
        plane_obj = self.get_origin_plane_object(plane)
        return part.CreateReferenceFromObject(plane_obj)

    def get_or_create_hybrid_body(self, name: str = "MCP_Construction") -> Any:
        part = self.get_active_part()
        hybrid_bodies = part.HybridBodies

        try:
            return hybrid_bodies.Item(name)
        except Exception:
            hb = hybrid_bodies.Add()
            hb.Name = name
            return hb

    def create_offset_plane_reference(self, plane: str, offset: float) -> Any:
        """Create an offset plane in a HybridBody and return a reference.

        This avoids the common CATIA problem where a transient plane reference
        cannot be used reliably for sketch creation.
        """
        part = self.get_active_part()

        if abs(float(offset)) < 1e-9:
            return self.get_origin_plane_reference(plane)

        base_ref = self.get_origin_plane_reference(plane)
        hsf = part.HybridShapeFactory
        try:
            import win32com.client
            hsf = win32com.client.Dispatch(hsf)
        except:
            pass
        hb = self.get_or_create_hybrid_body("MCP_Construction")

        offset_plane = hsf.AddNewPlaneOffset(base_ref, float(offset), False)
        offset_plane.Name = f"MCP_{plane.upper()}_Offset_{float(offset):g}"

        hb.AppendHybridShape(offset_plane)
        part.UpdateObject(offset_plane)

        return part.CreateReferenceFromObject(offset_plane)

    def create_reference_from_object(self, obj: Any) -> Any:
        part = self.get_active_part()
        return part.CreateReferenceFromObject(obj)

    def refresh_display(self) -> None:
        try:
            self.active_editor.ActiveViewer.Reframe()
            return
        except Exception:
            pass

        try:
            self.catia.ActiveWindow.ActiveViewer.Reframe()
        except Exception:
            pass

    def fit_all(self) -> None:
        self.refresh_display()

    def describe_document(self, doc: Any | None = None) -> dict[str, Any]:
        doc = doc or self.active_document

        info: dict[str, Any] = {
            "name": safe_str(getattr(doc, "Name", "")),
            "full_name": safe_str(getattr(doc, "FullName", "")),
            "type": safe_str(getattr(doc, "Type", "")),
            "saved": bool(getattr(doc, "Saved", False)) if hasattr(doc, "Saved") else False,
            "kind": "Unknown",
        }

        try:
            part = doc.Part
            info["kind"] = "CATPart"
            info["part_name"] = safe_str(getattr(part, "Name", ""))
            info["part_number"] = safe_str(getattr(part, "PartNumber", ""))
            return info
        except Exception:
            pass

        try:
            product = doc.Product
            info["kind"] = "CATProduct"
            info["product_name"] = safe_str(getattr(product, "Name", ""))
            info["part_number"] = safe_str(getattr(product, "PartNumber", ""))
            return info
        except Exception:
            pass

        if info["name"].lower().endswith(".catdrawing") or "drawing" in info["type"].lower():
            info["kind"] = "CATDrawing"

        return info

    def get_connection_info(self) -> dict[str, Any]:
        if not self.is_connected:
            return {
                "connected": False,
                "caption": "",
                "visible": False,
                "documents_count": 0,
                "active_document": None,
            }

        app = self.ensure_connected()

        documents_count = 0
        try:
            documents_count = int(app.Documents.Count)
        except Exception:
            pass

        active_document = None
        try:
            active_document = self.describe_document(app.ActiveDocument)
        except Exception:
            pass

        return {
            "connected": True,
            "caption": safe_str(getattr(app, "Caption", "")),
            "visible": bool(getattr(app, "Visible", False)),
            "documents_count": documents_count,
            "active_document": active_document,
        }


def safe_str(value: Any) -> str:
    try:
        return str(value)
    except Exception:
        return ""


def normalize_path(path: str) -> str:
    return str(Path(path).expanduser().resolve())