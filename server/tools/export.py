import os
from pathlib import Path

from core.connection import get_connection
from core.models import ToolResult

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent


# Export filter constants (CorelDRAW X6/16.1 COM filter IDs)
_CDR_DXF = 1296
_CDR_AI = 1305
_CDR_SVG = 1345
_CDR_PNG = 802
_CDR_JPEG = 774

_CDR_ALL_PAGES = 0
_CDR_CURRENT_PAGE = 1
_PDF_CURRENT_PAGE = 1
_CDR_RGB_COLOR_IMAGE = 4
_CDR_NORMAL_ANTIALIASING = 1
_CDR_COMPRESSION_NONE = 0
_CDR_MILLIMETER = 3

_DXF_VERSION = {"R12": 0, "R14": 1, "R2000": 3, "R2004": 4}


def resolve_export_path(path: str) -> str:
    """Resolve relative paths under the project root to avoid permission errors."""
    p = Path(path)
    if not p.is_absolute():
        p = _PROJECT_ROOT / p
    return str(p)


def _ensure_dir(path: str) -> str:
    """Create the destination directory and return an absolute path."""
    abs_path = resolve_export_path(path)
    dirpath = os.path.dirname(abs_path)
    if dirpath:
        os.makedirs(dirpath, exist_ok=True)
    return abs_path



def _page_size_mm(doc):
    """Return the active page size in millimetres."""
    try:
        doc.Unit = _CDR_MILLIMETER
        page = doc.ActivePage
        return page.SizeWidth, page.SizeHeight
    except Exception:
        return 0, 0


def _prepare_current_page_export(doc) -> int:
    """Read export metadata without changing selection or layer state.

    Preview/export tools are read-only with respect to the CDR document. Hidden
    and non-printable layers therefore remain exactly as the designer set them.
    """
    doc.Unit = _CDR_MILLIMETER
    page = doc.ActivePage
    shapes_count = 0
    try:
        shapes_count = page.Shapes.Count
    except Exception:
        pass
    try:
        page.Activate()
    except Exception:
        pass
    return shapes_count


def _finish_export_filter(export_filter) -> None:
    if export_filter is not None:
        export_filter.Finish()


def _file_size(path: str) -> int:
    try:
        return os.path.getsize(path)
    except Exception:
        return 0


def _verify_exported(path: str) -> int:
    size = _file_size(path)
    if size <= 0:
        raise RuntimeError(f"Export failed; output file was not created or is empty: {path}")
    return size


def _page_export_area(app, doc):
    page_w_mm, page_h_mm = _page_size_mm(doc)
    if page_w_mm <= 0 or page_h_mm <= 0:
        return None
    try:
        return app.CreateRect(0, 0, page_w_mm, page_h_mm)
    except Exception:
        return None


def _export_geometry(app, doc, scope="page", padding_mm=0):
    """Choose page or visible/printable artwork bounds without editing the page.

    Artwork bounds include objects outside the page. Padding is explicit and also
    provides room for outlines extending beyond the geometric object bounds.
    """
    if scope not in {"page", "artwork"}:
        raise ValueError("scope must be 'page' or 'artwork'")
    if padding_mm < 0:
        raise ValueError("padding_mm cannot be negative")
    if scope == "page":
        width, height = _page_size_mm(doc)
        return _page_export_area(app, doc), width, height
    doc.Unit = _CDR_MILLIMETER
    boxes = []
    for shape in doc.ActivePage.Shapes:
        layer = shape.Layer
        if not layer.Visible or not layer.Printable:
            continue
        boxes.append((float(shape.LeftX), float(shape.BottomY),
                      float(shape.RightX), float(shape.TopY)))
    if not boxes:
        raise ValueError("No visible printable artwork exists on the active page")
    left = min(b[0] for b in boxes) - padding_mm
    bottom = min(b[1] for b in boxes) - padding_mm
    right = max(b[2] for b in boxes) + padding_mm
    top = max(b[3] for b in boxes) + padding_mm
    width, height = right - left, top - bottom
    if width <= 0 or height <= 0:
        raise ValueError("Artwork export bounds must have positive width and height")
    return app.CreateRect(left, bottom, width, height), width, height


def _workspace_export(conn, path, exporter, padding_mm, **options):
    """Export visible printable artwork plus visible Desktop objects via a copy.

    X8 can move external objects back to its non-printable Desktop layer on save.
    Copy only the requested workspace into a disposable document; never enable
    printing on the source Desktop layer or resize/reposition the source artwork.
    """
    if padding_mm < 0:
        raise ValueError("padding_mm cannot be negative")
    app = conn.app
    source = app.ActiveDocument
    source_page = source.ActivePage
    old_unit = source.Unit
    source.Unit = _CDR_MILLIMETER
    temporary = None
    try:
        selected = []
        for i in range(1, source_page.Shapes.Count + 1):
            shape = source_page.Shapes.Item(i)
            layer = shape.Layer
            if layer.Visible and (layer.Printable or layer.IsDesktopLayer):
                selected.append(shape)
        if not selected:
            raise ValueError("No visible workspace artwork exists")
        boxes = [(float(s.LeftX), float(s.BottomY), float(s.RightX), float(s.TopY)) for s in selected]
        left, bottom = min(b[0] for b in boxes), min(b[1] for b in boxes)
        width = max(b[2] for b in boxes) - left + padding_mm * 2
        height = max(b[3] for b in boxes) - bottom + padding_mm * 2
        if width <= 0 or height <= 0:
            raise ValueError("Workspace export bounds must be positive")
        temporary = app.CreateDocument()
        temporary.Activate()
        temporary.Unit = _CDR_MILLIMETER
        temporary.WorldScale = source.WorldScale
        temporary.ActivePage.SetSize(width, height)
        layer = temporary.ActivePage.CreateLayer("Temporary workspace export")
        layer.Activate()
        for shape, box in zip(selected, boxes):
            source.Activate()
            source_page.Activate()
            shape.Copy()  # X8 cannot copy one mixed local/Desktop ShapeRange.
            temporary.Activate()
            layer.Activate()
            page = temporary.ActivePage
            before = {str(page.Shapes.Item(i).StaticID) for i in range(1, page.Shapes.Count + 1)}
            layer.Paste()
            added = [page.Shapes.Item(i) for i in range(1, page.Shapes.Count + 1)
                     if str(page.Shapes.Item(i).StaticID) not in before]
            if len(added) != 1:
                raise RuntimeError("Workspace root object did not paste as one object")
            pasted = added[0]
            pasted.LeftX = box[0] - left + padding_mm
            pasted.TopY = box[3] - bottom + padding_mm
            pasted.MoveToLayer(layer)
        layer.Printable = True
        temporary.Activate()
        temporary.ActivePage.Activate()
        temporary.Unit = _CDR_MILLIMETER
        if app.ActiveDocument.Name != temporary.Name:
            raise RuntimeError("Unable to activate the temporary workspace export document")
        result = exporter(path, scope="page", **options)
        if not result.success:
            raise RuntimeError(result.error or "Workspace export failed")
        return {**(result.data or {}), "scope": "workspace", "source_object_count": len(selected),
                "workspace_width_mm": width, "workspace_height_mm": height}
    finally:
        if temporary is not None:
            temporary.Dirty = False
            temporary.Close()
        source.Activate()
        source_page.Activate()
        source.Unit = old_unit


def export_pdf(
    path: str,
    color_profile: str = "ISO_Coated_v2",
    bleed: float = 3.0,
    crop_marks: bool = True,
    multi_page: bool = False,
    pdfx_version: str = "",
) -> ToolResult:
    """Export a production PDF with color, bleed, crop-mark, multipage, and PDF/X options."""
    conn = get_connection()
    if not conn.status.connected:
        return ToolResult.fail("CorelDRAW is not connected")
    output_path = _ensure_dir(path)

    def _export():
        doc = conn.app.ActiveDocument
        if not doc:
            raise RuntimeError("No document is open in CorelDRAW")
        shapes_count = _prepare_current_page_export(doc)
        try:
            doc.PDFSettings.ColorMode = 1  # cdrPDFCMYK
            doc.PDFSettings.BleedingLimit = bleed
            doc.PDFSettings.PrintCropMarks = crop_marks
            doc.PDFSettings.MultiPage = multi_page
            doc.PDFSettings.PublishRange = _CDR_ALL_PAGES if multi_page else _PDF_CURRENT_PAGE
            doc.PDFSettings.PageRange = "" if multi_page else str(doc.ActivePage.Index)
            doc.PDFSettings.TextAsCurves = True
            # PDF/X support
            if pdfx_version:
                try:
                    doc.PDFSettings.OutputAsPDFX = True
                except Exception:
                    pass
                try:
                    doc.PDFSettings.PDFXVersion = pdfx_version
                except Exception:
                    try:
                        pdfx_map = {"PDFX1a": 0, "PDFX3": 1, "PDFX4": 2}
                        doc.PDFSettings.PDFXVersion = pdfx_map.get(pdfx_version, 0)
                    except Exception:
                        pass
        except Exception:
            pass
        # Try known ICC profile property names across CorelDRAW versions.
        if color_profile:
            for attr in ("ColorProfileName", "ColorProfile", "OutputColorProfile", "ICCProfileName"):
                try:
                    setattr(doc.PDFSettings, attr, color_profile)
                    break
                except Exception:
                    continue
        doc.PublishToPDF(output_path)
        file_size = _verify_exported(output_path)
        return {
            "path": output_path,
            "color_profile": color_profile,
            "bleed": bleed,
            "crop_marks": crop_marks,
            "pdfx_version": pdfx_version,
            "shapes_count": shapes_count,
            "file_size": file_size,
        }

    result = conn.safe_call(_export, timeout=300.0)
    if result["success"]:
        return ToolResult.ok(f"PDF exported: {path}", **result["result"])
    return ToolResult.fail(result.get("error", "PDF export failed"))


def export_dxf(path: str, version: str = "R14", layer_filter: str = "", export_hidden: bool = False) -> ToolResult:
    """Export DXF for laser cutting/engraving with version and layer options."""
    conn = get_connection()
    if not conn.status.connected:
        return ToolResult.fail("CorelDRAW is not connected")
    output_path = _ensure_dir(path)

    dxf_ver = _DXF_VERSION.get(version, 1)

    def _export():
        doc = conn.app.ActiveDocument
        if not doc:
            raise RuntimeError("No document is open in CorelDRAW")
        shapes_count = _prepare_current_page_export(doc)
        # Try ExportEx (newer CorelDRAW). X6 often rejects this with COM object error;
        # fall back to basic Export which works silently in automation mode.
        try:
            expflt = doc.ExportEx(output_path, _CDR_DXF, _CDR_CURRENT_PAGE, None, None)
            for attr, val in [("BitmapType", 0), ("TextAsCurves", True),
                              ("Version", dxf_ver), ("Units", 3), ("FillUnmapped", True)]:
                try:
                    setattr(expflt, attr, val)
                except Exception:
                    pass
            if layer_filter:
                filter_layers = [l.strip() for l in layer_filter.split(",") if l.strip()]
                if filter_layers:
                    try:
                        expflt.LayerFilter = ",".join(filter_layers)
                    except Exception:
                        pass
            _finish_export_filter(expflt)
        except Exception:
            doc.Export(output_path, _CDR_DXF, _CDR_CURRENT_PAGE, None, None)
        file_size = _verify_exported(output_path)
        return {
            "path": output_path,
            "version": version,
            "text_as_curves": True,
            "shapes_count": shapes_count,
            "file_size": file_size,
        }

    result = conn.safe_call(_export, timeout=300.0)
    if result["success"]:
        return ToolResult.ok(f"DXF exported: {path} (v{version})", **result["result"])
    return ToolResult.fail(result.get("error", "DXF export failed"))


def export_ai(path: str) -> ToolResult:
    """Export Adobe Illustrator format."""
    conn = get_connection()
    if not conn.status.connected:
        return ToolResult.fail("CorelDRAW is not connected")
    output_path = _ensure_dir(path)

    def _export():
        doc = conn.app.ActiveDocument
        if not doc:
            raise RuntimeError("No document is open in CorelDRAW")
        shapes_count = _prepare_current_page_export(doc)
        doc.Export(output_path, _CDR_AI, _CDR_CURRENT_PAGE, None, None)
        file_size = _verify_exported(output_path)
        return {"path": output_path, "format": "ai", "shapes_count": shapes_count, "file_size": file_size}

    result = conn.safe_call(_export, timeout=300.0)
    if result["success"]:
        return ToolResult.ok(f"AI exported: {path}", **result["result"])
    return ToolResult.fail(result.get("error", "AI export failed"))


def export_svg(path: str) -> ToolResult:
    """Export SVG vector format."""
    conn = get_connection()
    if not conn.status.connected:
        return ToolResult.fail("CorelDRAW is not connected")
    output_path = _ensure_dir(path)

    def _export():
        doc = conn.app.ActiveDocument
        if not doc:
            raise RuntimeError("No document is open in CorelDRAW")
        shapes_count = _prepare_current_page_export(doc)
        doc.Export(output_path, _CDR_SVG, _CDR_CURRENT_PAGE, None, None)
        file_size = _verify_exported(output_path)
        return {"path": output_path, "format": "svg", "shapes_count": shapes_count, "file_size": file_size}

    result = conn.safe_call(_export, timeout=300.0)
    if result["success"]:
        return ToolResult.ok(f"SVG exported: {path}", **result["result"])
    return ToolResult.fail(result.get("error", "SVG export failed"))


def export_png(path: str, dpi: int = 300, width: int = 0, background_transparent: bool = False,
               scope: str = "page", padding_mm: float = 0) -> ToolResult:
    """Export PNG with page, artwork (printable), or workspace (also Desktop) scope."""
    conn = get_connection()
    if not conn.status.connected:
        return ToolResult.fail("CorelDRAW is not connected")
    output_path = _ensure_dir(path)

    def _export():
        if scope == "workspace":
            return _workspace_export(conn, output_path, export_png, padding_mm,
                                     dpi=dpi, width=width, background_transparent=background_transparent)
        doc = conn.app.ActiveDocument
        if not doc:
            raise RuntimeError("No document is open in CorelDRAW")
        shapes_count = _prepare_current_page_export(doc)
        export_area, page_w_mm, page_h_mm = _export_geometry(conn.app, doc, scope, padding_mm)
        if width > 0 and page_w_mm > 0:
            export_dpi = int(width / (page_w_mm / 25.4))
        else:
            export_dpi = dpi
        export_dpi = max(72, export_dpi)
        if page_w_mm > 0 and page_h_mm > 0:
            pixel_w = max(1, int(page_w_mm / 25.4 * export_dpi))
            pixel_h = max(1, int(page_h_mm / 25.4 * export_dpi))
        else:
            pixel_w = 0
            pixel_h = 0

        exported = False
        try:
            expflt = doc.ExportBitmap(
                output_path,
                _CDR_PNG,
                _CDR_CURRENT_PAGE,
                _CDR_RGB_COLOR_IMAGE,
                pixel_w,
                pixel_h,
                export_dpi,
                export_dpi,
                _CDR_NORMAL_ANTIALIASING,
                False,
                background_transparent,
                True,
                False,
                _CDR_COMPRESSION_NONE,
                None,
                export_area,
            )
            _finish_export_filter(expflt)
            exported = True
        except Exception:
            if scope != "page":
                raise  # Never silently fall back to a cropped page export.

        if not exported:
            doc.Export(output_path, _CDR_PNG, _CDR_CURRENT_PAGE, None, None)

        file_size = _verify_exported(output_path)
        return {
            "path": output_path,
            "dpi": export_dpi,
            "width_px": pixel_w,
            "height_px": pixel_h,
            "format": "png",
            "scope": scope,
            "shapes_count": shapes_count,
            "file_size": file_size,
        }

    result = conn.safe_call(_export, timeout=300.0)
    if result["success"]:
        return ToolResult.ok(f"PNG exported: {path}", **result["result"])
    return ToolResult.fail(result.get("error", "PNG export failed"))


def export_jpeg(path: str, quality: int = 85, dpi: int = 300,
                scope: str = "page", padding_mm: float = 0) -> ToolResult:
    """Export JPEG with page, artwork (printable), or workspace (also Desktop) scope."""
    conn = get_connection()
    if not conn.status.connected:
        return ToolResult.fail("CorelDRAW is not connected")
    output_path = _ensure_dir(path)

    def _export():
        if scope == "workspace":
            return _workspace_export(conn, output_path, export_jpeg, padding_mm, quality=quality, dpi=dpi)
        doc = conn.app.ActiveDocument
        if not doc:
            raise RuntimeError("No document is open in CorelDRAW")
        shapes_count = _prepare_current_page_export(doc)
        export_area, page_w_mm, page_h_mm = _export_geometry(conn.app, doc, scope, padding_mm)
        if page_w_mm > 0 and page_h_mm > 0:
            pixel_w = max(1, int(page_w_mm / 25.4 * dpi))
            pixel_h = max(1, int(page_h_mm / 25.4 * dpi))
        else:
            pixel_w = 0
            pixel_h = 0

        exported = False
        try:
            expflt = doc.ExportBitmap(
                output_path,
                _CDR_JPEG,
                _CDR_CURRENT_PAGE,
                _CDR_RGB_COLOR_IMAGE,
                pixel_w,
                pixel_h,
                dpi,
                dpi,
                _CDR_NORMAL_ANTIALIASING,
                False,
                False,
                True,
                False,
                _CDR_COMPRESSION_NONE,
                None,
                export_area,
            )
            # Set JPEG quality (0-100)
            try:
                expflt.JPEGQuality = quality
            except Exception:
                pass
            _finish_export_filter(expflt)
            exported = True
        except Exception:
            if scope != "page":
                raise

        if not exported:
            doc.Export(output_path, _CDR_JPEG, _CDR_CURRENT_PAGE, None, None)

        file_size = _verify_exported(output_path)
        return {
            "path": output_path,
            "quality": quality,
            "dpi": dpi,
            "width_px": pixel_w,
            "height_px": pixel_h,
            "format": "jpeg",
            "scope": scope,
            "shapes_count": shapes_count,
            "file_size": file_size,
        }

    result = conn.safe_call(_export, timeout=300.0)
    if result["success"]:
        return ToolResult.ok(f"JPEG exported: {path}", **result["result"])
    return ToolResult.fail(result.get("error", "JPEG export failed"))


def export_preview_png(path: str, width: int = 800, width_px: int = 0,
                       scope: str = "page", padding_mm: float = 0) -> ToolResult:
    """Preview page, printable artwork, or workspace including visible Desktop objects."""
    if width_px > 0:
        width = width_px
    conn = get_connection()
    if not conn.status.connected:
        return ToolResult.fail("CorelDRAW is not connected")
    output_path = _ensure_dir(path)

    def _export():
        if scope == "workspace":
            return _workspace_export(conn, output_path, export_preview_png, padding_mm, width=width)
        doc = conn.app.ActiveDocument
        if not doc:
            raise RuntimeError("No document is open in CorelDRAW")
        shapes_count = _prepare_current_page_export(doc)
        export_area, page_w_mm, page_h_mm = _export_geometry(conn.app, doc, scope, padding_mm)
        if page_w_mm > 0:
            export_dpi = int(width / (page_w_mm / 25.4))
        else:
            export_dpi = 72
        export_dpi = max(72, min(export_dpi, 300))
        if page_w_mm > 0 and page_h_mm > 0:
            pixel_w = max(1, int(page_w_mm / 25.4 * export_dpi))
            pixel_h = max(1, int(page_h_mm / 25.4 * export_dpi))
        else:
            pixel_w = width
            pixel_h = 0

        exported = False
        try:
            expflt = doc.ExportBitmap(
                output_path,
                _CDR_PNG,
                _CDR_CURRENT_PAGE,
                _CDR_RGB_COLOR_IMAGE,
                pixel_w,
                pixel_h,
                export_dpi,
                export_dpi,
                _CDR_NORMAL_ANTIALIASING,
                False,
                False,
                True,
                False,
                _CDR_COMPRESSION_NONE,
                None,
                export_area,
            )
            _finish_export_filter(expflt)
            exported = True
        except Exception:
            if scope != "page":
                raise

        if not exported:
            doc.Export(output_path, _CDR_PNG, _CDR_CURRENT_PAGE, None, None)

        file_size = _verify_exported(output_path)
        return {
            "path": output_path,
            "width_px": width,
            "scope": scope,
            "actual_width_px": pixel_w,
            "actual_height_px": pixel_h,
            "dpi": export_dpi,
            "shapes_count": shapes_count,
            "file_size": file_size,
        }

    result = conn.safe_call(_export, timeout=300.0)
    if result["success"]:
        return ToolResult.ok(f"Preview exported: {path}", **result["result"])
    return ToolResult.fail(result.get("error", "Preview export failed"))


def batch_export(pages: str, format: str, output_dir: str) -> ToolResult:
    """Export selected pages; pages accepts lists, ranges, or all."""
    conn = get_connection()
    if not conn.status.connected:
        return ToolResult.fail("CorelDRAW is not connected")

    format = format.lower()

    def _batch():
        doc = conn.app.ActiveDocument
        if not doc:
            raise RuntimeError("No document is open in CorelDRAW")
        _ensure_dir(os.path.join(output_dir, "dummy.txt"))

        all_pages = list(doc.Pages)
        if pages.strip().lower() == "all":
            selected_pages = all_pages
        elif "-" in pages:
            start, end = [int(p.strip()) for p in pages.split("-")]
            selected_pages = [p for p in all_pages if start <= p.Index <= end]
        else:
            indices = [int(p.strip()) for p in pages.split(",") if p.strip()]
            selected_pages = []
            for p in all_pages:
                if p.Index in indices:
                    selected_pages.append(p)

        results = []
        doc_name = doc.FileName or "Untitled"
        base = os.path.splitext(doc_name)[0] if doc_name else "output"

        for page in selected_pages:
            page.Activate()
            ext = format if format != "ai" else "ai"
            out_path = os.path.join(output_dir, f"{base}_p{page.Index}.{ext}")
            try:
                _prepare_current_page_export(doc)
                if format == "pdf":
                    try:
                        doc.PDFSettings.PublishRange = _PDF_CURRENT_PAGE
                        doc.PDFSettings.PageRange = str(page.Index)
                    except Exception:
                        pass
                    doc.PublishToPDF(out_path)
                elif format == "dxf":
                    try:
                        expflt = doc.ExportEx(out_path, _CDR_DXF, _CDR_CURRENT_PAGE, None, None)
                        expflt.BitmapType = 0
                        expflt.TextAsCurves = True
                        expflt.Version = 1
                        expflt.Units = 3
                        _finish_export_filter(expflt)
                    except Exception:
                        doc.Export(out_path, _CDR_DXF, _CDR_CURRENT_PAGE, None, None)
                elif format == "png":
                    page_w_mm, page_h_mm = _page_size_mm(doc)
                    if page_w_mm > 0 and page_h_mm > 0:
                        pixel_w = max(1, int(page_w_mm / 25.4 * 150))
                        pixel_h = max(1, int(page_h_mm / 25.4 * 150))
                    else:
                        pixel_w = 0
                        pixel_h = 0
                    export_area = _page_export_area(conn.app, doc)
                    expflt = doc.ExportBitmap(
                        out_path,
                        _CDR_PNG,
                        _CDR_CURRENT_PAGE,
                        _CDR_RGB_COLOR_IMAGE,
                        pixel_w,
                        pixel_h,
                        150,
                        150,
                        _CDR_NORMAL_ANTIALIASING,
                        False,
                        False,
                        True,
                        False,
                        _CDR_COMPRESSION_NONE,
                        None,
                        export_area,
                    )
                    _finish_export_filter(expflt)
                elif format == "svg":
                    doc.Export(out_path, _CDR_SVG, _CDR_CURRENT_PAGE, None, None)
                else:
                    raise ValueError(f"Unsupported export format: {format}")
                _verify_exported(out_path)
                results.append({"page": page.Index, "path": out_path, "status": "success"})
            except Exception as e:
                results.append({"page": page.Index, "path": out_path, "status": "failed", "error": str(e)})

        return {"format": format, "total": len(results), "results": results}

    result = conn.safe_call(_batch, timeout=300.0)
    if result["success"]:
        r = result["result"]
        succeeded = sum(1 for x in r["results"] if x["status"] == "success")
        return ToolResult.ok(f"Batch export completed: {succeeded}/{r['total']} succeeded", **r)
    return ToolResult.fail(result.get("error", "Batch export failed"))
