"""Explicit file opening, isolated saved-file inspection, and safe template merge."""
from __future__ import annotations

import hashlib
import os
from pathlib import Path
import shutil
import tempfile
from core.connection import get_connection
from core.models import ToolResult
from core.sessions import session_manager, _full_path
from core.operations import inspect_document_raw, _walk_shapes
from tools import live
from tools.text import _set_text_content


def _source(path):
    p = Path(path)
    if not p.is_absolute() or not p.is_file() or p.suffix.lower() not in {".cdr", ".cdt"}:
        raise ValueError("An existing absolute .cdr/.cdt source path is required")
    return p.resolve()


def open_document(path: str) -> ToolResult:
    """Open the exact absolute CDR/CDT path, or attach if already open. Returns a document session."""
    try:
        p = _source(path)
        docs = session_manager.list_documents()
        if not any(os.path.normcase(d["path"]) == os.path.normcase(str(p)) for d in docs):
            conn = get_connection()
            result = conn.safe_call(lambda: conn.app.OpenDocument(str(p)), idempotent=False)
            if not result["success"]:
                raise RuntimeError(result["error"])
        return live.attach_open_document(str(p))
    except Exception as exc:
        return live._error(exc, "OPEN_FAILED")


def verify_saved_document(path: str, preview_path: str = "", width_px: int = 1200,
                          preview_scope: str = "page", padding_mm: float = 0) -> ToolResult:
    """Reopen an isolated byte-identical copy of saved CDR; report pages, units, scale, editable text/font/overflow.

    Does not save/close the user's open document. Optional PNG is rendered from the reopened disk copy.
    Content completeness and visual composition still require comparison to the task's source inventory.
    """
    try:
        source = _source(path)
        dest = live._validated_destination(preview_path, suffixes={".png"}, overwrite=False) if preview_path else None
        session_manager._ensure_connected()
        conn = get_connection()
        def verify():
            previous = conn.app.ActiveDocument if conn.app.Documents.Count else None
            with tempfile.TemporaryDirectory(prefix="corel-verify-") as tmp:
                copy = Path(tmp) / source.name
                shutil.copyfile(source, copy)
                digest = hashlib.sha256(copy.read_bytes()).hexdigest()
                doc = conn.app.OpenDocument(str(copy))
                try:
                    doc.Activate()
                    original_unit = int(doc.Unit)
                    ruler_units = {"horizontal": int(doc.Rulers.HUnits), "vertical": int(doc.Rulers.VUnits)}
                    doc.Unit = 3
                    pages, texts = [], []
                    for page in doc.Pages:
                        page.Activate()
                        data = inspect_document_raw(doc, int(page.Index), True, 10000)
                        try:
                            scale = float(doc.WorldScale)
                        except Exception:
                            scale = None
                        data["world_scale"] = scale
                        pages.append(data)
                        for shape, _ in _walk_shapes(page.Shapes):
                            if int(shape.Type) == 6:
                                item = {"page": int(page.Index), "id": str(shape.StaticID),
                                        "text": str(shape.Text.Story.Text), "font": str(shape.Text.Story.Font),
                                        "size_pt": float(shape.Text.Story.Size)}
                                try:
                                    item["overflow"] = bool(shape.Text.Overflow)
                                except Exception:
                                    item["overflow"] = None
                                texts.append(item)
                    if dest:
                        doc.Pages(1).Activate()
                        result = live.export_tools.export_preview_png(str(dest), width_px=width_px,
                                                                     scope=preview_scope, padding_mm=padding_mm)
                        if not result.success:
                            raise RuntimeError(result.error)
                    return {"source": str(source), "sha256": digest, "verification_mode": "reopened_byte_copy",
                            "opened_unit": original_unit, "ruler_units": ruler_units, "inspection_unit": "mm", "pages": pages,
                            "editable_texts": texts, "preview": str(dest) if dest else None,
                            "preview_scope": preview_scope,
                            "overflow_count": sum(t["overflow"] is True for t in texts)}
                finally:
                    doc.Dirty = False  # Only the disposable verification copy.
                    doc.Close()
                    if previous is not None:
                        previous.Activate()
        result = conn.safe_call(verify, timeout=300, idempotent=False)
        if not result["success"]:
            raise RuntimeError(result["error"])
        return ToolResult.ok("Saved bytes reopened and inspected; compare report and preview to task requirements", **result["result"])
    except Exception as exc:
        return live._error(exc, "SAVED_VERIFICATION_FAILED")


def merge_files(name, args):
    source = _source(args["template_path"])
    if name == "batch_merge":
        from tools.data_merge import read_excel_data
        data = read_excel_data(args["data_path"])
        if not data.success:
            return data
        rows = data.data["records"]
        root = Path(args["output_dir"])
        if not root.is_absolute():
            raise ValueError("output_dir must be absolute")
        root = root.resolve()
        plans = []
        for row in rows:
            filename = args["naming_pattern"].format(**row)
            if not filename or Path(filename).name != filename or any(c in filename for c in '<>:"/\\|?*') or filename.endswith((".", " ")):
                raise ValueError("naming_pattern must produce a plain valid filename")
            plans.append((row, root / (filename + ".cdr")))
    else:
        plans = [(args["data"], Path(args["output_path"]))]
    targets = [live._validated_destination(str(p), suffixes={".cdr"}, overwrite=False) for _, p in plans]
    if len({os.path.normcase(str(p)) for p in targets}) != len(targets):
        raise ValueError("Duplicate output filenames; nothing was generated")
    session_manager._ensure_connected()
    conn = get_connection()
    def merge():
        previous = conn.app.ActiveDocument if conn.app.Documents.Count else None
        outputs = []
        try:
            for (row, _), destination in zip(plans, targets):
                # Work on a disposable disk copy; an already-open template cannot be overwritten.
                with tempfile.TemporaryDirectory(prefix="corel-merge-") as tmp:
                    copy = Path(tmp) / source.name
                    shutil.copyfile(source, copy)
                    doc = conn.app.OpenDocument(str(copy))
                    try:
                        found = set()
                        for page in doc.Pages:
                            for shape, _ in _walk_shapes(page.Shapes):
                                if int(shape.Type) == 6 and str(shape.Name) in row:
                                    key = str(shape.Name)
                                    _set_text_content(shape.Text, str(row[key]))
                                    found.add(key)
                        destination.parent.mkdir(parents=True, exist_ok=True)
                        if destination.exists():
                            raise FileExistsError(destination)
                        doc.SaveAs(str(destination), conn.app.CreateStructSaveAsOptions())
                        outputs.append({"path": str(destination), "matched_fields": sorted(found),
                                        "unmatched_fields": sorted(set(row) - found)})
                    finally:
                        doc.Dirty = False  # Only the generated working copy.
                        doc.Close()
        finally:
            if previous is not None:
                previous.Activate()
        return outputs
    result = conn.safe_call(merge, timeout=300, idempotent=False)
    if not result["success"]:
        return ToolResult.fail(result["error"] + "; inspect output directory before retrying; completed files are retained")
    return ToolResult.ok("Template merge saved; run verify_saved_document on outputs", outputs=result["result"], total=len(targets))
