"""One CorelDRAW MCP, one shared STA connection, compact tools plus feature discovery."""
from __future__ import annotations

import functools
import importlib
import inspect
import json
import os
from pathlib import Path
import threading

from fastmcp import FastMCP
from pydantic import TypeAdapter
from loguru import logger

from core.connection import close_connection, get_connection
from core.models import ToolResult
from core.sessions import session_manager
from tools import live
from tools.unified import open_document, verify_saved_document, merge_files

mcp = FastMCP("coreldraw")
_tool_lock = threading.RLock()
CATALOG = json.loads((Path(__file__).parent / "feature_catalog.json").read_text(encoding="utf-8"))
GLOBALS = {"read_excel_data", "list_templates", "get_template_info", "get_size_variant"}
READS = {"get_document_info", "list_all_text_shapes", "find_shape_by_name", "get_layers",
         "check_text_overflow", "check_dimensions", "check_text_overflow_all",
         "check_missing_fonts", "get_color_report", "check_rgb_colors"}


def _function(name):
    module = importlib.import_module("tools." + CATALOG[name])
    return getattr(module, name)


def list_feature_tools(category: str = "", name: str = "") -> ToolResult:
    """Discover advanced tools. Specify name for its JSON argument schema and coordinate convention."""
    try:
        names = [n for n, cat in CATALOG.items() if (not category or cat == category) and (not name or n == name)]
        if (category or name) and not names:
            raise ValueError("Unknown feature/category; query without filters for the catalog")
        entries = []
        for n in names:
            fn = _function(n)
            item = {"name": n, "category": CATALOG[n], "description": inspect.getdoc(fn),
                    "session_required": n not in GLOBALS | {"open_template", "create_document", "batch_merge", "merge_record"}}
            if name:
                item["input_schema"] = TypeAdapter(fn).json_schema()
                item["coordinates"] = (
                    "page-top-left mm; top_left=true" if n in {"create_rectangle", "create_ellipse", "create_line", "create_text_frame"}
                    else "legacy Corel native coordinates; inspect function description and known object before writing")
            entries.append(item)
        return ToolResult.ok("Advanced features; prefer apply_operations for common edits", features=entries, total=len(entries))
    except Exception as exc:
        return live._error(exc, "FEATURE_DISCOVERY_FAILED")


def run_feature_tool(name: str, arguments: dict, session_id: str = "", expected_revision: int = -1) -> ToolResult:
    """Run a discovered advanced feature. Document edits require a bound session and current revision.

    Call list_feature_tools(name=...) first for the argument schema. Template merges use explicit
    template/output paths and refuse overwrites. Legacy coordinates vary; common edits use apply_operations.
    """
    try:
        if name not in CATALOG:
            raise ValueError("Unknown feature; call list_feature_tools")
        fn = _function(name)
        inspect.signature(fn).bind(**arguments)
        # Validate types, including required fields, before any COM side effect.
        from pydantic import create_model, ConfigDict
        params = inspect.signature(fn).parameters
        model = create_model("FeatureArguments", __config__=ConfigDict(extra="forbid"), **{
            k: (p.annotation if p.annotation is not inspect.Parameter.empty else object,
                p.default if p.default is not inspect.Parameter.empty else ...)
            for k, p in params.items()})
        args = model.model_validate(arguments).model_dump()
        if name in GLOBALS:
            return fn(**args)
        if name == "open_template":
            return open_document(args["path"])
        if name == "create_document":
            factors = {"mm": 1, "cm": 10, "inch": 25.4, "in": 25.4}
            if args["unit"] not in factors:
                raise ValueError("Supported document units: mm, cm, inch, in")
            factor = factors[args["unit"]]
            return live.create_new_document(args["width"] * factor, args["height"] * factor)
        if name in {"batch_merge", "merge_record"}:
            return merge_files(name, args)
        if not session_id:
            raise ValueError("session_id is required; attach the exact document first")
        if name == "save_document":
            return live.save_document_as(session_id, args["path"]) if args.get("path") else live.save_current_document(session_id)
        if expected_revision < 0 and name not in READS and not name.startswith("export_"):
            raise ValueError("expected_revision is required for document changes")
        if name == "close_document":
            if args.get("save"):
                raise ValueError("Save explicitly with save_current_document before closing")
            session = session_manager.get(session_id)
            if session.revision != expected_revision:
                raise ValueError("Revision conflict; inspect the document again")
            def close_clean(doc):
                if doc.Dirty:
                    raise RuntimeError("Refusing to close a document with unsaved changes")
                doc.Close()
            session_manager.execute(session_id, close_clean)
            session_manager.detach(session_id)
            return ToolResult.ok("Closed the saved document")
        # Explicit destinations never overwrite existing files through the advanced route.
        if CATALOG[name] == "export":
            for key in ("path", "output_dir"):
                if args.get(key):
                    p = Path(args[key])
                    if not p.is_absolute() or p.exists():
                        raise ValueError("Advanced export requires a new absolute output path/directory")
        mutation = name not in READS and CATALOG[name] != "export"
        def invoke(doc):
            result = fn(**args)
            if not result.success:
                raise RuntimeError(result.error or result.message)
            return result.model_dump()
        payload = session_manager.execute(session_id, invoke, mutation=mutation,
            expected_revision=expected_revision if mutation else None, label="Feature: " + name, timeout=300)
        return ToolResult.ok("Advanced feature completed", **payload)
    except Exception as exc:
        return live._error(exc, "FEATURE_FAILED")


def _serialized(fn):
    @functools.wraps(fn)
    def call(*args, **kwargs):
        with _tool_lock:
            return fn(*args, **kwargs)
    return call


def register_tools():
    names = ("corel_status", "list_open_documents", "attach_active_document", "attach_open_document",
             "create_new_document", "inspect_document", "apply_operations", "render_preview",
             "undo_last_edit", "redo_last_edit", "save_current_document", "save_document_as",
             "export_document", "detach_document")
    for name in names:
        mcp.tool()(_serialized(getattr(live, name)))
    for fn in (open_document, list_feature_tools, run_feature_tool, verify_saved_document):
        mcp.tool()(_serialized(fn))


def main():
    logger.remove()
    logger.add(str(Path(__file__).parent.parent / "server.log"), rotation="5 MB", retention=3)
    register_tools()
    try:
        # Lazy attachment: starting MCP does not launch or edit CorelDRAW.
        mcp.run(transport="stdio", show_banner=False)
    finally:
        close_connection()


if __name__ == "__main__":
    main()
