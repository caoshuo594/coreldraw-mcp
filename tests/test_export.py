import sys
import tempfile
import unittest
from types import SimpleNamespace
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "server"))

from tools import export as export_tools


class Status:
    connected = True


class Collection:
    def __init__(self, values):
        self.values = values
    @property
    def Count(self):
        return len(self.values)
    def __iter__(self):
        return iter(self.values)


class Layer:
    def __init__(self, visible, printable):
        self.Visible = visible
        self.Printable = printable


class Page:
    def __init__(self):
        self.SizeWidth = 210.0
        self.SizeHeight = 297.0
        self.Index = 1
        self.Shapes = Collection([object()])
        self.Layers = Collection([Layer(False, False), Layer(True, False)])
    def Activate(self):
        return None


class Settings:
    pass


class ExportFilter:
    def __init__(self, path):
        self.path = path
    def Finish(self):
        Path(self.path).write_bytes(b"exported")


class Document:
    def __init__(self):
        self.ActivePage = Page()
        self.Unit = 3
        self.PDFSettings = Settings()
        self.clear_selection_calls = 0
    def ClearSelection(self):
        self.clear_selection_calls += 1
    def PublishToPDF(self, path):
        Path(path).write_bytes(b"pdf")
    def Export(self, path, *args):
        Path(path).write_bytes(b"export")
    def ExportEx(self, path, *args):
        raise RuntimeError("exercise documented fallback")
    def ExportBitmap(self, path, *args):
        return ExportFilter(path)


class App:
    def __init__(self, doc):
        self.ActiveDocument = doc
    def CreateRect(self, *args):
        return args


class Connection:
    def __init__(self):
        self.status = Status()
        self.doc = Document()
        self.app = App(self.doc)
        self.timeouts = []
    def safe_call(self, func, *args, **kwargs):
        self.timeouts.append(kwargs.get("timeout"))
        try:
            return {"success": True, "result": func(*args)}
        except Exception as exc:
            return {"success": False, "result": None, "error": str(exc)}


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.conn = Connection()
        self.original_get_connection = export_tools.get_connection
        export_tools.get_connection = lambda: self.conn
        self.temp_dir = tempfile.TemporaryDirectory()

    def tearDown(self):
        export_tools.get_connection = self.original_get_connection
        self.temp_dir.cleanup()

    def path(self, name):
        return str(Path(self.temp_dir.name) / name)

    def test_all_single_export_functions_execute_without_closure_errors(self):
        calls = [
            (export_tools.export_pdf, self.path("a.pdf")),
            (export_tools.export_dxf, self.path("a.dxf")),
            (export_tools.export_ai, self.path("a.ai")),
            (export_tools.export_svg, self.path("a.svg")),
            (export_tools.export_png, self.path("a.png")),
            (export_tools.export_jpeg, self.path("a.jpg")),
            (export_tools.export_preview_png, self.path("preview.png")),
        ]
        for function, path in calls:
            with self.subTest(function=function.__name__):
                result = function(path)
                self.assertTrue(result.success, result.error)
                self.assertGreater(Path(path).stat().st_size, 0)
        self.assertEqual(self.conn.timeouts, [300.0] * len(calls))

    def test_preview_does_not_change_layer_state_or_selection(self):
        before = [(layer.Visible, layer.Printable) for layer in self.conn.doc.ActivePage.Layers]
        result = export_tools.export_preview_png(self.path("preview.png"))
        after = [(layer.Visible, layer.Printable) for layer in self.conn.doc.ActivePage.Layers]
        self.assertTrue(result.success, result.error)
        self.assertEqual(before, after)
        self.assertEqual(self.conn.doc.clear_selection_calls, 0)

    def test_artwork_bounds_include_external_objects_but_exclude_hidden_layers(self):
        visible = SimpleNamespace(Visible=True, Printable=True)
        hidden = SimpleNamespace(Visible=False, Printable=True)
        def box(layer, left, bottom, right, top):
            return SimpleNamespace(Layer=layer, LeftX=left, BottomY=bottom, RightX=right, TopY=top)
        self.conn.doc.ActivePage.Shapes = Collection([
            box(visible, 10, 20, 30, 40), box(visible, 260, -5, 280, 25),
            box(hidden, -500, -500, 500, 500),
        ])
        old_size = (self.conn.doc.ActivePage.SizeWidth, self.conn.doc.ActivePage.SizeHeight)
        rect, width, height = export_tools._export_geometry(self.conn.app, self.conn.doc, "artwork", 2)
        self.assertEqual(rect, (8, -7, 274, 49))
        self.assertEqual((width, height), (274, 49))
        self.assertEqual(old_size, (self.conn.doc.ActivePage.SizeWidth, self.conn.doc.ActivePage.SizeHeight))

    def test_artwork_export_failure_never_falls_back_to_cropped_page(self):
        layer = SimpleNamespace(Visible=True, Printable=True)
        self.conn.doc.ActivePage.Shapes = Collection([
            SimpleNamespace(Layer=layer, LeftX=260, BottomY=10, RightX=280, TopY=30)
        ])
        def fail(*args):
            raise RuntimeError("bounds unsupported")
        self.conn.doc.ExportBitmap = fail
        for fn, name in [(export_tools.export_preview_png, "fail.png"),
                         (export_tools.export_png, "fail2.png"),
                         (export_tools.export_jpeg, "fail.jpg")]:
            result = fn(self.path(name), scope="artwork")
            self.assertFalse(result.success)
            self.assertFalse(Path(self.path(name)).exists())

    def test_invalid_or_empty_artwork_scope_is_rejected(self):
        with self.assertRaises(ValueError):
            export_tools._export_geometry(self.conn.app, self.conn.doc, "unknown")
        with self.assertRaises(ValueError):
            export_tools._export_geometry(self.conn.app, self.conn.doc, "page", -1)
        self.conn.doc.ActivePage.Shapes = Collection([])
        result = export_tools.export_preview_png(self.path("empty.png"), scope="artwork")
        self.assertFalse(result.success)

    def test_workspace_uses_temporary_copy_route_for_all_raster_exports(self):
        original = export_tools._workspace_export
        calls = []
        def temporary_export(conn, path, exporter, padding_mm, **options):
            calls.append((exporter.__name__, padding_mm))
            Path(path).write_bytes(b"workspace")
            return {"scope": "workspace", "source_object_count": 4}
        export_tools._workspace_export = temporary_export
        try:
            for fn, name in [(export_tools.export_preview_png, "workspace.png"),
                             (export_tools.export_png, "workspace2.png"),
                             (export_tools.export_jpeg, "workspace.jpg")]:
                result = fn(self.path(name), scope="workspace", padding_mm=4)
                self.assertTrue(result.success, result.error)
                self.assertEqual(result.data["scope"], "workspace")
            self.assertEqual(calls, [("export_preview_png", 4), ("export_png", 4), ("export_jpeg", 4)])
        finally:
            export_tools._workspace_export = original


if __name__ == "__main__":
    unittest.main()
