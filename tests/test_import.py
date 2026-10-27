import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "server"))
from tools import shapes


class Collection:
    def __init__(self, values):
        self.values = values
    @property
    def Count(self):
        return len(self.values)
    def Item(self, index):
        return self.values[index - 1]


class Shape:
    def __init__(self, sid):
        self.StaticID = sid
        self.Name = str(sid)
        self.Type = 3
        self.SizeWidth, self.SizeHeight = 10, 20
        self.moves = []
    def SetPosition(self, x, y):
        self.moves.append((x, y))
    def SetSize(self, width, height):
        self.SizeWidth, self.SizeHeight = width, height


class ImportTests(unittest.TestCase):
    def setUp(self):
        self.old = Shape(1)
        self.added = [Shape(2), Shape(3)]
        self.page = SimpleNamespace(Shapes=Collection([self.old]))
        self.calls = []
        def import_ex(path, filter_id, options):
            self.calls.append((path, filter_id, options))
            def finish():
                self.page.Shapes.values[0:0] = self.added
            return SimpleNamespace(Finish=finish)
        self.page.ActiveLayer = SimpleNamespace(ImportEx=import_ex)
        self.doc = SimpleNamespace(ActivePage=self.page, Unit=1, ReferencePoint=9)
        self.grouped_members = []
        self.group = Shape(99)
        class Range:
            def Add(inner, shape):
                self.grouped_members.append(shape.StaticID)
            def Group(inner):
                return self.group
        self.app = SimpleNamespace(ActiveDocument=self.doc, CreateStructImportOptions=lambda: "options",
                                   CreateShapeRange=Range)
        def safe_call(fn):
            try:
                return {"success": True, "result": fn()}
            except Exception as exc:
                return {"success": False, "error": str(exc)}
        self.conn = SimpleNamespace(app=self.app, status=SimpleNamespace(connected=True), safe_call=safe_call)
        self.original = shapes.get_connection
        shapes.get_connection = lambda: self.conn
        self.tmp = tempfile.TemporaryDirectory()
        self.path = str(Path(self.tmp.name) / "source.svg")
        Path(self.path).write_text("<svg/>")
    def tearDown(self):
        shapes.get_connection = self.original
        self.tmp.cleanup()

    def test_layer_import_groups_only_new_objects_and_preserves_existing(self):
        result = shapes.import_svg(self.path, x=15, y=25)
        self.assertTrue(result.success, result.error)
        self.assertEqual(self.calls, [(self.path, 0, "options")])
        self.assertEqual(self.grouped_members, [2, 3])
        self.assertEqual(result.data["shape_id"], "99")
        self.assertEqual(self.group.moves, [(15, 25)])
        self.assertEqual(self.old.moves, [])
        self.assertEqual(self.doc.ReferencePoint, 9)

    def test_single_import_does_not_group_and_bitmap_uses_same_x8_interface(self):
        self.added = [Shape(4)]
        result = shapes.import_image(self.path, width=30, height=40)
        self.assertTrue(result.success, result.error)
        self.assertEqual(result.data["shape_id"], "4")
        self.assertEqual(self.grouped_members, [])
        self.assertEqual((self.added[0].SizeWidth, self.added[0].SizeHeight), (30, 40))

    def test_missing_file_is_rejected_before_import(self):
        result = shapes.import_svg(self.path + "missing")
        self.assertFalse(result.success)
        self.assertEqual(self.calls, [])


if __name__ == "__main__":
    unittest.main()
