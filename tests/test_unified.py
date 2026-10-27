import os
from pathlib import Path
import subprocess
import sys
import tempfile
import time
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "server"))
from core import sessions
from core.connection import ConnectionConfig, CorelDrawConnection
from core.process_lock import corel_process_lock
from test_sessions import Document, DocumentConnection
from test_connection import FailingThread
from tools import preflight


class UnifiedTests(unittest.TestCase):
    def test_saved_session_cannot_fall_back_to_same_basename_in_other_directory(self):
        doc = Document()
        conn = DocumentConnection(doc)
        manager = sessions.SessionManager()
        bound = manager._register({"name": "test.cdr", "path": os.path.abspath("missing/test.cdr"), "ordinal": 1}, "edit_file")
        old = sessions.get_connection
        sessions.get_connection = lambda: conn
        try:
            with self.assertRaisesRegex(RuntimeError, "no longer open"):
                manager._resolve_document(bound)
        finally:
            sessions.get_connection = old

    def test_legacy_mutations_are_not_replayed_by_default(self):
        conn = object.__new__(CorelDrawConnection)
        conn.config = ConnectionConfig(max_retries=3)
        conn._com_thread = FailingThread(RuntimeError("RPC rejected after write"))
        result = conn.safe_call(lambda: None)
        self.assertFalse(result["success"])
        self.assertEqual(conn._com_thread.calls, 1)

    def test_unknown_overflow_cannot_pass(self):
        doc = Document()
        text = type("Text", (), {"Type": 1})()
        shape = type("Shape", (), {"Type": 6, "StaticID": 5, "Text": text})()
        doc.ActivePage.Shapes.values = [shape]
        old = preflight.get_connection
        preflight.get_connection = lambda: DocumentConnection(doc)
        try:
            result = preflight.check_text_overflow_all()
            self.assertTrue(result.success)
            self.assertFalse(result.data["passed"])
            self.assertEqual(len(result.data["unknown_items"]), 1)
        finally:
            preflight.get_connection = old

    @unittest.skipUnless(sys.platform == "win32", "Windows named mutex")
    def test_mutex_serializes_separate_server_processes(self):
        with tempfile.TemporaryDirectory() as folder:
            marker = Path(folder) / "acquired"
            source = (
                "import sys; from pathlib import Path; "
                f"sys.path.insert(0,{str(Path(__file__).resolve().parents[1] / 'server')!r}); "
                "from core.process_lock import corel_process_lock; print('ready',flush=True); "
                f"\nwith corel_process_lock(): Path({str(marker)!r}).write_text('acquired')"
            )
            with corel_process_lock():
                child = subprocess.Popen([sys.executable, "-c", source], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
                self.assertEqual(child.stdout.readline().strip(), "ready")
                time.sleep(0.2)
                self.assertFalse(marker.exists())
            stdout, stderr = child.communicate(timeout=10)
            self.assertEqual(child.returncode, 0, stderr)
            self.assertTrue(marker.exists())


if __name__ == "__main__":
    unittest.main()
