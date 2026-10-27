"""Serialize unified MCP instances across chats; a timed-out STA task keeps its lock."""
from contextlib import contextmanager
import sys
import threading

_local = threading.RLock()


@contextmanager
def corel_process_lock():
    with _local:
        if sys.platform != "win32":
            yield
            return
        import win32event
        import win32api
        mutex = win32event.CreateMutex(None, False, r"Local\CodexCorelDrawUnifiedSTA")
        acquired = False
        try:
            outcome = win32event.WaitForSingleObject(mutex, 30000)
            if outcome not in (0, 0x80):
                raise TimeoutError("Another CorelDRAW operation is still running; inspect state before retrying")
            acquired = True
            yield
        finally:
            if acquired:
                win32event.ReleaseMutex(mutex)
            win32api.CloseHandle(mutex)
