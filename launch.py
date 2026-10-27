"""Load spreadsheet DLLs before starting the single COM STA worker."""
import os
from pathlib import Path
import runpy
import sys

root = Path(__file__).resolve().parent
os.chdir(root)
sys.path.insert(0, str(root / "server"))
import pandas  # noqa: F401 -- X8 host DLL initialization order
runpy.run_path(str(root / "server" / "server.py"), run_name="__main__")
