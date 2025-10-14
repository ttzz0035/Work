import os, sys
from pathlib import Path

# EXE内または開発環境で動作可能にする
base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
browsers_dir = base_dir / "playwright" / "ms-playwright"

os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers_dir)
os.environ["PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD"] = "1"
