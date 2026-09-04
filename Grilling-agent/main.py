import os
import sys
from pathlib import Path

from aiohttp.web import run_app

# Make sure imports from src/ resolve when launched from repo root.
ROOT = Path(__file__).resolve().parent
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from app import app  # noqa: E402


if __name__ == "__main__":
    run_app(app, host="0.0.0.0", port=int(os.environ.get("PORT", 3978)))
