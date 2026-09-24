from pathlib import Path
import os
import sys


ROOT = Path(__file__).resolve().parents[1]
APP_DIR = ROOT / "Gestion_Celular"
DATA_DIR = Path(os.environ.get("SIGE_DATA_DIR", "/tmp/sige"))

os.environ.setdefault("SIGE_DATA_DIR", str(DATA_DIR))
sys.path.insert(0, str(APP_DIR))

from app_web import app, prepare_database  # noqa: E402


prepare_database()

