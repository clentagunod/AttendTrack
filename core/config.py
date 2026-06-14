"""core/config.py — Settings loader for AttendTrack."""

import json
import sys
from pathlib import Path


def _base_path() -> Path:
    """Resolve base directory whether running as script or frozen exe."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent


SETTINGS_FILE = _base_path() / "settings.json"


def load() -> dict:
    if not SETTINGS_FILE.exists():
        raise FileNotFoundError(f"settings.json not found at: {SETTINGS_FILE}")
    with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save(settings: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


# Convenience accessors
def app(s: dict) -> dict:      return s["app"]
def attendance(s: dict) -> dict: return s["attendance"]
def excel_cfg(s: dict) -> dict: return s["excel"]
def report_cfg(s: dict) -> dict: return s["report"]
def status_labels(s: dict) -> dict: return s["status_labels"]