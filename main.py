"""
main.py — Entry point for AttendTrack.

Usage:
    python main.py
    python main.py --file "C:/path/to/attendance.xlsx"

MACROS — Edit these to change app-wide behavior quickly:
"""

# ── MACROS ────────────────────────────────────────────────────────────────────
APP_NAME         = "AttendTrack"
DEFAULT_THEME    = "darkly"          # Fallback if settings.json missing
SETTINGS_FILE    = "settings.json"
EXIT_ON_BAD_CFG  = True
# ──────────────────────────────────────────────────────────────────────────────

import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

def main():
    import argparse
    from core import config

    parser = argparse.ArgumentParser(description=f"{APP_NAME} — Attendance Tracker")
    parser.add_argument("--file", "-f", help="Auto-load an attendance Excel file on startup")
    args = parser.parse_args()

    # Load settings (create default if missing)
    try:
        settings = config.load()
    except FileNotFoundError as e:
        print(f"[ERROR] {e}")
        if EXIT_ON_BAD_CFG:
            sys.exit(2)
        settings = {}

    # Override theme from MACRO if settings missing
    if "app" not in settings:
        settings["app"] = {}
    settings["app"].setdefault("theme", DEFAULT_THEME)
    settings["app"].setdefault("name",  APP_NAME)

    # Auto-load file from CLI flag
    if args.file:
        settings["app"]["last_file"] = args.file
        settings["app"]["remember_last_file"] = True

    from gui.app import AttendTrackApp
    app = AttendTrackApp()
    app.run()


if __name__ == "__main__":
    main()