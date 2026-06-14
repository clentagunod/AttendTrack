"""utils/logger.py — Logging for AttendTrack (file + in-app)."""

import logging
import sys
from pathlib import Path
from datetime import datetime


class AppLogger:
    """Dual logger: writes to file and also to an in-app callback."""

    def __init__(self, name: str, log_file: Path | None = None):
        self._log   = logging.getLogger(name)
        self._log.setLevel(logging.DEBUG)
        self._log.handlers.clear()
        self._callbacks = []

        # Console handler
        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(logging.DEBUG)
        ch.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
        self._log.addHandler(ch)

        # File handler
        if log_file:
            log_file.parent.mkdir(parents=True, exist_ok=True)
            fh = logging.FileHandler(log_file, encoding="utf-8")
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(logging.Formatter(
                "%(asctime)s [%(levelname)-8s] %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S"
            ))
            self._log.addHandler(fh)

    def add_callback(self, fn):
        """Register a UI callback fn(level, message)."""
        self._callbacks.append(fn)

    def _emit(self, level: str, msg: str):
        for fn in self._callbacks:
            try:
                fn(level, msg)
            except Exception:
                pass

    def info(self, msg: str):
        self._log.info(msg)
        self._emit("INFO", msg)

    def warning(self, msg: str):
        self._log.warning(msg)
        self._emit("WARN", msg)

    def error(self, msg: str):
        self._log.error(msg)
        self._emit("ERROR", msg)

    def debug(self, msg: str):
        self._log.debug(msg)
        self._emit("DEBUG", msg)

    def success(self, msg: str):
        self._log.info(f"✓ {msg}")
        self._emit("OK", msg)


_instance: AppLogger | None = None


def setup(name: str, log_file: Path | None = None) -> AppLogger:
    global _instance
    _instance = AppLogger(name, log_file)
    return _instance


def get() -> AppLogger:
    global _instance
    if _instance is None:
        _instance = AppLogger("AttendTrack")
    return _instance