"""Where things live: the app folder and the settings file."""
from __future__ import annotations

import json
import os
import sys

SLUG = "draftsman"
SETTINGS_NAME = SLUG + "-settings.json"


def app_dir() -> str:
    """Folder that holds the app: beside the exe, beside the .app bundle (never
    inside it - writing there breaks the signature), or the repo root from source."""
    if getattr(sys, "frozen", False):
        exe = os.path.abspath(sys.executable)
        parts = exe.split(os.sep)
        for i, p in enumerate(parts):
            if p.endswith(".app"):
                return os.sep.join(parts[:i]) or os.sep
        return os.path.dirname(exe)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _writable(folder: str) -> bool:
    try:
        probe = os.path.join(folder, "." + SLUG + "-write-test")
        with open(probe, "w") as fh:
            fh.write("x")
        os.remove(probe)
        return True
    except OSError:
        return False


def data_dir() -> str:
    """Beside the app when that can be written to, otherwise ~/.<slug>."""
    folder = app_dir()
    if not _writable(folder):  # e.g. the app was dropped into /Applications
        folder = os.path.join(os.path.expanduser("~"), "." + SLUG)
        try:
            os.makedirs(folder, exist_ok=True)
        except OSError:
            pass
    return folder


def settings_path() -> str:
    return os.path.join(data_dir(), SETTINGS_NAME)


def load_settings() -> dict:
    try:
        with open(settings_path(), "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(settings: dict) -> None:
    try:
        path = settings_path()
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(settings, fh, indent=2)
        os.replace(tmp, path)
    except (OSError, TypeError, ValueError):
        pass


def documents_dir() -> str:
    home = os.path.expanduser("~")
    path = os.path.join(home, "Documents")
    return path if os.path.isdir(path) else home


def open_in_system(path: str) -> None:
    """Open a file or folder with whatever the system uses for it."""
    import subprocess
    try:
        if sys.platform == "win32":
            os.startfile(path)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass
