from __future__ import annotations

import os
import sys
from pathlib import Path


def is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def runtime_dir() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def bundle_dir() -> Path:
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return runtime_dir()


def portable_mode_enabled() -> bool:
    if is_frozen():
        return True
    if os.environ.get("DATACOLISA_PORTABLE", "").strip() == "1":
        return True
    base = runtime_dir()
    return (base / "portable.flag").exists() or (base / "DATACOLISA.portable").exists()


def data_dir() -> Path:
    base = runtime_dir()
    if portable_mode_enabled():
        return base / "data"
    return Path(os.environ.get("APPDATA", str(Path.home()))) / "DATACOLISA"


def exports_dir() -> Path:
    if portable_mode_enabled():
        return runtime_dir() / "exports"
    # En mode dev (non-portable), on utilise Documents/DATACOLISA
    docs = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Documents" / "DATACOLISA"
    return docs


def ensure_runtime_dirs() -> None:
    for path in (data_dir(), exports_dir()):
        path.mkdir(parents=True, exist_ok=True)


def settings_dir(app_name: str = "DATACOLISA") -> Path:
    if portable_mode_enabled():
        return data_dir()
    return Path(os.environ.get("APPDATA", str(Path.home()))) / app_name


def app_assets_dir() -> Path:
    candidates = [
        bundle_dir() / "assets",
        runtime_dir() / "assets",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[-1]


def presentation_assets_dir() -> Path:
    candidates = [
        bundle_dir() / "presentation" / "assets",
        runtime_dir() / "presentation" / "assets",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[-1]
