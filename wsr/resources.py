"""Locate bundled runtime assets, both in source checkouts and frozen (PyInstaller) apps."""

from __future__ import annotations

import sys
from pathlib import Path


def resource_root() -> Path:
    """Base directory for bundled read-only assets (templates, images).

    When packaged with PyInstaller the data files are unpacked to a temporary
    directory exposed as ``sys._MEIPASS``; otherwise they live next to the repo.
    """
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        return Path(bundle_dir)
    return Path(__file__).resolve().parent.parent


def resource_path(*parts: str) -> Path:
    return resource_root().joinpath(*parts)
