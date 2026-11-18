"""Select one journal input location. Missing inputs never fall back elsewhere."""
from __future__ import annotations

import os
from pathlib import PurePosixPath

LOCATIONS = {
    "desktop": "デスクトップ/【完成版】授業日誌システム",
    "bentan": "デスクトップ/勉たん/業務データ/授業日誌",
}


def source_remote(profile=None):
    selected = os.environ.get("JOURNAL_INPUT_LOCATION", "desktop") if profile is None else profile
    if selected not in LOCATIONS:
        raise ValueError("Unknown journal input location; refusing to use a different original")
    return "onedrive:" + LOCATIONS[selected]


def source_file(name, profile=None):
    if not name or PurePosixPath(name).name != name or "\\" in name or name in (".", ".."):
        raise ValueError("A single input filename is required")
    return source_remote(profile) + "/" + name


if __name__ == "__main__":
    print(source_remote())
