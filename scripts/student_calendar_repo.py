#!/usr/bin/env python3
"""Resolve the dedicated student-calendar publishing clone.

The Git working tree should not live under OneDrive because syncing `.git`
metadata can corrupt an in-progress Git operation.  An environment override is
supported for recovery and migration.
"""

from __future__ import annotations

import os
from pathlib import Path


def candidates() -> list[Path]:
    result: list[Path] = []
    override = os.environ.get("BENKYO_STUDENT_CALENDAR_REPO", "").strip()
    if override:
        result.append(Path(override).expanduser())

    local_app_data = os.environ.get("LOCALAPPDATA", "").strip()
    if local_app_data:
        result.append(Path(local_app_data) / "BenkyoClub" / "student-calendar")

    # Migration fallback only. New installations should use LOCALAPPDATA above.
    result.append(Path.home() / "OneDrive" / "デスクトップ" / "生徒スケジュール表")
    return result


def resolve(required: bool = True) -> Path | None:
    for path in candidates():
        if (path / ".git").is_dir():
            return path.resolve()
    if required:
        checked = "\n".join(f"- {path}" for path in candidates())
        raise FileNotFoundError(f"student-calendar公開用リポジトリが見つかりません:\n{checked}")
    return None
