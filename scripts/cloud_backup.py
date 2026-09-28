"""Verified backups written directly to the OneDrive root ``backup`` folder."""

from __future__ import annotations

import hashlib
import shutil
import subprocess
import tempfile
from pathlib import Path, PurePosixPath

RCLONE_REMOTE = "onedrive"
BACKUP_ROOT = "backup"


def find_rclone() -> str:
    found = shutil.which("rclone") or shutil.which("rclone.exe")
    if found:
        return found
    packages = Path.home() / "AppData" / "Local" / "Microsoft" / "WinGet" / "Packages"
    matches = sorted(packages.glob("Rclone.Rclone*/rclone-*/rclone.exe"), reverse=True)
    if matches:
        return str(matches[0])
    raise FileNotFoundError("rclone.exe が見つかりません")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_relative(value: str | Path) -> str:
    path = PurePosixPath(str(value).replace("\\", "/").strip("/"))
    if not path.parts or any(part in {"", ".", ".."} for part in path.parts):
        raise ValueError(f"invalid backup path: {value}")
    return path.as_posix()


def backup_remote_path(category: str, timestamp: str, relative_path: str | Path) -> str:
    category = _safe_relative(category)
    timestamp = _safe_relative(timestamp)
    relative = _safe_relative(relative_path)
    return f"{RCLONE_REMOTE}:{BACKUP_ROOT}/{category}/{timestamp}/{relative}"


def _run(command: list[str]) -> None:
    result = subprocess.run(
        command, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=900
    )
    if result.returncode:
        raise RuntimeError((result.stderr or result.stdout).strip())


def backup_file(source: Path, *, category: str, timestamp: str,
                relative_path: str | Path | None = None) -> str:
    """Upload a local file to OneDrive/backup and verify its SHA-256."""
    source = source.resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    destination = backup_remote_path(category, timestamp, relative_path or source.name)
    rclone = find_rclone()
    _run([rclone, "copyto", str(source), destination, "--metadata"])
    with tempfile.TemporaryDirectory(prefix="onedrive-backup-verify-") as temp:
        downloaded = Path(temp) / "verify.bin"
        _run([rclone, "copyto", destination, str(downloaded)])
        if _sha256(source) != _sha256(downloaded):
            raise RuntimeError(f"OneDrive backup verification failed: {destination}")
    return destination


def backup_cloud_file(source_remote: str, *, category: str, timestamp: str,
                      relative_path: str | Path) -> str:
    """Copy an existing cloud file into OneDrive/backup and verify the copy."""
    destination = backup_remote_path(category, timestamp, relative_path)
    rclone = find_rclone()
    _run([rclone, "copyto", source_remote, destination, "--metadata"])
    with tempfile.TemporaryDirectory(prefix="onedrive-backup-verify-") as temp:
        source = Path(temp) / "source.bin"
        copied = Path(temp) / "backup.bin"
        _run([rclone, "copyto", source_remote, str(source)])
        _run([rclone, "copyto", destination, str(copied)])
        if _sha256(source) != _sha256(copied):
            raise RuntimeError(f"OneDrive backup verification failed: {destination}")
    return destination
