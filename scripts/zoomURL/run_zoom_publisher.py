"""Run Zoom publication from a stable installation outside OneDrive."""
import contextlib
from datetime import datetime
import json
import os
from pathlib import Path
import subprocess
import sys
import traceback

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "zoomURL")]

def publish():
    import publish_zoom_recording_url_json as publisher
    from publish_lock import PublishLock
    from student_calendar_repo import resolve
    config = json.loads((ROOT / "runtime.json").read_text("utf-8"))
    repo = resolve()
    # All Windows publishers normalize this lock to the same stable location.
    lock = PublishLock(Path(config["sharedLock"]), purpose="Zoom録画URL公開")
    lock.acquire()
    try:
        subprocess.run(["git", "pull", "--rebase", "origin", "main"], cwd=repo, check=True, timeout=120)
        publisher.url_list.z.SYSTEM_DIR = repo
        publisher.url_list.SYSTEM_DIR = ROOT
        publisher.PublishLock = lambda *args, **kwargs: lock
        print("[runtime]", ROOT, flush=True)
        print("[schedule]", repo, flush=True)
        return publisher.main()
    finally:
        lock.release()

if __name__ == "__main__":
    logs = ROOT / "logs"
    logs.mkdir(exist_ok=True)
    with (logs / "zoom_recording_json.log").open("a", encoding="utf-8", buffering=1) as output:
        with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
            print(f"[{datetime.now().isoformat()}] START", flush=True)
            try:
                code = publish()
            except SystemExit as exc:
                code = exc.code if isinstance(exc.code, int) else 1
            except Exception:
                traceback.print_exc()
                code = 1
            print(f"[{datetime.now().isoformat()}] END exit={code}", flush=True)
    raise SystemExit(code)
