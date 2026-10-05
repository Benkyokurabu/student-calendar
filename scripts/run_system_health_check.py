"""Run and log the health monitor from the stable runtime."""
import contextlib
from datetime import datetime
from pathlib import Path
import sys
import traceback

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
if __name__ == "__main__":
    logs = ROOT / "logs"
    logs.mkdir(exist_ok=True)
    with (logs / "system_health_task.log").open("a", encoding="utf-8", buffering=1) as output:
        with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
            print(f"[{datetime.now().isoformat()}] START", flush=True)
            try:
                import system_health_check as health
                code = health.main()
            except Exception:
                traceback.print_exc()
                code = 1
            print(f"[{datetime.now().isoformat()}] END exit={code}", flush=True)
    raise SystemExit(code)
