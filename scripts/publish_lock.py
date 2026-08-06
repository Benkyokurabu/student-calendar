#!/usr/bin/env python3
"""student-calendarへの書き込み・pushを直列化するプロセス間ロック。"""

from __future__ import annotations

import json
import os
import time
import uuid
from pathlib import Path


class PublishLock:
    def __init__(
        self,
        lock_path: Path,
        *,
        purpose: str,
        timeout_seconds: int = 900,
        poll_seconds: float = 5.0,
        stale_seconds: int = 7200,
    ) -> None:
        self.lock_path = lock_path
        self.purpose = purpose
        self.timeout_seconds = timeout_seconds
        self.poll_seconds = poll_seconds
        self.stale_seconds = stale_seconds
        self.token = uuid.uuid4().hex
        self.acquired = False

    def _payload(self) -> str:
        return json.dumps(
            {
                "token": self.token,
                "pid": os.getpid(),
                "purpose": self.purpose,
                "createdAt": time.time(),
            },
            ensure_ascii=False,
        )

    def _remove_if_stale(self) -> bool:
        try:
            age = time.time() - self.lock_path.stat().st_mtime
        except FileNotFoundError:
            return True
        if age <= self.stale_seconds:
            return False
        stale_path = self.lock_path.with_name(
            self.lock_path.name + f".stale-{int(time.time())}"
        )
        try:
            self.lock_path.replace(stale_path)
            print(f"[LOCK] 古いロックを退避しました: {stale_path}")
            return True
        except FileNotFoundError:
            return True
        except OSError:
            return False

    def acquire(self) -> None:
        self.lock_path.parent.mkdir(parents=True, exist_ok=True)
        deadline = time.monotonic() + self.timeout_seconds
        announced = False
        while True:
            try:
                fd = os.open(str(self.lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                with os.fdopen(fd, "w", encoding="utf-8") as handle:
                    handle.write(self._payload())
                self.acquired = True
                print(f"[LOCK] 取得: {self.purpose}")
                return
            except FileExistsError:
                if self._remove_if_stale():
                    continue
                if not announced:
                    print(f"[LOCK] 別の公開処理の完了を待機中: {self.lock_path}")
                    announced = True
                if time.monotonic() >= deadline:
                    raise TimeoutError(f"公開ロック待機がタイムアウトしました: {self.lock_path}")
                time.sleep(self.poll_seconds)

    def release(self) -> None:
        if not self.acquired:
            return
        try:
            payload = json.loads(self.lock_path.read_text(encoding="utf-8"))
            if payload.get("token") == self.token:
                self.lock_path.unlink()
                print(f"[LOCK] 解放: {self.purpose}")
        except FileNotFoundError:
            pass
        finally:
            self.acquired = False

    def __enter__(self) -> "PublishLock":
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, traceback) -> None:
        self.release()
