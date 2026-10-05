import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from publish_lock import PublishLock, process_is_alive

class ProcessLockTests(unittest.TestCase):
    def test_live_owner_is_not_stale_or_terminated(self):
        with tempfile.TemporaryDirectory() as tmp:
            child = subprocess.Popen([sys.executable, "-c", "import time; time.sleep(30)"])
            try:
                lock = PublishLock(Path(tmp) / "lock", purpose="test")
                lock.lock_path.write_text(json.dumps({"pid": child.pid}), "utf-8")
                self.assertTrue(process_is_alive(child.pid))
                self.assertFalse(lock._remove_if_stale())
                self.assertIsNone(child.poll())
            finally:
                child.terminate()
                child.wait()

    def test_dead_owner_can_be_recovered(self):
        with tempfile.TemporaryDirectory() as tmp:
            child = subprocess.Popen([sys.executable, "-c", "pass"])
            child.wait()
            lock = PublishLock(Path(tmp) / "lock", purpose="test")
            lock.lock_path.write_text(json.dumps({"pid": child.pid}), "utf-8")
            self.assertFalse(process_is_alive(child.pid))
            self.assertTrue(lock._remove_if_stale())
            self.assertFalse(lock.lock_path.exists())

    def test_reused_lock_remains_owned_until_release(self):
        with tempfile.TemporaryDirectory() as tmp:
            lock = PublishLock(Path(tmp) / "lock", purpose="test")
            lock.acquire()
            token = lock.lock_path.read_text("utf-8")
            lock.acquire()
            self.assertEqual(token, lock.lock_path.read_text("utf-8"))
            lock.release()
            self.assertFalse(lock.lock_path.exists())

if __name__ == "__main__":
    unittest.main()
