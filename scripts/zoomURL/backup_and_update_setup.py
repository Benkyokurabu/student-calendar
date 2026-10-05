from datetime import datetime
import hashlib
import json
from pathlib import Path
import subprocess
import sys

BASE = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path(__file__).resolve().parent
STAGE = BASE / 'stable-runtime'
sys.path.insert(0, str(STAGE))
from cloud_backup import backup_file, backup_cloud_file, find_rclone

stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
category = 'zoom_task_runtime'
print(backup_file(BASE / 'tasks-before.json', category=category, timestamp=stamp, relative_path='scheduled_tasks/tasks-before.json'))
cloud = 'onedrive:デスクトップ/【完成版】授業日誌システム/zoomURL/setup_zoom_json_publish_tasks.py'
original = BASE / 'setup-original.py'
rclone = find_rclone()
subprocess.run([rclone, 'copyto', cloud, str(original)], check=True)
code = original.read_text('utf-8-sig')
needle = '    launcher_path = script_dir / "scheduled_zoom_recording_url_json_publish_hidden.vbs"'
if 'stable_launcher' in code:
    print('Cloud task installer already selects the stable runtime.')
    raise SystemExit(0)
assert needle in code
code = code.replace('import subprocess\n', 'import subprocess\nimport os\n')
code = code.replace(needle, '''    local_runtime = Path(os.environ.get("LOCALAPPDATA", "")) / "BenkyoClub" / "zoom-publisher"
    stable_launcher = local_runtime / "zoomURL" / "run_zoom_publisher_hidden.vbs"
    launcher_path = stable_launcher if stable_launcher.exists() else script_dir / "scheduled_zoom_recording_url_json_publish_hidden.vbs"''')
compile(code, 'setup_zoom_json_publish_tasks.py', 'exec')
print(backup_cloud_file(cloud, category=category, timestamp=stamp, relative_path='zoomURL/setup_zoom_json_publish_tasks.py'))
updated = BASE / 'setup-updated.py'
updated.write_text(code, 'utf-8')
subprocess.run([rclone, 'copyto', str(updated), cloud], check=True)
verified = BASE / 'setup-verified.py'
subprocess.run([rclone, 'copyto', cloud, str(verified)], check=True)
assert hashlib.sha256(updated.read_bytes()).digest() == hashlib.sha256(verified.read_bytes()).digest()
print('Task installer updated and verified directly on OneDrive cloud.')
