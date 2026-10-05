import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile

SOURCE = Path(__file__).resolve().parent
BASE = Path(tempfile.mkdtemp(prefix='zoom-runtime-install-'))
STAGE = BASE / 'stable-runtime'
CLOUD = 'onedrive:デスクトップ/【完成版】授業日誌システム'
FILES = [
    'cloud_backup.py', 'publish_lock.py', 'student_calendar_repo.py',
    'recording_publication_filter.py',
    'zoomURL/publish_zoom_recording_url_json.py',
    'zoomURL/zoom_recording_urls.py', 'zoomURL/zoom_recording_url_list.py',
    'zoomURL/zoom_meeting_ids.json', 'zoomURL/zoom_recording_overrides.json',
    'zoomURL/.env', 'zoomURL/test_publish_zoom_recording_url_json.py',
    'zoomURL/test_zoom_lesson_only.py',
]
RCLONE = next((Path(os.environ['LOCALAPPDATA']) / 'Microsoft/WinGet/Packages').glob('Rclone.Rclone*/rclone-*/rclone.exe'))
STAGE.mkdir(exist_ok=True)
file_list = BASE / 'runtime-files.txt'
file_list.write_text('\n'.join(FILES), encoding='utf-8')
subprocess.run([str(RCLONE), 'copy', CLOUD, str(STAGE), '--files-from', str(file_list)], check=True)

# Use the committed and tested Windows lock implementation.
lock_file = STAGE / 'publish_lock.py'
lock_file.write_text((SOURCE.parent / 'publish_lock.py').read_text('utf-8'), 'utf-8')
for name in ('system_health_check.py', 'run_system_health_check.py'):
    (STAGE / name).write_text((SOURCE.parent / name).read_text('utf-8'), 'utf-8')

wrapper = (SOURCE / 'run_zoom_publisher.py').read_text('utf-8')
(STAGE / 'zoomURL/run_zoom_publisher.py').write_text(wrapper, 'utf-8')

runtime = Path(os.environ['LOCALAPPDATA']) / 'BenkyoClub/zoom-publisher'
python = Path(sys.executable)
launcher = f'''Option Explicit
Dim shell, code
Set shell = CreateObject("WScript.Shell")
code = shell.Run("""{python}"" -X utf8 -B ""{runtime / 'zoomURL/run_zoom_publisher.py'}""", 0, True)
WScript.Quit code
'''
(STAGE / 'zoomURL/run_zoom_publisher_hidden.vbs').write_text(launcher, 'ascii')
health_launcher = launcher.replace(str(runtime / 'zoomURL/run_zoom_publisher.py'), str(runtime / 'run_system_health_check.py') + ' --notify')
# Keep the argument outside the quoted script filename.
health_launcher = health_launcher.replace('run_system_health_check.py --notify""', 'run_system_health_check.py"" --notify')
(STAGE / 'run_system_health_check_hidden.vbs').write_text(health_launcher, 'ascii')

original = Path.home() / 'OneDrive/デスクトップ/【完成版】授業日誌システム'
(STAGE / 'runtime.json').write_text(json.dumps({'sharedLock': str(runtime.parent / 'student_calendar_publish.lock'), 'cloudSource': CLOUD}, ensure_ascii=False, indent=2), 'utf-8')

for file in STAGE.rglob('*.py'):
    compile(file.read_text('utf-8-sig'), str(file), 'exec')
for pattern in ('test_publish_zoom_recording_url_json.py', 'test_zoom_lesson_only.py'):
    subprocess.run([sys.executable, '-X', 'utf8', '-B', '-m', 'unittest', 'discover', '-s', str(STAGE / 'zoomURL'), '-p', pattern], check=True, cwd=STAGE)
sys.path.insert(0, str(STAGE))
from publish_lock import PublishLock, process_is_alive
import tempfile
with tempfile.TemporaryDirectory() as temp:
    owner = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(10)'])
    try:
        assert process_is_alive(owner.pid)
        lock = PublishLock(Path(temp) / 'lock', purpose='probe')
        lock.lock_path.write_text(json.dumps({'pid': owner.pid}), 'utf-8')
        assert lock._remove_if_stale() is False
        assert owner.poll() is None, 'Liveness check terminated the owner'
    finally:
        owner.terminate()
        owner.wait()
    assert not process_is_alive(owner.pid)
    assert lock._remove_if_stale() is True
print('Prepared stable runtime; payload tests, syntax and live/dead owner checks passed.')
integrity = {str(p.relative_to(STAGE)): hashlib.sha256(p.read_bytes()).hexdigest() for p in STAGE.rglob('*') if p.suffix in ('.py', '.vbs')}
(STAGE / 'runtime_integrity.json').write_text(json.dumps({'files': integrity}, indent=2), 'utf-8')

manifest = {str(p.relative_to(STAGE)): hashlib.sha256(p.read_bytes()).hexdigest() for p in STAGE.rglob('*') if p.is_file()}
(BASE / 'stable-runtime-manifest.json').write_text(json.dumps(manifest, indent=2), 'utf-8')

print('Preparation directory:', BASE)
