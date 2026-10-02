"""Administrator desktop UI; credentials stay in Git Credential Manager."""
import json
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
from urllib.request import Request, urlopen
import webbrowser

REPO = 'Benkyokurabu/student-calendar'
RAW = f'https://raw.githubusercontent.com/{REPO}/main/'
WORKFLOW = f'https://github.com/{REPO}/actions/workflows/recording-release.yml'

def request_json(url, data=None, token=None):
    headers = {'Accept': 'application/vnd.github+json', 'User-Agent': 'BenkyoRecordingRelease'}
    if token:
        headers['Authorization'] = 'Bearer ' + token
    body = json.dumps(data).encode() if data is not None else None
    with urlopen(Request(url, data=body, headers=headers), timeout=45) as response:
        raw = response.read()
        return json.loads(raw) if raw else None

def dispatch(event_key, mode, release_at):
    result = subprocess.run(['git', 'credential', 'fill'],
        input='protocol=https\nhost=github.com\n\n', text=True,
        capture_output=True, timeout=60, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
    credential = dict(line.split('=', 1) for line in result.stdout.splitlines() if '=' in line)
    token = credential.get('password')
    if result.returncode or not token:
        raise RuntimeError('GitHubの保存済みログインを取得できません。公開予約のGitHub画面から設定してください。')
    request_json(f'https://api.github.com/repos/{REPO}/actions/workflows/recording-release.yml/dispatches',
        {'ref': 'main', 'inputs': {'event_key': event_key, 'mode': mode, 'release_at': release_at}}, token)

class App:
    def __init__(self, root):
        self.root = root
        root.title('録画の公開設定')
        root.geometry('980x620')
        self.entries = {}
        top = ttk.Frame(root, padding=15)
        top.pack(fill='x')
        ttk.Label(top, text='録画の公開設定', font=('Yu Gothic UI', 18)).pack(anchor='w')
        ttk.Label(top, text='単元テスト解説は「非公開」または「日時を指定」。Zoom側で視聴を制限します。').pack(anchor='w', pady=8)
        self.month = tk.StringVar(value=datetime.now().strftime('%Y-%m'))
        row = ttk.Frame(top)
        row.pack(fill='x')
        ttk.Label(row, text='対象年月').pack(side='left')
        ttk.Entry(row, textvariable=self.month, width=9).pack(side='left', padx=8)
        self.reload = ttk.Button(row, text='録画と公開状態を読み込む', command=self.load)
        self.reload.pack(side='left')
        self.tree = ttk.Treeview(root, columns=('lesson', 'status'), show='headings', selectmode='browse')
        self.tree.heading('lesson', text='授業（日付・時間・校舎・クラス）')
        self.tree.heading('status', text='公開設定')
        self.tree.column('lesson', width=620)
        self.tree.column('status', width=260)
        self.tree.pack(fill='both', expand=True, padx=15)
        bottom = ttk.Frame(root, padding=15)
        bottom.pack(fill='x')
        self.mode = tk.StringVar(value='日時を指定して公開')
        ttk.Combobox(bottom, textvariable=self.mode, state='readonly', width=23,
                     values=['日時を指定して公開', '非公開', '今すぐ公開']).pack(side='left')
        self.release = tk.StringVar(value=(datetime.now() + timedelta(days=7)).strftime('%Y-%m-%dT%H:%M'))
        ttk.Label(bottom, text='日本時間').pack(side='left', padx=8)
        ttk.Entry(bottom, textvariable=self.release, width=20).pack(side='left')
        self.save = ttk.Button(bottom, text='選択した録画に設定', command=self.submit)
        self.save.pack(side='left', padx=10)
        ttk.Button(bottom, text='クラウドの実行結果', command=lambda: webbrowser.open(WORKFLOW)).pack(side='left')
        self.status = tk.StringVar(value='読み込み中…')
        ttk.Label(root, textvariable=self.status, wraplength=950).pack(fill='x', padx=15, pady=(0, 15))
        self.load()

    def background(self, work, done):
        self.save.configure(state='disabled')
        self.reload.configure(state='disabled')
        def worker():
            try:
                result = work()
                self.root.after(0, lambda: complete(result, None))
            except Exception as error:
                message = str(error)
                self.root.after(0, lambda: complete(None, message))
        def complete(result, error):
            self.save.configure(state='normal')
            self.reload.configure(state='normal')
            if error:
                self.status.set('設定・取得に失敗しました。成功扱いにはしていません。')
                messagebox.showerror('録画公開設定', error)
            else:
                done(result)
        threading.Thread(target=worker, daemon=True).start()

    def load(self):
        month = self.month.get()
        try:
            datetime.strptime(month, '%Y-%m')
        except ValueError:
            messagebox.showerror('対象年月', '2026-10 の形式で入力してください。')
            return
        def work():
            stamp = str(datetime.now().timestamp())
            entries = request_json(RAW + f'zoom_recording_urls_{month}.json?t={stamp}')['entries']
            state = request_json(RAW + f'recording_release.json?t={stamp}')
            return entries, state
        def done(result):
            self.entries, state = result
            rules = {k: r for r in state['rules'] for k in r['eventKeys']}
            self.tree.delete(*self.tree.get_children())
            for key, entry in sorted(self.entries.items()):
                rule = rules.get(key)
                status = '通常公開' if rule is None else ('設定未完了（再試行中）' if rule.get('status') == 'pending' else ('公開済み' if rule.get('status') == 'released' else ('非公開' if rule['mode'] == 'private' else '公開待ち ' + rule['releaseAt'])))
                label = f"{entry['date']} {entry['time']} {'本校' if entry['campus'] == 'hon' else '南教室'} {entry.get('label', '')}"
                self.tree.insert('', 'end', iid=key, values=(label, status))
            self.status.set('授業を選び、公開方法を設定してください。同じZoom録画の両校舎・再開分にも適用されます。')
        self.background(work, done)

    def submit(self):
        selection = self.tree.selection()
        if not selection:
            messagebox.showinfo('授業を選択', '対象の授業を一覧から選んでください。')
            return
        mode = {'日時を指定して公開': 'scheduled', '非公開': 'private', '今すぐ公開': 'public'}[self.mode.get()]
        release = self.release.get().strip() if mode == 'scheduled' else ''
        if mode == 'scheduled':
            try:
                if datetime.strptime(release, '%Y-%m-%dT%H:%M') <= datetime.now():
                    raise ValueError()
            except ValueError:
                messagebox.showerror('公開日時', '未来の日本時間を 2026-10-10T22:00 の形式で入力してください。')
                return
        self.status.set('クラウドへ設定を送信しています…')
        sent_at = datetime.now().astimezone()
        event_key = selection[0]
        def sent(_):
            self.status.set('クラウドへ送信しました。Zoom側への反映を確認しています…')
            self.root.after(10000, lambda: self.poll(event_key, mode, release, sent_at, 0))
        self.background(lambda: dispatch(event_key, mode, release), sent)

    def poll(self, key, mode, release, sent_at, attempt):
        def work():
            return request_json(RAW + 'recording_release.json?t=' + str(datetime.now().timestamp()))
        def done(state):
            rule = next((r for r in state['rules'] if key in r['eventKeys']), None)
            fresh = datetime.fromisoformat(state['generatedAt']) >= sent_at
            if rule and fresh and rule['mode'] == mode and rule['releaseAt'] == release:
                if rule['status'] == 'pending':
                    self.status.set('Zoom側への反映が未完了です。権限不足などの理由は「クラウドの実行結果」に表示されています。')
                else:
                    self.load()
                    messagebox.showinfo('設定完了', 'Zoom側の共有設定の反映を確認しました。' + ('指定日時以降にクラウドで公開します。' if mode == 'scheduled' else ''))
                return
            if attempt >= 60:
                self.status.set('クラウドの確認に時間がかかっています。まだ設定完了ではありません。「クラウドの実行結果」を確認してください。')
                return
            self.root.after(10000, lambda: self.poll(key, mode, release, sent_at, attempt + 1))
        self.background(work, done)

if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()
