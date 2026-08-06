# Zoom録画URLの手動差し替え

生徒用カレンダーで録画URLを手動で貼り替える場合は、月別ファイル `zoom_recording_overrides_YYYY-MM.json` の `overrides` に1件ずつ追加します。

例:

```json
{
  "date": "2026-08-06",
  "time": "6:35～8:05",
  "campus": "hon",
  "grade": "j3",
  "class": "A",
  "subject": "eng",
  "room": "1",
  "url": "https://us06web.zoom.us/rec/play/..."
}
```

`url` を空文字 `""` にすると、その授業の録画リンクを非表示にできます。

優先順位:

1. `zoom_recording_overrides_YYYY-MM.json`
2. `zoom_recording_overrides_latest.json`
3. 授業日誌Excelから抽出された `recordingUrl`
4. Zoomから自動取得した `zoom_recording_urls_YYYY-MM.json`
5. `zoom_recording_urls_latest.json`

コード:

- `campus`: `hon` または `minami`
- `grade`: `e4`, `e5`, `e6`, `j1`, `j2`, `j3`
- `subject`: `arith`, `eng`, `jp`, `math`, `sci`, `soc`
