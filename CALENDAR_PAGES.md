# 運用中のカレンダー

- `calendar.html`：普段使うスケジュール表。画面修正・案内・動作確認はこちらを基準にする。
- `teacher_calendar.html`：先生用スケジュール表。
- `lesson_prep.html`：授業準備シート。

2026-10-19から、3画面のスケジュールは `calendar-schedule-source.js` を通じて勉たんの読み取り専用 `/api/schedule/calendar?month=YYYY-MM` から取得する。原本の取り込みが完了すれば同じ授業DBを表示し、別のJSON公開を待たない。開いたままの画面は約1分で再確認する。通信に失敗した場合のみ同じ月の保存済みJSONに切り替え、注意文を表示する。別月のlatestデータを対象月として代用しない。

先生用アーカイブは `scripts/build_teacher_calendar.py` で生徒用と同じイベントから生成する。日誌本文・録画は従来の公開処理を維持する。担当が空欄の授業も省略せず、先生用画面では「未判定」で選択できる。回帰確認は `node scripts/test_calendar_schedule_source.mjs`、`node scripts/test_calendar_live.mjs`、既存の `node scripts/test_calendar_mobile.mjs`、`python -B -m unittest discover -s scripts -p test_build_teacher_calendar.py -v`。

`calendar_journal.html` の専用画面は利用していないため廃止した。古いブックマークを壊さないため、通常の `calendar.html` へ転送する小さなファイルだけを残す。独自UIや機能を追加せず、日誌付き画面としての保守・ブラウザ操作テストは不要。

`calendar_journal0309backup.html` も削除済み。必要な旧コードはGit履歴から取得できる。

`journal_*.json` / `journal_map_*.json` と授業日誌Excelは、通常のスケジュール表・授業準備で使用する現役のデータであり、画面廃止の削除対象ではない。
