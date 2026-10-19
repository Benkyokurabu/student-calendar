# 2026-10-19 カレンダー共通配信の保存・復元

対象：生徒・先生用カレンダー

- 変更前の復元タグ：`calendar-sync-before-20261019`（8be128af00232945976c4c2f619795f9e9e1c116）
- 今回の検証済み復元タグ：`calendar-sync-after-20261019`（5915539c45f56f27ea7c03346daba8de8e390b39）
- 復元パッチの検証元：4d6e1dba6aaa230ef811de6cd2ef4aa241cf4832
- 生成日：2026-10-19

今回の機能変更はコミット・mainへのpush・本番反映済み。上記の2タグもGitHubに保存済み。今回の配信は11月386件（本校194・南教室192）を3画面と本番APIで確認した。

## 旧処理へ戻す場合

`calendar-sync-20261019.patch` は今回の画面・処理・関連設定だけを戻す逆差分。公開済みの月別スケジュール・先生用データ・日誌JSON、DB、OneDrive原本は変更しない。勉たんの他機能や後から追加されたテストを保持する。

作成時点のコードに対して `git apply --check` が成功した。まだ適用しておらず、本番は今回の版のまま。

作業時は最新mainから専用ブランチを作り、以下を順に実行する。

```powershell
git fetch origin main --tags
git switch -c rollback-calendar-sync-20261019 origin/main
git apply --check docs/rollback/calendar-sync-20261019.patch
git apply --index docs/rollback/calendar-sync-20261019.patch
git diff --cached --check
git diff --cached --stat
```

確認が失敗した場合は止め、変更された対象ファイルと逆差分を照合する。force、reset、履歴の上書きや、現在のデータを古いタグへ丸ごと置換する方法は使わない。最新mainとの競合を解消し、型・関連lint・関連テスト・ビルド（勉たん）とカレンダー表示を確認してから、戻す変更を新しいコミットにしてmainへpushする。

カレンダー側を先に旧処理へ戻し、静的JSONで11月が見えることを確認した後で、勉たん側のAPIを戻す。今回の方式へ再び切り替える場合は、その復元コミットをrevertし、同じ検証を行う。タグは元の2版を保持するため動かさない。

全ファイルを当時の状態で参照したい場合は、復元タグを別の作業フォルダへチェックアウトして比較する。11月データや別作業を保持する運用復元では上記パッチを使用する。

## 復元対象

```text
.github/workflows/test-student-calendar.yml |    8 ---
 .github/workflows/update-journal.yml        |    8 ---
 CALENDAR_PAGES.md                           |    4 -
 calendar-schedule-source.js                 |   77 ---------------------------
 calendar.html                               |   29 ++++++----
 lesson_prep.html                            |   12 +---
 scripts/build_teacher_calendar.py           |   44 ---------------
 scripts/test_build_teacher_calendar.py      |   36 -------------
 scripts/test_calendar_live.mjs              |   77 ---------------------------
 scripts/test_calendar_mobile.mjs            |    1 
 scripts/test_calendar_schedule_source.mjs   |   59 ---------------------
 teacher_calendar.html                       |   41 +++++++++-----
 12 files changed, 45 insertions(+), 351 deletions(-)
```
