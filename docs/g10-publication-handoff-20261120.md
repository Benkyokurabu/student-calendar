# G10への公開処理移行

2026-11-20の移行準備。GitHub Pagesの公開元は、切替確認後に `main` から `g10-live-20261120` へ変更する。公開先URLは維持する。

## 変更箇所

- 日誌更新、Zoom録画URL公開、録画公開判定、録画解禁、日誌管理の既存5 workflowは、定義とOIDC認証の参照元を `main` に維持する。実際に読み書きするcheckoutとJSONのpush先は `g10-live-20261120` にする。
- `BENTAN_PUBLICATION_BRANCH` は、未設定時に旧PC互換の `main` を使用する。新workflowは `g10-live-20261120` を明示する。その指定で別branchをcheckoutしている場合、書込み前に停止する。
- 新公開branchの `bentan-endpoints.json` は新PCの公開APIと校内職員画面を指定する。main側の同ファイルは旧入口のまま保持し、更新された旧PC処理が新DBへ無断で接続することを避ける。
- 旧入口は移行中に処理を保留し、移行後は必要な署名Webhook・既存PDF作成worker・認証済みGitHub callbackだけを新PCへ転送する。旧PCからの録画公開policy取得とPAT captureは旧入口で退役させる。
- 公開HTMLのGitHub raw予備読込み先も `g10-live-20261120` とする。mainへの古いpending commit pushで、公開中の新branchのHTMLやJSONは変わらない。

## ほかのPCで変更する場合

アプリ本体の機能変更は、従来どおり `line-check-system` のmainで行う。この公開branchの分離とは別である。

時間割・日誌の公開処理を変更する場合は、mainのworkflow定義と、新公開branchの実行コードを両方確認する。新公開branchへmain全体を機械的に上書きしない。最新の公開JSONと新PC向けendpoint設定を維持したうえで、対象コードだけを反映する。

旧PCのOSやプロセスが実際に停止したとは扱わない。旧DBのclaim停止・書込拒否と、新公開branchへの分離で二重処理を防ぐ。北辰の自動作成は今回の変更対象に含めない。

## 確認

実bare Gitを使い、旧mainへのHTML/JSONのpending commit push後も新公開branchのtree全体が変わらないこと、新担当のpending commit回復が新公開branchへだけpushされることを試験している。

既存 `publish-zoom-recording-urls.yml` の手動実行には `verify_connection_only` を追加した。trueの場合はGitHub OIDCで新PCへ空のrecordingsを送って認証を確認し、録画公開・Git commit・障害Issueへの投稿は実行しない。

準備commitのpushだけで運用切替済みとは扱わない。新PCのDB・添付・アプリ採用とバックアップ検証後に、新公開branchのendpoint、repository変数、Pages公開元、既存5workflowの担当を切り替える。
