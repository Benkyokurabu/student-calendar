# 日誌原本の競合保護（2026-11-04）

旧処理は取得した原本を長時間加工したあと、ローカル保存時刻を基準に
`rclone copy --update --ignore-errors` で書き戻していた。その間の先生の
追記より加工ファイルの時刻が新しくなり、古い取得内容で原本を置換できた。
アップロード失敗も成功扱いにしていたため、原本と異なるJSONの公開も可能だった。

現在の処理は次の順序で実行する。

1. 加工前に、全原本の不変DriveItem ID・共有先drive ID・eTag・SHA256を保存する。
   Graphから取得したbytesと、加工対象のローカルbytesの一致を確認する。
   取得の前後でeTagが変わった原本や、未取得・未知の原本は処理を停止する。
2. 既存の月作成・前月同期・入力欄保守を行う。
3. 全原本のeTagを再確認する。書き戻しが必要なdriveでは、専用scratch内の
   新規UUIDテキストに対し、誤eTag拒否・現eTag成功・旧eTag拒否を毎回実証する。
4. 変更した原本だけを不変item IDで置換する。直前に元bytesのSHA256も確認し、
   加工前に固定したeTagを`If-Match`へ指定する。競合を無視する再試行は行わない。
5. 書き戻したbytesと新eTag、さらに未変更原本を含む全eTagを確認してから公開する。
   HTTP 412・423、認証エラー、通信断、再編集などはジョブを失敗にし、既存の公開JSONを保つ。

設定はリポジトリ変数 `JOURNAL_UPLOAD_PROBE_PARENTS` に
`{"対象drive ID":"専用scratch folder ID"}` の形式で登録する。
folder名は `.bentan-cas-probe-` で始まる専用folderに限定し、原本folderを指定できない。
実行時probeは新規に作成した自分のitem IDだけを内容・eTag確認後に削除する。
設定がない場合、変更原本の書き戻しは拒否する。新規・削除原本も既存版の保護なしで
作成せず、再取得を伴う別途対応として停止する。

実対象の共有driveで2026-11-04に隔離ファイルを用いて検証し、誤eTagは412かつ
初期bytes維持、現eTagは200かつ期待bytes一致、旧eTagは412かつ更新bytes維持を確認した。
原本への試験書き込みは行っていない。証拠は非公開の
`.private/journal-loss-20261104/conditional-put-capability-proof.json` に保存した。

自動試験は `python -m unittest discover -s scripts -p test_safe_journal_upload.py -v`。
取得中・保守中・PUT直前・PUT後の編集、旧baseline、ロック、認証失敗、条件無視、
応答喪失、未知の原本、共有drive取り違えを検証する。

安全処理24件・旧期限付きwriter11件を独立したTest student calendar workflowでも実行する。
原本更新workflowが停止中でも、保護コード・workflowの変更時に試験を実行する。
Update Schedule & Journalのvalidate_only=trueは検証用であり、全原本の取得・版固定・
実接続での条件付き保存拒否の検証だけを行う。保守・原本更新・JSON公開はプログラムの
条件でスキップする。通常の本番定期更新はfalseで実行し、上記すべての保護を通す。

改名・移動の競合も、Graph名と取得時ファイル名の一致、capture前後の全パスとID対応の
再一覧照合によって拒否する。重複パス・大小文字衝突・同一原本の別名登録も受理しない。
実接続41837233016は60冊の版固定後、南中1数学の版変更を検出して安全停止した。
保守・原本更新・JSON公開がすべてskipされたことをログで確認済み。

Graphの[ファイル内容置換](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0)を使用する。
小PUTのIf-MatchはMicrosoft担当者の[説明](https://learn.microsoft.com/en-au/answers/questions/1515186/how-to-specify-conflict-behavior-for-a-put-request)に加え、
上記の実接続実証と毎回の実行時検証を必須にしている。
複数原本全体のトランザクションではないため、途中失敗時に正常反映済みの原本を
古い内容へ戻す処理は行わない。次回は新しい原本を再取得して整合した内容を公開する。
