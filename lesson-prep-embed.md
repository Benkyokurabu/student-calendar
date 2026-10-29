# 勉たん内の授業準備シート

2026-10-29。単独ページの操作・先生プルダウン・「表示」「印刷」は維持する。

親URLは `https://line-check-system.vercel.app`。開発検証ではHTTPのloopback（127.0.0.1/localhost）のみ許可する。`embedded=1` が指定され、iframe内で、document.referrerのoriginが許可対象の場合だけ橋渡しを有効にする。第三者の認証cookieは使わない。

初回URL：`lesson_prep.html?embedded=1&teacher=工藤&date=2026-11-02`。

dateは実在する日付のみ適用。先生は初回候補ロード後に、その日の割当と一意に一致し、先生プルダウンに存在する場合だけ選択・表示する。授業なし・不明・候補重複時は通常の未選択プルダウンにする。確認済みalias（髙山/高山、佐藤（治）/治彦等）も候補が1件だけの場合に適用する。ロード中を含め利用者が先生・日付を操作した後は初期選択を上書きしない。

子→親は、初期読込後の `{type:'lesson-prep:ready',hasLessons,matchedTeacher?,height}`、表示状態変更時の `{type:'lesson-prep:availability',hasLessons,matchedTeacher?}`、高さ変更時の `{type:'lesson-prep:height',height}`。送信先は検証済みreferrer originで固定し、`*` を使わない。

親→子の `{type:'lesson-prep:print'}` は、sourceが直属parent、originが検証済みreferrerと一致、初期読込完了、現在選択と表示中の授業カードが一致する場合のみ `window.print()` を実行する。先生や日付を変えて「表示」する前は、前のシートの誤印刷を防ぐため `hasLessons:false` を通知する。従来のシート内印刷ボタンも維持する。

初期の最新API取得待ちと保存済みJSONへの切替は通常のロードとして扱う。実際のAPI失敗・タイムアウト時のみ保存済みデータの説明を表示する。APIと保存版の両方が取得できない場合は取得失敗を表示する。

回帰テスト：`node scripts/test_lesson_prep_loading_notice.mjs`、`node scripts/test_lesson_prep_embed.mjs`。GitHub Actionsにも登録する。実データ変更や通知送信は行わない。
