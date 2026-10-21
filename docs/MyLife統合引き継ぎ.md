# MyLife統合引き継ぎ v2.2

## 境界と保存先

core.jsは日本時間・起床計算、sync.jsはボタン操作に起因する統括のみ、app.jsは画面と接続、backend/Code.gsは認証・ACL・Sheets保存・手動reconcile・勤務元読取を担当する。OCR依存なし。

予定DBは既存のevents!A:T：id,owner,scope,title,start,end,allDay,notes,sync,calendarId,googleEventId,source,deleted,createdAt,updatedAt,location,revision,syncState,googleEtag,lastError。schema変更なし。設定と予定は別。tokenはRAMだけで永続化・ログ出力しない。

## 手動同期契約

通常保存はSheetsへ即時反映し外部送信はpending状態へ保留。起動、認証、月変更、復帰、保存ではCalendar events APIを呼ばない。認証・ACL確認やMyLife DB読取/保存は外部同期とは別。

「すべて同期」では勤務元をGETして共有DBへupsertし、その後Googleのshared/private reconcileとCalendar表示読取を実行する。各系統の例外を分離。Google未接続ならGoogle系統だけskip。勤務だけの実行にはMyLifeのSheets用ログインが必要であり、完全未ログインで会社の勤務表を取得する設計ではない。

Google Calendar連携解除はgoogleCalendarEnabled=false。Sheets保存の認証・予定データは保持する。連携を再有効化してもボタン操作まで予定を同期しない。last syncは成功日時とは別の最終試行日時で結果併記。localStorageにイベントやtokenを新規キャッシュしない（既存の体験用予定は従来キーを維持）。

## 安全性

共有/個人のGoogle ACL検証、revision、lock、Google ETag、安定したGoogle event IDを維持。sync OFFはGoogle側を残す。旧サーバーは自動送信するためprotocol 3確認まで新UIから書き込み禁止。pendingでも再編集できるがGoogle側ETagが変わっていれば上書きしない。競合解決の選択も次回手動同期で実行する。

Google由来予定は本人のみのCalendarならprivate DB、検証済み共有Calendarならshared DBに保留保存。他の共有Calendarをprivateへ誤取り込みしない。勤務はsource=shiftでshared DBに保存し、通常save/delete APIは拒否。日時ごとの既存行を再利用し、時刻未入力日は保持。元シート・備考・元シフトアプリに書かない。

## 実運用の残作業

backend/config.example.jsonの実設定、OAuthclient、gatewayデプロイ、2人の実ACL検証が未完了。GitHub Pages UIだけではGoogle接続が完成しない。docs/検証状況.mdの未完了項目を引き継ぐ。既存DBを初期化せず移行する。
