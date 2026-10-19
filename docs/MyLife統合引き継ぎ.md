# MyLife本体への統合

## 境界

UI(app.js) → Google読み取りAPI／同期GAS → 本人のSheets・Calendar。勤務同期は独立GAS → 勤務カレンダー → PWA。家計簿や既存勤務システムのコードには依存しません。

core.jsは副作用のない日付・OCR候補・起床計算です。UIを移行しても再利用できます。app.jsにはUIとAPI呼び出しが残っており、完全なアダプター分割は今後の作業です。サーバーの保存・同期処理はbackendに分離済みです。

## eventsの20列（順序固定）

|列|意味|
|---|---|
|id|UUID、MyLifeの安定ID|
|owner|作成者メール、サーバーが設定|
|scope|shared / private|
|title|200文字以内|
|start / end|RFC3339の日時、終日はYYYY-MM-DD。endは排他的|
|allDay|true / false|
|notes|3000文字以内|
|sync|Google同期ON/OFF。共有範囲とは独立|
|calendarId / googleEventId|同期先。OFFにしても参照は保持|
|source|manual。外部Google予定は別表示で原則コピーしない|
|deleted|論理削除|
|createdAt / updatedAt|更新日時|
|location|300文字以内|
|revision|楽観ロック用の整数|
|syncState|local / pending-create / pending-update / pending-delete / synced / google-deleted|
|googleEtag|Googleの競合検出用ETag|
|lastError|同期失敗の概要。秘密情報を保存しない|

設定はlocalStorageのmylife.calendar.config.v1に保存。ユーザートークンはメモリのみ。クラウド予定は端末に永続キャッシュしません。体験モードの予定だけ別キーに保存します。GASのMYLIFE_CONFIGはサーバー側の信頼する設定で、本人メールから個人DBを決定します。

## 同期・競合

- 同じGASデプロイを2人で使用し、ScriptLockで保存を直列化。revisionが古ければ拒否。
- Sheetsに未完了の意図を先に書き、Google操作を後から実行。Google失敗時もローカルの意図を失わない。
- UUID由来のGoogleイベントIDで再試行時の二重作成を抑制。衝突はマーカーで照合。
- Googleの変更はreconcileでSheetsへ取り込み。通常は前面5分ごと／復帰時／手動。バックグラウンド永続同期はしない。
- GoogleとMyLifeの同時変更はETagを照合して停止。エラーの予定を開き、確認後「Google側の内容を採用して再開」が可能。MyLife案を優先する自動上書きは行わない。必要ならメモなどへ退避してからGoogle案を採用し、再編集する。
- 同期OFFはGoogle予定を残す。再ON時にGoogleのETagが変わっていれば競合として扱う。
- Googleでの削除はMyLifeの内容を同期OFFで保持する。外部Google予定は表示元から消える。
- 共有範囲の変更は既存予定では禁止。別DBへ移動する原子的なトランザクションがないため、データの移し替え機能は未実装。
- SheetsとCalendarに分散トランザクションはない。Google書込後にSheets反映が失敗した場合も意図を残して止めるが、完全なexactly-onceは保証しない。直接Sheets編集・別GASからの書込はロック対象外で禁止する。

## 移行

1. 2人の個人DBは統合後も別権限のまま維持する。
2. UUID、Google ID、revision、墓石を保持して移す。新IDを採番しない。
3. 既存v1の15列DBは自動移行しない。バックアップ、列追加とrevision初期化、既存Google IDの検証を別途行う。
4. OAuth生成元の追加と、必要ならGASのclientIdを同じクライアントへ維持する。
5. サービスワーカーはMyLife Calendar専用プレフィックスのみ削除し、他機能のキャッシュに触れない。
6. 元シフト表と勤務シフトアプリは読み取り専用連携のまま維持する。
