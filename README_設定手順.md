# v2.2 接続設定（運用者向け）

通常ユーザーの設定画面は接続状態・接続/解除・すべて同期・最終同期だけです。個別IDを2人の端末へ入力させる方式は廃止しました。既に端末にある設定は削除しません。

1. 既存の共有Sheets・個人Sheets・共有Calendar・個人Calendarはそのまま使う。予定を削除・初期化しない。DBのevents!A:Tの列は変更しない。
2. 本体の独立したApps Scriptにbackend/Code.gsとbackend/appsscript.jsonを配置する。勤務シフト管理アプリ自体は変更しない。
3. Script PropertiesのMYLIFE_CONFIGにbackend/config.example.jsonの形式で設定する。usersは本人と恵里子の正しいGoogleメール2件。共有Sheetsは2人だけ編集可、個人Sheetsと個人Calendarは本人のみ。共有Calendarは2人だけ編集可。
4. shiftSourceIdに元の「辻岡シフト」IDを設定。同期を押す本人に元シートの閲覧権限が必要。恵里子に元の勤務表のアクセス権を付与する必要はない。本人が取り込んだ勤務は共有DBに保存される。
5. Google CloudにPWA用OAuth web clientとGitHub Pages originを登録し、Sheets/Calendar/Drive APIsを有効化する。現在この初期設定は未完了。
6. gatewayを所有者として実行・アクセス「全員」でデプロイする。全リクエストはaudience検証したGoogle access tokenと2人のallowlistで認証し、Google APIには常に呼出者のtokenを使う。所有者権限で予定操作しない。デプロイにより匿名で予定を読めるようにはならない。
7. deployment.jsに公開用clientIdとgatewayUrlのみ設定。メール・元シートID・個人DB ID・tokenはGitHubへ置かない。新backendはprotocol:3を返すことを確認する。旧backendのままではアプリは書き込みを拒否する。
8. 本体と旧勤務専用GASに時間トリガーがないことを確認。旧勤務専用GASのdisableMyLifeAutoSyncは自分のsyncMyLifeShiftsだけ解除する。他アプリのトリガーは変更しない。
9. 公開ページを再読み込みして接続。「すべて同期」のみ勤務取得・Googleイベントの送受信を実行。接続や月移動はDBを読むだけ。

「接続解除」はGoogle Calendar連携を無効化する。MyLifeのSheets保存用Googleセッションと保存済みデータは保持する。Google自体の許可の取り消しが必要ならGoogleアカウントの接続管理で行う。

勤務の時刻未入力日（開始時間/終了時間の見出し）は既存予定を保持し保留。明示的なH・公休等のみ勤務の削除候補とする。データ取得失敗・不正時刻・未知店舗では予定を消去しない。
