# MyLife Calendar v2.2 — 完全手動同期

既存のMyLife Calendarを継続するPWA。公開先：https://makoto6709.github.io/mylife-calendar/

- 画像アップロード・撮影・OCRを廃止。
- 「すべて同期」だけが勤務元の取得とGoogle Calendar双方向同期を実行する。カレンダーと設定画面に同じ操作を配置。
- 起動・復帰・月移動・接続・予定保存時にGoogle Calendarの予定を同期しない。
- 予定保存はGoogle Sheetsへ即時保存し、Googleへの変更はキューに保持。同期OFFはGoogleイベントを削除しない。
- 勤務は共有Sheetsへ登録して2人に表示。元シートはGETのみ、備考行は取得しない。勤務をアプリから変更することは禁止。
- Google Calendar連携OFFでも、MyLifeのGoogle Sheetsへの認証が有効なら勤務だけ同期できる。
- 勤務・Googleの失敗を独立して扱い、結果と最終試行日時を表示。片方の失敗で他方を中断しない。

設定・接続情報、既存DBとlocalStorageのキーは保持する。旧サーバーへの意図しない即時同期を防ぐため、protocol 3が確認できるまで新アプリからの書き込みを停止する。

## 現在の稼働制限

公開UIとソースは更新するが、Google OAuth clientId・本体gatewayUrlの初期設定は未完了。Google実接続と2アカウントの実機検証は未完了。体験モードを実同期成功として表示しない。勤務用の旧独立GASは本体gatewayとは別物で、PWAから呼ばない。

構成：core.js（日付・起床計算）、sync.js（手動実行の統括）、app.js（UIとアダプター）、backend/Code.gs（認証・保存・外部同期）。手順はREADME_設定手順.md、検証はdocs/検証状況.md。
