# README

Googleスプレッドシートに作成した月次タスク一覧表をGoogleカレンダーをトリガーに更新していく。

## 準備

### Googleカレンダー


### Googleスプレッドシート


## For development:

[clasp](https://github.com/google/clasp) で環境構築。

- `npm run push` で更新
- `clasp open-script` でエディタページを開く

スクリプト プロパティとして以下の3つを設定すること。

- CALENDAR_ID …… GoogleカレンダーのID
- SHEET_ID …… GoogleスプレッドシートのID
- WEBHOOK_URL …… Slackアプリとしてチャンネルに投稿するためのURL。
    - https://api.slack.com/apps/ からアプリのページに移動して `Incoming Webhooks` セクションからチャンネルごとに発行する
