#Slack_Send_DirectMessage
---
## 概要
スプレッドシートに記載されているメールアドレスに対して、DMを送信します。 / Send DM to email address listed in the Google Spreadsheet.

---
## 機能
Slackで同一文面のDirectMessageを一斉送信します。
送信対象はGoogleスプレッドシートに入力します。複数人のグループチャットにも対応しています。
送信する文章はGoogleドキュメントに入力します。文頭にメンションを「さん」付けで付与します。

---
## 使い方
使用するためには以下の手順が必要となります。
 * スプレッドシートの作成
 * GASファイルの適用
 * ドキュメントの作成・割り当て
 * SlackAPIの申請


---
## データのダウンロード・Googleドライブへのアップロード
1. 右上緑のアイコンの「Code」→「Download Zip」でZipファイルをダウンロードしてください。
2. Zipファイル解凍してください。
3. 「SlackDM送信.xlsx」「SlackDM送信.docx」をGoogleドライブにアップロードしてください。

---
## GASファイルの適用
1. Googleドライブ上で先ほどアップロードした「SlackDM送信.xlsx」を開いてください。
2. 「ツール」→「スクリプトエディタ」を選択してください。
3. 開いた「無題のプロジェクト」に先ほどダウンロードした「index.gs」をコピー＆ペーストしてください。

---
## ドキュメントの作成・割り当て
1. Googleドライブ上で先ほどアップロードした「SlackDM送信.docx」を開いてください。
2. Slackに投稿する文面を作成します。
3. 冒頭の「${name}」はDM送信時に相手の名前のメンション[例：@taro_tanaka]に置換されます。

---
## SlackAPIの申請
1.[SlackAPI](https://api.slack.com/apps)で以下の権限を持つAppを作成します。Appの作成は[このあたり](https://qiita.com/yuukiw00w/items/94e4495fc593cfbda45c)を読んでください。このとき「Add an OAuth Scope」が二つありますが、必ず「User Token Scopes」の方で権限を追加してください。


* chat:write
* channels:write
* groups:write
* im:write
* mpim:write
* users:read
* users:read.email

2. Slackの管理者に申請が行きますので、承認されるまで待ちます。

---
## TOKENの貼り付け
1.承認されたらトークンを取得します。「User OAuth Token」のTOKENをコピーしてください。

TOKENの例：
>xoxp--123456789012-1234567890123-1234567890123-12345678901234567890123456789012

2.「SlackDM送信.xlsx」のスクリプトエディタの8行目のダミーのTOKENと差し替えてください。

> const TOKEN = **"(太字部分を差し替え)xoxp--123456789012-1234567890123-1234567890123-12345678901234567890123456789012"**

---
## テスト起動と承認

1.「SlackDM送信.xlsx」のシートに実験送信する方のメールアドレスを入力します。Slackのメールアドレスは送る方のプロフィールに掲載されています。
2.「SlackDM送信.docx」のURLをA2セルに貼り付けます。
3.「ツール」→「マクロ」→「インポート」→「main」の「関数を追加」をクリックしてください。
4.「ツール」→「マクロ」→「main」を実行してください。初回は承認を求められるので承認してください。
5. 送信完了となると一番右のセルに「送信済み」と表示されます。
6. 送信完了となると送信結果をメッセージボックスで教えてくれます。

---
## おまけ：一定の周期で動作させる
[このあたり](https://tonari-it.com/gas-timed-driven-trigger/)を読むと良いです。
