# Bijiris

QRコードで回答できるアンケートアプリです。  
お客さんは共通QRコードからアンケートタイトルを選んで回答し、管理者はフォーム作成、回答一覧、回答者ごとの履歴確認、計測管理、運用設定をブラウザで行えます。

## 保存場所

- データベース: `data/bijiris.db`
- アップロード画像: `data/uploads/`
- バックアップ: `data/backups/`
- 設定ファイル: `data/config.json`

## 起動

通常起動:

```bash
cd /Users/Kenmo/bijiris
python3 server.py
```

運用用スクリプト:

```bash
cd /Users/Kenmo/bijiris
./scripts/start_bijiris.sh
```

管理画面:

```text
http://127.0.0.1:8123/
```

## GitHub から常設公開する

このアプリは Python サーバー、SQLite、画像アップロードを使うため、GitHub Pages のような静的ホスティングでは動きません。  
GitHub にコードを置き、Render の Web Service と永続ディスクで公開する構成を推奨します。

追加済みファイル:

- Render Blueprint: `render.yaml`
- Git 管理除外: `.gitignore`
- 環境変数例: `.env.example`

基本手順:

1. `bijiris` フォルダだけを新しい GitHub リポジトリへ push
2. Render で `New +` -> `Blueprint` を選び、その GitHub リポジトリを接続
3. `render.yaml` の内容で Web Service と永続ディスクを作成
4. Render が発行する `https://xxxx.onrender.com` で動作確認
5. 必要なら管理画面の `設定` タブで公開URLを独自ドメインへ変更

注意:

- 初回デプロイ直後は管理者パスワードが `bijiris-admin` なので、すぐ変更してください
- `render-data` ディスクに DB と画像が保存されます
- リポジトリへ `data/` の本番データは入れないでください
- 永続ディスクは Render の有料 Web Service が必要です

## 管理画面でできる運用設定

`設定` タブから次を操作できます。

- 公開ベースURLの保存
- 管理者パスワード変更
- バックアップ作成とダウンロード
- 現在の運用状態確認

初期パスワードは `bijiris-admin` です。初回ログイン後に変更してください。

## 公開URL

固定の公開URLや独自ドメインを使う場合は、管理画面の `設定` タブで保存できます。  
保存した値は `data/config.json` に入ります。

環境変数を使う場合:

```bash
BIJIRIS_PUBLIC_BASE_URL=https://example.com python3 server.py
```

環境変数がある場合は、管理画面で保存した公開URLより環境変数が優先されます。

## お客様アプリとして使う

回答者画面 `/f/` は PWA として動作します。

- Android: Chrome で開くと `ホーム画面に追加` を出せます
- iPhone / iPad: Safari の共有メニューから `ホーム画面に追加` を使います
- 回答者画面にインストール案内カードを表示しています
- `最新版に更新` ボタンと再起動で最新版へ更新できます

## バックアップ

管理画面の `設定` タブからバックアップを作成できます。  
ローカルで手動作成する場合:

```bash
cd /Users/Kenmo/bijiris
./scripts/backup_bijiris.sh
```

生成物:

```text
data/backups/bijiris-backup-YYYYMMDD-HHMMSS.tar.gz
```

## LaunchAgent

Macで常駐運用する場合はテンプレートを使えます。

テンプレート:

```text
scripts/com.bijiris.server.plist
```

必要に応じて次へ配置してください。

```text
~/Library/LaunchAgents/com.bijiris.server.plist
```

固定URLの Cloudflare Tunnel を常駐する場合:

```text
scripts/com.bijiris.tunnel.plist
```

## 固定URLの Cloudflare Tunnel

`trycloudflare.com` は一時URLなので、本番運用では名前付きトンネルに切り替えてください。

下準備済みのファイル:

- テンプレート: `cloudflared/config.template.yml`
- 起動スクリプト: `scripts/start_cloudflared_named_tunnel.sh`
- LaunchAgent: `scripts/com.bijiris.tunnel.plist`

切り替え手順の概要:

1. `cloudflared tunnel login` で Cloudflare にログイン
2. `cloudflared tunnel create bijiris` で名前付きトンネル作成
3. `cloudflared tunnel route dns bijiris app.example.com` で固定ホスト名を割り当て
4. `cloudflared/config.template.yml` を `cloudflared/config.yml` にコピーして、`tunnel`, `credentials-file`, `hostname` を実値へ変更
5. 管理画面の `設定` タブか `data/config.json` の `public_base_url` を固定URLへ変更
6. `scripts/start_cloudflared_named_tunnel.sh` でトンネル起動

固定URLへ切り替えた後は、QRコードとホーム画面ショートカットをそのURL基準で配り直してください。

## 注意点

- アンケートから新しくアップロードされる画像は、アプリ内保存されるため管理画面で表示できます。
- 過去にGoogle DriveのURL参照で取り込んだ画像は、元URLの共有設定に依存します。安定運用する場合はアプリ内保存へ移行してください。
- `trycloudflare.com` のURLは一時URLです。常設運用では固定ドメインを推奨します。
