# インストールガイド - 完全セットアップ手順

> **クイックスタート**: 基本的な使い方は [README.md](README.md) を参照してください。このガイドは詳細なセットアップ手順を提供します。

Claude Code対話型Ganttチャート自動生成システムの完全セットアップ手順を説明します。

## 📋 前提条件

- **Node.js**: v20.x以上（`node --version`で確認）
- **npm**: v10.x以上（`npm --version`で確認）
- **Claude Code CLI**: インストール済み
- **Google Cloud Platform**: アカウント作成済み
- **Google Apps Script**: 基本的な知識（推奨）

## 🚀 セットアップ手順

### 1. リポジトリのクローン

```bash
git clone https://github.com/{your-username}/claude-gantt-chart.git
cd claude-gantt-chart
```

### 2. 依存関係のインストール

```bash
npm install
```

### 3. 環境変数の設定

`.env.example`をコピーして`.env`を作成:

```bash
cp .env.example .env
```

`.env`ファイルを編集:

```env
# Google Drive設定
GOOGLE_DRIVE_FOLDER_ID=your_folder_id_here
GOOGLE_CLIENT_ID=your_client_id_here
GOOGLE_CLIENT_SECRET=your_client_secret_here
GOOGLE_REDIRECT_URI=http://localhost:19204/oauth2callback

# スプレッドシートID（Discord通知用、オプション）
GOOGLE_SPREADSHEET_ID=your_spreadsheet_id_here

# Discord通知設定（オプション）
DISCORD_WEBHOOK_URL=your_webhook_url_here

# Todoist API設定（オプション）
TODOIST_API_KEY=your_todoist_api_token_here
```

### 4. Google Cloud Platformの設定

#### 4.1 プロジェクト作成

1. [Google Cloud Console](https://console.cloud.google.com/)にアクセス
2. 新規プロジェクトを作成

#### 4.2 Google Drive API有効化

1. 「APIとサービス」→「ライブラリ」
2. 「Google Drive API」を検索して有効化
   - URL: https://console.cloud.google.com/apis/library/drive.googleapis.com

#### 4.3 OAuth 2.0認証情報の作成

1. 「APIとサービス」→「認証情報」
2. 「認証情報を作成」→「OAuth 2.0 クライアントID」
3. アプリケーションの種類: **デスクトップアプリ**
4. リダイレクトURI: `http://localhost:19204/oauth2callback`
5. クライアントIDとシークレットを`.env`に設定

#### 4.4 Google Driveフォルダ作成

1. Google Driveで新規フォルダを作成（例: `Gantt-JSON-Files`）
2. フォルダURLからフォルダIDを取得:
   - URL形式: `https://drive.google.com/drive/folders/{FOLDER_ID}`
   - `{FOLDER_ID}`の部分を`.env`の`GOOGLE_DRIVE_FOLDER_ID`に設定

### 5. Google OAuth認証

初回認証トークンを取得:

```bash
npm run gantt:auth
```

ブラウザが開き、Googleアカウントでログイン・認証します。
認証完了後、`.google-token.json`が自動生成されます（gitignoreに含まれます）。

### 6. TypeScriptのビルド

```bash
npm run build
```

### 7. Google Apps Scriptの設定

#### 7.1 スプレッドシート作成

1. [Google Sheets](https://sheets.google.com/)で新規スプレッドシートを作成
2. スプレッドシートのURLからIDを取得:
   - URL形式: `https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit`
   - `{SPREADSHEET_ID}`を`.env`の`GOOGLE_SPREADSHEET_ID`に設定（オプション）

#### 7.2 Apps Scriptプロジェクト作成

1. スプレッドシートを開く
2. 「拡張機能」→「Apps Script」
3. 新規プロジェクトが作成されます

#### 7.3 GASコードのコピー

`gas/`フォルダ内の全ファイルをApps Scriptエディタにコピー:

- `Config.gs`
- `SheetManager.gs`
- `GanttRenderer.gs`
- `Code.gs`

#### 7.4 Config.gsの設定

`Config.gs`を開き、以下を設定:

```javascript
const CONFIG = {
  DRIVE_FOLDER_ID: 'あなたのフォルダID',  // .envと同じ値
  DISCORD_WEBHOOK_URL: 'あなたのWebhook URL',  // オプション
  TODOIST_API_KEY: 'あなたのTodoist APIトークン',  // オプション
  // ...その他の設定
};
```

#### 7.5 時間トリガーの設定

1. Apps Scriptエディタで「トリガー」アイコン（⏰）をクリック
2. 「トリガーを追加」をクリック
3. 以下のように設定:
   - **実行する関数**: `checkForNewJsonFiles`
   - **イベントのソース**: `時間主導型`
   - **時間ベースのトリガーのタイプ**: `分ベースのタイマー`
   - **時間の間隔**: `1分おき`
4. 「保存」をクリック

### 8. Todoist統合（オプション）

Todoistと連携する場合:

#### 8.1 Todoist APIトークン取得

1. [Todoist設定](https://todoist.com/app/settings/integrations/developer)にアクセス
2. 「開発者」タブ→「APIトークン」をコピー
3. `.env`の`TODOIST_API_KEY`に設定
4. `Config.gs`の`TODOIST_API_KEY`にも設定

#### 8.2 Todoistトリガー設定

Apps Scriptで以下のトリガーを追加:

1. **Todoistタスク同期**:
   - 実行する関数: `syncTodoistTasks`
   - イベントのソース: `時間主導型`
   - 時間の間隔: `30分おき`

2. **期日切れタスク通知**:
   - 実行する関数: `sendOverdueTasksNotification`
   - イベントのソース: `時間主導型`
   - 時間ベースのトリガーのタイプ: `日タイマー`
   - 時刻: `午前9時～10時`

### 9. Discord通知設定（オプション）

Discord通知を有効にする場合:

1. Discordサーバーで「サーバー設定」→「連携サービス」→「ウェブフック」
2. 新しいウェブフックを作成
3. ウェブフックURLをコピー
4. `.env`の`DISCORD_WEBHOOK_URL`に設定
5. `Config.gs`の`DISCORD_WEBHOOK_URL`にも設定

### 10. clasp を使った自動デプロイ（オプション）

手動でApps Scriptエディタにコードをコピーする代わりに、claspを使ってコマンドラインから自動デプロイできます。

#### 10.1. claspのインストール

```bash
npm install -g @google/clasp
```

#### 10.2. Google認証

```bash
clasp login
```

ブラウザが開くので、Googleアカウントでログインして認証します。

#### 10.3. 新規GASプロジェクト作成

```bash
# スプレッドシート紐付け型のプロジェクトを作成
clasp create --title "Gantt Chart Generator" --type sheets
```

または、既存のGASプロジェクトを使用する場合:

```bash
# Apps Scriptエディタで「プロジェクトの設定」からスクリプトIDを取得
clasp clone <スクリプトID>
```

#### 10.4. .clasp.json 設定

プロジェクトルートに`.clasp.json`を作成:

```json
{
  "scriptId": "your-script-id-here",
  "rootDir": "./gas"
}
```

**⚠️ 注意**: `.clasp.json`はスクリプトIDを含むため、`.gitignore`で除外されています。

#### 10.5. コードをプッシュ

```bash
# gas/ディレクトリ内のコードをGASプロジェクトにアップロード
clasp push
```

#### 10.6. ブラウザで開く

```bash
# Apps Scriptエディタをブラウザで開く
clasp open
```

#### 10.7. トリガー設定

claspでコードをプッシュした後、Apps Scriptエディタで時間トリガーを設定してください（上記のセクション7を参照）。

#### clasp コマンド一覧

| コマンド | 説明 |
|---------|------|
| `clasp push` | ローカルコード → GASにアップロード |
| `clasp pull` | GASコード → ローカルにダウンロード |
| `clasp open` | ブラウザでGASエディタを開く |
| `clasp logs` | 実行ログを表示 |
| `clasp deploy` | デプロイメント作成 |

## ✅ 動作確認

### 1. Claude Codeで動作確認

```bash
# プロジェクトディレクトリで
claude code
```

Claude Codeのプロンプトで:

```
/gantt
```

スキルが正常に認識されれば成功です。

### 2. JSONファイル生成テスト

1. `/gantt`コマンドで対話形式でプロジェクト情報を入力
2. 完了後、以下を実行:

```bash
npm run gantt:save
```

3. `output/`フォルダにJSONファイルが生成されることを確認
4. Google Driveにアップロードされることを確認

### 3. Ganttチャート生成テスト

1. JSONファイルがGoogle Driveにアップロードされた後、1分以内に自動処理される
2. スプレッドシートに以下のシートが作成されることを確認:
   - プロジェクト一覧
   - 全プロジェクトタスク
   - 期日切れタスク
   - 個別Ganttチャート

## 🛠️ トラブルシューティング

### セットアップ関連

#### Q1. `npm install` でエラーが発生する

**A**: Node.jsのバージョンを確認してください。

```bash
node --version  # v20.x.x 以上を確認
npm --version   # v10.x.x 以上を確認
```

古いバージョンの場合は [Node.js公式サイト](https://nodejs.org/) から最新LTS版をインストールしてください。

#### Q2. OAuth認証がうまくいかない

**A**: 以下を段階的に確認:

1. **Google Drive APIが有効化されているか**:
   - [Google Cloud Console](https://console.cloud.google.com/apis/library/drive.googleapis.com) で確認
   - 「有効」と表示されていればOK

2. **`.env`ファイルの設定が正確か**:
   ```bash
   cat .env | grep GOOGLE_  # 設定値を確認
   ```
   - クライアントID、シークレットにスペースや改行が含まれていないか確認
   - リダイレクトURIが正確に `http://localhost:19204/oauth2callback` か確認

3. **OAuth同意画面が設定されているか**:
   - Google Cloud Consoleで「OAuth同意画面」を確認
   - テストユーザーに自分のGoogleアカウントが追加されているか確認

4. **認証トークンを再取得**:
   ```bash
   rm .google-token.json  # 古いトークンを削除
   npm run gantt:auth     # 再認証
   ```

#### Q3. JSONファイルが生成されない

**A**: 以下の順で確認:

1. **TypeScriptがビルドされているか**:
   ```bash
   npm run build
   # dist/ フォルダが生成されることを確認
   ls -la dist/
   ```

2. **ヘルパースクリプトが実行されているか**:
   ```bash
   npm run gantt:save
   ```
   エラーメッセージが表示される場合は内容を確認

3. **outputs/ フォルダの権限**:
   ```bash
   ls -la outputs/  # フォルダの存在と権限を確認
   mkdir -p outputs # 存在しない場合は作成
   ```

#### Q4. Google Driveアップロードがスキップされる

**A**: 以下を確認:

1. **`.env`に`GOOGLE_DRIVE_FOLDER_ID`が設定されているか**:
   ```bash
   grep GOOGLE_DRIVE_FOLDER_ID .env
   ```

2. **認証トークンが存在するか**:
   ```bash
   ls -la .google-token.json
   ```
   存在しない場合は `npm run gantt:auth` を実行

3. **フォルダIDが正しいか**:
   - Google DriveでフォルダのURLを確認
   - URL形式: `https://drive.google.com/drive/folders/{FOLDER_ID}`
   - `{FOLDER_ID}` 部分が `.env` と一致しているか確認

### GAS関連

#### Q5. Ganttチャートが自動生成されない

**A**: 以下を段階的に確認:

1. **時間トリガーが正しく設定されているか**:
   - Apps Scriptエディタで「トリガー」アイコン（⏰）をクリック
   - `checkForNewJsonFiles` 関数のトリガーが存在するか確認
   - 間隔が「1分おき」に設定されているか確認

2. **`Config.gs`の`DRIVE_FOLDER_ID`が正しいか**:
   ```javascript
   // Config.gs を開いて確認
   DRIVE_FOLDER_ID: 'あなたのフォルダID',  // .envと同じ値
   ```

3. **JSONファイルが正しいフォルダにアップロードされているか**:
   - Google Driveで該当フォルダを開く
   - JSONファイルが存在するか確認
   - ファイル名が `processed_` で始まっていないか確認

4. **GASの実行ログを確認**:
   - Apps Scriptエディタで「実行ログ」を開く
   - `checkForNewJsonFiles` の実行ログを確認
   - エラーメッセージがあれば内容を確認

5. **手動実行でテスト**:
   - Apps Scriptエディタで `testGenerateGantt()` 関数を実行
   - エラーが出る場合は実行ログで詳細を確認

#### Q6. 処理済みファイルが増えすぎた場合

**A**: 処理済みファイル（`processed_` で始まるファイル）は手動で削除またはアーカイブできます:

1. Google Driveで該当フォルダを開く
2. `processed_` で始まるファイルを選択（複数選択可能）
3. 別のフォルダに移動するか削除する

**推奨**: 定期的にアーカイブフォルダを作成して移動することで、メインフォルダを整理できます。

#### Q7. 特定のJSONファイルだけ処理したい

**A**: 手動実行を使用:

1. Google DriveでJSONファイルを右クリック → 「リンクを取得」
2. URLからファイルIDを取得（`/d/` と `/view` の間の文字列）
3. Apps Scriptエディタで `Code.gs` を開く
4. `processJsonFile('ファイルID')` を直接実行

### Todoist統合関連

#### Q8. Todoistタスクが同期されない

**A**: 以下を確認:

1. **`Config.gs`の`TODOIST_API_KEY`が正しく設定されているか**:
   - [Todoist設定](https://todoist.com/app/settings/integrations/developer) でAPIトークンを確認
   - `Config.gs`に正しく貼り付けられているか確認
   - APIキーにスペースや改行が含まれていないか確認

2. **トリガーが正しく設定されているか**:
   - Apps Scriptエディタで「トリガー」アイコン（⏰）をクリック
   - `syncTodoistTasks` 関数のトリガーが存在するか確認

3. **手動実行でテスト**:
   ```javascript
   // Apps Scriptエディタで実行
   testTodoistIntegration()
   ```
   実行ログでエラーメッセージを確認

4. **Inboxプロジェクトが存在するか**:
   - Todoistアプリで「Inbox」プロジェクトが存在するか確認
   - プロジェクト名が正確に「Inbox」であることを確認（大文字小文字区別）

5. **API接続をテスト**:
   ```javascript
   // Apps Scriptエディタで以下を実行
   function testTodoistAPI() {
     const url = 'https://api.todoist.com/rest/v2/projects';
     const options = {
       'method': 'get',
       'headers': {
         'Authorization': 'Bearer ' + CONFIG.TODOIST_API_KEY
       }
     };
     const response = UrlFetchApp.fetch(url, options);
     Logger.log(response.getContentText());
   }
   ```

### Discord通知関連

#### Q9. 期日切れタスク通知が送信されない

**A**: 以下を確認:

1. **Discord Webhook URLが正しく設定されているか**:
   - `Config.gs`の`DISCORD_WEBHOOK_URL`が正しいか確認
   - URLが `https://discord.com/api/webhooks/...` の形式か確認

2. **期日切れタスクが存在するか**:
   - 「期日切れタスク」シートを開く
   - 実際に期日切れタスクが存在するか確認

3. **トリガーが正しく設定されているか**:
   - `sendOverdueTasksNotification` 関数の日タイマーが設定されているか確認

4. **手動実行でテスト**:
   ```javascript
   // Apps Scriptエディタで実行
   testOverdueTasksFeature()
   ```

5. **Webhook URLをテスト**:
   ```javascript
   // Apps Scriptエディタで以下を実行
   function testDiscordWebhook() {
     const payload = {
       'content': 'テスト通知: GASからの接続確認'
     };
     const options = {
       'method': 'post',
       'contentType': 'application/json',
       'payload': JSON.stringify(payload)
     };
     UrlFetchApp.fetch(CONFIG.DISCORD_WEBHOOK_URL, options);
   }
   ```

### clasp関連

#### Q10. claspコマンドがエラーになる

**A**: 以下を確認:

1. **claspがインストールされているか**:
   ```bash
   clasp --version
   ```
   エラーが出る場合は再インストール:
   ```bash
   npm install -g @google/clasp
   ```

2. **Google認証が完了しているか**:
   ```bash
   clasp login
   ```

3. **`.clasp.json`が正しく設定されているか**:
   ```bash
   cat .clasp.json
   ```
   `scriptId`と`rootDir`が正しく設定されているか確認

4. **スクリプトIDが正しいか**:
   - Apps Scriptエディタで「プロジェクトの設定」を開く
   - スクリプトIDが`.clasp.json`と一致しているか確認

## 📚 関連ドキュメント

- **[README.md](README.md)** - プロジェクト概要、クイックスタート、基本的な使い方
- **[gas/README.md](gas/README.md)** - GAS側の詳細ドキュメント、API仕様
- **問題が発生した場合** - このページのトラブルシューティングセクションを参照

## 🎯 セットアップ完了後

セットアップが完了したら、[README.md](README.md) の「使い方」セクションを参照して、実際にプロジェクトを作成してみましょう。

```bash
# Claude Codeを起動
claude code

# 対話形式でプロジェクト作成
> /gantt
```

## 📝 ライセンス

MIT License
