# インストールガイド

Claude Code対話型Ganttチャート自動生成システムのセットアップ手順を説明します。

## 📋 前提条件

- **Node.js**: v20.x以上
- **npm**: v10.x以上
- **Claude Code CLI**: インストール済み
- **Google Cloud Platform**: アカウント作成済み
- **Google Apps Script**: 基本的な知識

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

### Q1. OAuth認証がうまくいかない

**A**: 以下を確認:
1. Google Drive APIが有効化されているか
2. `.env`のクライアントID・シークレットが正確か
3. リダイレクトURIが正確に設定されているか

### Q2. JSONファイルが生成されない

**A**: 以下を確認:
1. `npm run build`を実行したか
2. TypeScriptのコンパイルエラーがないか
3. `npm run gantt:save`を実行したか

### Q3. Ganttチャートが自動生成されない

**A**: 以下を確認:
1. 時間トリガーが正しく設定されているか
2. `Config.gs`の`DRIVE_FOLDER_ID`が正しいか
3. Apps Scriptの実行ログでエラーがないか
4. JSONファイルが正しいフォルダにアップロードされているか

### Q4. Todoistタスクが同期されない

**A**: 以下を確認:
1. `TODOIST_API_KEY`が正しく設定されているか
2. Todoistトリガーが正しく設定されているか
3. `testTodoistIntegration()`関数で手動テストする

## 📚 次のステップ

- [README.md](README.md) - 詳細な使い方
- [gas/README.md](gas/README.md) - GAS側の詳細ドキュメント
- [DESIGN.md](DESIGN.md) - システム設計ドキュメント

## 🆘 サポート

問題が発生した場合:
1. [Issues](https://github.com/{your-username}/claude-gantt-chart/issues)で報告
2. トラブルシューティングセクションを確認
3. Apps Scriptの実行ログを確認

## 📝 ライセンス

MIT License
