# Claude Code対話型Ganttチャート自動生成システム

Claude Codeとの対話でプロジェクトを設計し、JSONファイルを自動生成して、Google Apps Scriptでスプレッドシートに美しいGanttチャートを自動作成するシステムです。

## 📋 目次

- [機能概要](#機能概要)
- [システム構成](#システム構成)
- [セットアップ](#セットアップ)
- [使い方](#使い方)
- [データ仕様](#データ仕様)
- [GAS統合](#gas統合)
- [トラブルシューティング](#トラブルシューティング)

---

## 🎯 機能概要

### Phase 1: Claude Code側（このリポジトリ）

- **対話型プロジェクト設計**: `/gantt` スラッシュコマンドで段階的にプロジェクト情報を入力
- **JSON自動生成**: プロジェクト基本情報とタスク一覧を構造化JSONで出力
- **対話履歴保存**: プロジェクト作成プロセスをMarkdown形式で記録
- **Google Driveアップロード**: 生成したJSONファイルを自動アップロード

### Phase 2: GAS側（✅ 実装完了）

- **自動Ganttチャート生成**: JSONファイルからスプレッドシートに視覚的なGanttチャートを作成
- **4シート構成**: プロジェクト一覧、全プロジェクトタスク、期日切れタスク、Todoistタスク、個別Ganttチャート
- **タスク管理機能**: 進捗率、依存関係、担当者、優先度の管理
- **Todoist統合**: TodoistのInboxタスクを自動同期してスプレッドシートで一元管理
- **期日切れタスク通知**: 期日を過ぎたタスクを自動検出してDiscord通知
- **Discord通知**: 成功/エラー通知をDiscordに自動送信

---

## 🏗️ システム構成

```
185-automatic-gantt-chart-creation/
├── .claude/
│   └── commands/
│       └── gantt.md            # /gantt スラッシュコマンド定義
├── src/
│   └── gantt-helper.ts         # ヘルパースクリプト（状態管理・JSON生成・Discord通知）
├── gas/                         # Google Apps Script ファイル ✨ NEW
│   ├── Config.gs               # 設定管理（スプレッドシートID、Discord、色設定）
│   ├── SheetManager.gs         # シート操作（作成、更新、データ投入）
│   ├── GanttRenderer.gs        # Ganttチャート描画（バー、ヘッダー、スタイル）
│   ├── Code.gs                 # メインスクリプト（実行エントリーポイント）
│   └── README.md               # GAS導入手順・使い方
├── outputs/                     # JSONファイル保存先
├── docs/                        # 対話履歴ファイル保存先
├── package.json                 # 依存関係
├── tsconfig.json                # TypeScript設定
├── .env.example                 # 環境変数テンプレート
├── .gitignore
└── README.md
```

---

## 🚀 セットアップ

### 1. 依存関係のインストール

```bash
npm install
```

### 2. 環境変数の設定

`.env.example` をコピーして `.env` を作成：

```bash
cp .env.example .env
```

`.env` ファイルを編集して以下の値を設定：

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

### 3. Google OAuth認証の初回セットアップ

**Google Cloud Consoleで認証情報を取得**:

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 新規プロジェクトを作成
3. 「APIとサービス」→「ライブラリ」から以下のAPIを有効化：
   - **Google Drive API**: https://console.cloud.google.com/apis/library/drive.googleapis.com
4. 「APIとサービス」→「認証情報」
5. 「OAuth 2.0 クライアントID」を作成
   - アプリケーションの種類: デスクトップアプリ
   - リダイレクトURI: `http://localhost:19204/oauth2callback`
6. クライアントIDとクライアントシークレットを`.env`に設定

**初回認証トークン取得**:

```bash
npm run gantt:auth  # 初回OAuth認証フロー
```

認証が完了すると `.google-token.json` が生成されます（.gitignoreに含まれます）。

### 4. GAS側の時間トリガー設定

**Google Apps ScriptにDRIVE_FOLDER_IDを設定**:

1. Apps Scriptエディタを開く
2. `Config.gs`を開く
3. `DRIVE_FOLDER_ID`に`.env`と同じフォルダIDを設定：
   ```javascript
   DRIVE_FOLDER_ID: 'あなたのフォルダID',
   ```

**時間ベーストリガーを追加**:

1. Apps Scriptエディタで左側の「トリガー」アイコン（⏰）をクリック
2. 右下の「トリガーを追加」をクリック
3. 以下のように設定：
   - **実行する関数を選択**: `checkForNewJsonFiles`
   - **イベントのソースを選択**: `時間主導型`
   - **時間ベースのトリガーのタイプを選択**: `分ベースのタイマー`
   - **時間の間隔を選択**: `1分おき`
4. 「保存」をクリック

これで、1分ごとにGoogle Driveフォルダをチェックし、新しいJSONファイルが見つかると自動的にGanttチャートが生成されます。

**Todoist統合の追加設定（オプション）**:

Todoistタスクを同期する場合、以下の追加トリガーを設定してください：

1. **Todoistタスク同期トリガー**:
   - **実行する関数**: `syncTodoistTasks`
   - **イベントのソース**: `時間主導型`
   - **時間ベースのトリガーのタイプ**: `分ベースのタイマー`
   - **時間の間隔**: `30分おき` または `1時間おき`

2. **期日切れタスク通知トリガー**:
   - **実行する関数**: `sendOverdueTasksNotification`
   - **イベントのソース**: `時間主導型`
   - **時間ベースのトリガーのタイプ**: `日タイマー`
   - **時刻**: `午前9時～10時`

3. **Config.gsにTODOIST_API_KEYを設定**:
   ```javascript
   TODOIST_API_KEY: 'your_todoist_api_token_here',
   ```

   Todoist APIトークンの取得方法:
   - [Todoist設定](https://todoist.com/app/settings/integrations/developer) → 「開発者」タブ
   - 「APIトークン」をコピー

### 5. ビルド（TypeScript → JavaScript変換）

```bash
npm run build
```

---

## 📖 使い方

### 基本ワークフロー

#### 1. プロジェクト作成開始

Claude Codeで以下のコマンドを実行：

```
/gantt
```

#### 2. 対話でプロジェクト情報を入力

システムが段階的に質問するので、順次回答：

- **Phase 1: 基本情報収集**
  - プロジェクトID: `PRJ001` など
  - プロジェクト名: `新規ECサイト開発`
  - プロジェクト目的: `個人事業` or `HRteam`
  - プロジェクトジャンル: 開発/マーケティング/イベント企画/業務改善/汎用
  - プロジェクト期限: `2025-06-30`

- **Phase 2-4: タスク設計**
  - タスクの追加・修正
  - 依存関係の設定
  - 担当者・優先度・工数の入力

#### 3. 完了を宣言

すべての情報入力が完了したら：

```
完了
```

#### 4. ファイル生成・保存

ヘルパースクリプトを実行：

```bash
npm run gantt:save
```

**生成されるファイル**:

- `outputs/{プロジェクトID}_{プロジェクト名}.json` - プロジェクトデータ（JSON）
- `docs/{プロジェクトID}_{プロジェクト名}.md` - 対話履歴（Markdown）

ファイルは自動的にGoogle Driveにもアップロードされます。

#### 5. GASでGanttチャート作成

**✅ 自動実行（時間トリガー方式）**:

`npm run gantt:save` でJSONファイルをGoogle Driveにアップロード後、**1分以内に自動的にGanttチャートが生成されます**。

**動作フロー**:
1. Node.jsがJSONファイルをGoogle Driveにアップロード
2. GAS側の時間トリガー（1分ごと）がフォルダをチェック
3. 未処理のJSONファイルを検出したら自動的に処理
4. 処理済みファイル名を `processed_XXXX.json` に変更
5. Discord通知が送信される（設定している場合）

**手動実行の場合**:
1. スプレッドシートを開く
2. メニュー「Ganttチャート」→「実際のJSONから生成」
3. または、Apps Scriptエディタで`generateFromDriveFile()`関数を実行

---

## 📊 データ仕様

### プロジェクトJSON形式

```json
{
  "project_id": "PRJ001",
  "project_name": "新規ECサイト開発",
  "project_purpose": "個人事業",
  "project_type": "開発",
  "project_deadline": "2025-06-30",
  "tasks": [
    {
      "task_id": "T001",
      "task_name": "要件定義",
      "start_date": "2025-01-01",
      "end_date": "2025-01-15",
      "assignee": "山田太郎",
      "dependencies": [],
      "progress": 0,
      "priority": "★★★★★",
      "parent_task_id": null,
      "tags": ["企画・計画"],
      "estimated_hours": 40,
      "is_milestone": false
    }
  ]
}
```

### タスクフィールド詳細

| フィールド | 型 | 説明 | 例 |
|-----------|-----|------|-----|
| `task_id` | string | タスクID（自動採番） | `T001`, `T002` |
| `task_name` | string | タスク名 | `要件定義` |
| `start_date` | string | 開始日（YYYY-MM-DD） | `2025-01-01` |
| `end_date` | string | 終了日（YYYY-MM-DD） | `2025-01-15` |
| `assignee` | string | 担当者（フリーテキスト） | `山田太郎` |
| `dependencies` | array | 依存タスクID配列 | `["T001", "T002"]` |
| `progress` | number | 進捗率（0-100） | `50` |
| `priority` | string | 優先度（6段階★） | `★★★★★` |
| `parent_task_id` | string\|null | 親タスクID | `T001` or `null` |
| `tags` | array | タグ配列 | `["企画・計画", "技術"]` |
| `estimated_hours` | number | 見積工数（時間） | `40` |
| `is_milestone` | boolean | マイルストーンフラグ | `true` or `false` |

### 優先度（6段階）

- `★★★★★` - 最重要
- `★★★★☆` - 重要
- `★★★☆☆` - 中
- `★★☆☆☆` - 普通
- `★☆☆☆☆` - 低
- `☆☆☆☆☆` - 最低

### タグ分類（10種類）

| タグ名 | 色 | 説明 |
|--------|-----|------|
| 企画・計画 | 青 | プロジェクト企画、要件定義 |
| 設計 | 緑 | システム設計、UI/UX設計 |
| 開発・実装 | 黄 | プログラミング、コーディング |
| テスト・検証 | 橙 | テスト、品質保証 |
| リリース・デプロイ | 赤 | リリース作業、デプロイ |
| 運用・保守 | 紫 | 運用、保守、監視 |
| マーケティング | ピンク | 広告、プロモーション |
| 営業・商談 | 水色 | 営業活動、商談 |
| 事務・管理 | グレー | 事務作業、管理業務 |
| その他 | 白 | 上記以外 |

---

## 🔗 GAS統合

> **📚 詳細な導入手順**: [gas/README.md](gas/README.md) を参照してください

### GAS側の主な機能（✅ 実装完了）

#### 1. プロジェクト一覧シート

- 全プロジェクトのサマリー表示
- プロジェクトID、名前、目的、ジャンル、期限、進捗率
- ステータス管理（企画中/進行中/完了/保留/中止）

#### 2. 全プロジェクトタスクシート

- 全プロジェクトのタスクを一元管理
- プロジェクトID + タスクIDでフィルタリング可能
- 新規プロジェクト作成時に自動追加
- onEditトリガーで自動同期
- 手動更新ボタンあり

#### 3. 期日切れタスクシート

- 全プロジェクトから期日を過ぎたタスクを自動抽出
- 期日切れ日数を表示（例: 3日遅れ）
- プロジェクトごとにグループ化
- Discord通知機能連携
- 自動更新（全シート同期時）

#### 4. Todoistタスクシート（オプション）

- TodoistのInboxタスクを自動同期
- 表示項目: タスクID、タスク名、説明、期日、優先度、ラベル、完了状態、Todoistリンク、最終更新日時
- 優先度別の色分け表示（最高=赤、高=橙、中=黄）
- 完了タスクは緑色で表示
- クリック可能なTodoistリンク（HYPERLINK式）
- 自動更新（トリガー設定で30分〜1時間ごと）

#### 5. 個別プロジェクトGanttチャートシート

- プロジェクトごとに1シート作成
- タイムライン可視化（日単位）
- 進捗バー表示
- 依存関係の矢印表示
- 列グループ化（担当者/優先度/タグ/工数は折りたたみ）

### GASトリガー設定

- **時間ベーストリガー** ✅: 1分ごとにGoogle Driveフォルダをチェックし、未処理JSONファイルを自動処理
- **手動実行** ✅: `testGenerateGantt` または `generateFromDriveFile` 関数で即座にGanttチャート生成
- **onOpenトリガー** ✅: スプレッドシート起動時のカスタムメニュー追加
- **onEditトリガー** 🔲: シート編集時の自動同期（今後実装予定）

**時間トリガーの動作**:
- 関数: `checkForNewJsonFiles`
- 間隔: 1分おき
- 処理: 未処理のJSONファイル（`processed_` で始まらないファイル）を検出して自動処理
- 処理後: ファイル名を `processed_XXXX.json` に変更

---

## 🛠️ トラブルシューティング

### Q1. `npm install` でエラーが発生する

**A**: Node.jsのバージョンを確認してください。推奨: Node.js 20.x 以上

```bash
node --version  # v20.x.x 以上を確認
```

### Q2. Google認証がうまくいかない

**A**: 以下を確認：

1. `.env` ファイルのクライアントID・シークレットが正しいか
2. Google Drive APIが有効化されているか
3. リダイレクトURIが正確に設定されているか (`http://localhost:3000/oauth2callback`)

### Q3. JSONファイルが生成されない

**A**: 以下を確認：

1. `npm run gantt:save` を実行したか
2. TypeScriptがビルドされているか (`npm run build`)
3. ヘルパースクリプトでエラーが出ていないか

### Q4. Google Driveアップロードがスキップされる

**A**: 以下を確認：

1. `.env` に `GOOGLE_DRIVE_FOLDER_ID` が設定されているか
2. `.google-token.json` が存在するか（初回認証が完了しているか）

### Q5. Ganttチャートが自動生成されない

**A**: 以下を確認：

1. **時間トリガーが正しく設定されているか**:
   - Apps Scriptエディタで「トリガー」アイコン（⏰）をクリック
   - `checkForNewJsonFiles` 関数のトリガーが「1分おき」で設定されているか確認

2. **Config.gsのDRIVE_FOLDER_IDが正しく設定されているか**:
   - `.env`のGOOGLE_DRIVE_FOLDER_IDと同じ値が設定されているか確認

3. **JSONファイルが正しいフォルダにアップロードされているか**:
   - Google Driveで該当フォルダを開き、JSONファイルが存在するか確認

4. **GASの実行ログを確認**:
   - Apps Scriptエディタで「実行ログ」を開く
   - `checkForNewJsonFiles` の実行ログを確認
   - エラーメッセージがあれば内容を確認

### Q6. 処理済みファイルが増えすぎた場合

**A**: 処理済みファイル（`processed_` で始まるファイル）は手動で削除またはアーカイブできます：

1. Google Driveで該当フォルダを開く
2. `processed_` で始まるファイルを選択
3. 別のフォルダに移動するか削除する

### Q7. 特定のJSONファイルだけ処理したい

**A**: 手動実行を使用してください：

1. Google DriveでJSONファイルを右クリック → 「リンクを取得」
2. URLからファイルIDを取得（`/d/` と `/view` の間の文字列）
3. Apps Scriptエディタで `Code.gs` を開く
4. `processJsonFile('ファイルID')` を直接実行

### Q8. Todoistタスクが同期されない

**A**: 以下を確認：

1. **Config.gsのTODOIST_API_KEYが正しく設定されているか**:
   - [Todoist設定](https://todoist.com/app/settings/integrations/developer) でAPIトークンを確認
   - Config.gsに正しく貼り付けられているか確認

2. **トリガーが正しく設定されているか**:
   - Apps Scriptエディタで「トリガー」アイコン（⏰）をクリック
   - `syncTodoistTasks` 関数のトリガーが存在するか確認

3. **手動実行でテスト**:
   - Apps Scriptエディタで `testTodoistIntegration()` 関数を実行
   - 実行ログでエラーメッセージを確認

4. **Inboxプロジェクトが存在するか**:
   - Todoistアプリで「Inbox」プロジェクトが存在するか確認
   - プロジェクト名が正確に「Inbox」であることを確認（大文字小文字区別）

### Q9. 期日切れタスク通知が送信されない

**A**: 以下を確認：

1. **Discord Webhook URLが正しく設定されているか**:
   - Config.gsのDISCORD_WEBHOOK_URLが正しいか確認

2. **期日切れタスクが存在するか**:
   - 「期日切れタスク」シートを開いて実際に期日切れタスクが存在するか確認

3. **トリガーが正しく設定されているか**:
   - `sendOverdueTasksNotification` 関数の日タイマーが設定されているか確認

4. **手動実行でテスト**:
   - `testOverdueTasksFeature()` 関数を実行してDiscord通知をテスト

---

## 📝 ライセンス

MIT License

---

## 🙏 貢献

プルリクエストを歓迎します！バグ報告や機能要望はIssuesでお願いします。

---

## 📞 サポート

質問・問題がある場合は、Issuesでお問い合わせください。
