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

## 🚀 クイックスタート

**初めての方**: 詳細なセットアップ手順は [INSTALL.md](INSTALL.md) をご覧ください。

### 3ステップでスタート

#### ステップ1: 依存関係のインストール

```bash
npm install
npm run build
```

#### ステップ2: 環境変数の設定

```bash
cp .env.example .env
# .envファイルを編集してGoogle Drive設定を追加
```

**最小限の設定項目**:
```env
GOOGLE_DRIVE_FOLDER_ID=your_folder_id_here
GOOGLE_CLIENT_ID=your_client_id_here
GOOGLE_CLIENT_SECRET=your_client_secret_here
```

> **📝 詳細**: Google Cloud設定、OAuth認証の完全ガイドは [INSTALL.md](INSTALL.md) を参照

#### ステップ3: 初回認証とGantt作成

```bash
# Google認証（初回のみ）
npm run gantt:auth

# Claude Codeでプロジェクト作成
claude code
> /gantt

# JSONファイル生成・アップロード
npm run gantt:save
```

> **📝 GAS連携**: スプレッドシート連携とトリガー設定は [INSTALL.md](INSTALL.md) を参照

---

## 📖 使い方

### 基本ワークフロー

#### 1. Claude Codeで対話形式で作成

```bash
claude code
> /gantt
```

システムが段階的に質問します：
- プロジェクト基本情報（ID、名前、目的、期限）
- タスク設計（タスク名、日程、担当者、依存関係）

#### 2. JSONファイル生成

対話完了後、「完了」と入力してから：

```bash
npm run gantt:save
```

これで以下が自動生成されます：
- `outputs/{プロジェクトID}_{プロジェクト名}.json`
- `docs/{プロジェクトID}_{プロジェクト名}.md`（対話履歴）

#### 3. Ganttチャート自動生成

Google Driveにアップロード後、**1分以内に自動的にスプレッドシートが更新されます**。

> **📝 詳細**: GASトリガー設定、手動実行方法は [INSTALL.md](INSTALL.md) を参照

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

### 生成される5つのシート

1. **プロジェクト一覧**: 全プロジェクトのサマリー（ステータス、進捗率）
2. **全プロジェクトタスク**: タスクを一元管理、自動同期
3. **期日切れタスク**: 遅延タスクの自動抽出、Discord通知連携
4. **Todoistタスク**（オプション）: Inbox自動同期、優先度色分け
5. **個別Ganttチャート**: プロジェクトごとのタイムライン可視化

### 自動処理の仕組み

GAS側の時間トリガー（1分間隔）がGoogle Driveフォルダを監視し、新しいJSONファイルを自動処理します。

> **📝 詳細**: GAS導入手順、clasp自動デプロイ、トリガー設定は [INSTALL.md](INSTALL.md) および [gas/README.md](gas/README.md) を参照

---

## 🛠️ よくある質問

### Q1. JSONファイルが生成されない

**A**: `npm run build` → `npm run gantt:save` を実行してください。

### Q2. Google認証エラーが発生する

**A**: `.env`ファイルのクライアントID/シークレット、リダイレクトURIを確認してください。

### Q3. Ganttチャートが自動生成されない

**A**: GAS側の時間トリガー設定と`Config.gs`のフォルダIDを確認してください。

> **📝 詳しいトラブルシューティング**: [INSTALL.md](INSTALL.md) を参照

---

## 📝 ライセンス

MIT License

---

## 🙏 貢献

プルリクエストを歓迎します！バグ報告や機能要望はIssuesでお願いします。

---

## 📞 サポート

質問・問題がある場合は、Issuesでお問い合わせください。
