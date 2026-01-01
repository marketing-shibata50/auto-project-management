# Google Apps Script (GAS) 導入手順

このディレクトリには、スプレッドシートに自動的にGanttチャートを生成するGASコードが含まれています。

## 📋 前提条件

- Googleアカウント
- スプレッドシート作成済み（ID: `1y7U-3hVfdubQPh-H39k-5bvuHF7cOkHkOuA121GtxCI`）
- Discord Webhook URL設定済み

## 🚀 セットアップ手順

### 1. Google Apps Scriptプロジェクトを開く

1. スプレッドシートを開く: https://docs.google.com/spreadsheets/d/1y7U-3hVfdubQPh-H39k-5bvuHF7cOkHkOuA121GtxCI
2. メニューから「拡張機能」→「Apps Script」をクリック
3. 新しいタブでGASエディタが開きます

### 2. スクリプトファイルをアップロード

GASエディタで以下のファイルを作成します:

#### **Config.gs**
1. 左側のファイル一覧で「+」ボタンをクリック → 「スクリプト」を選択
2. ファイル名を `Config.gs` に変更
3. `gas/Config.gs` の内容をコピーして貼り付け

#### **SheetManager.gs**
1. 同様に新しいスクリプトを作成
2. ファイル名を `SheetManager.gs` に変更  
3. `gas/SheetManager.gs` の内容をコピーして貼り付け

#### **GanttRenderer.gs**
1. 同様に新しいスクリプトを作成
2. ファイル名を `GanttRenderer.gs` に変更
3. `gas/GanttRenderer.gs` の内容をコピーして貼り付け

#### **Code.gs**
1. 既存の `Code.gs` ファイルの内容を削除
2. `gas/Code.gs` の内容をコピーして貼り付け

### 3. 権限の承認

1. GASエディタ上部の「実行」ボタンの隣にあるプルダウンから `testGenerateGantt` 関数を選択
2. 「実行」ボタンをクリック
3. 初回実行時に権限の承認が必要です:
   - 「権限を確認」をクリック
   - Googleアカウントを選択
   - 「詳細」→「プロジェクト名（安全ではないページ）に移動」をクリック
   - 「許可」をクリック

## 🧪 テスト実行

### テストデータで動作確認

1. GASエディタで `testGenerateGantt` 関数が選択されていることを確認
2. 「実行」ボタンをクリック
3. 実行ログを確認:
   - GASエディタ下部の「実行ログ」をクリック
   - ✓マークが表示されれば成功

### 確認事項

実行後、スプレッドシートで以下を確認:

#### ✅ プロジェクト一覧シート
- シート「プロジェクト一覧」が作成されている
- プロジェクトID `0009` の情報が表示されている
- 進捗率、総工数が自動計算されている

#### ✅ 全プロジェクトタスクシート
- シート「全プロジェクトタスク」が作成されている
- 36個のタスクが一覧表示されている

#### ✅ Ganttチャート
- シート「Gantt_0009 - セミナー振り分けシステム」が作成されている
- タスク情報が左側に表示されている
- 右側に月/日のヘッダーが表示されている
- タスクごとにカラフルなバーが表示されている
- 担当者/優先度/タグ/工数の列がデフォルトで折りたたまれている

#### ✅ Discord通知
- Discordチャンネルに成功通知が届いている
- プロジェクト名、ID、タスク数、スプレッドシートURLが表示されている

## 📊 実際のデータで実行

### 方法1: JSONファイルをGoogle Driveにアップロード

1. プロジェクトのJSONファイル（例: `0009_セミナー振り分けシステム.json`）をGoogle Driveにアップロード
2. ファイルを右クリック → 「リンクを取得」→ ファイルIDをコピー
   - URL例: `https://drive.google.com/file/d/{ファイルID}/view`
3. GASエディタで `processJsonFile` 関数を編集:
   ```javascript
   processJsonFile('ここにファイルIDを貼り付け');
   ```
4. `processJsonFile` を実行

### 方法2: JSONデータを直接Code.gsに貼り付け

1. JSONファイルの内容をコピー
2. `Code.gs` の `testGenerateGantt` 関数内の `testData` 変数を更新:
   ```javascript
   const testData = {
     // ここにJSONの内容を貼り付け
   };
   ```
3. `testGenerateGantt` を実行

## 🔄 自動化

### 期日切れタスクのDiscord通知

毎日自動的に期日切れタスクをDiscordに通知する機能を設定できます。

#### トリガー設定手順（朝7時と夜20時に通知）

1. GASエディタ左側の「トリガー」（時計アイコン）をクリック
2. 「トリガーを追加」をクリック
3. 以下を設定（朝7時の通知）:
   - 実行する関数: `sendOverdueTasksNotification`
   - イベントのソース: 時間主導型
   - 時間ベースのトリガーのタイプ: 日付ベースのタイマー
   - 時刻: 午前7時～8時
4. 「保存」をクリック
5. もう一度「トリガーを追加」をクリックして夜20時の通知を設定:
   - 実行する関数: `sendOverdueTasksNotification`
   - イベントのソース: 時間主導型
   - 時間ベースのトリガーのタイプ: 日付ベースのタイマー
   - 時刻: 午後8時～9時
6. 「保存」をクリック

#### 通知内容

- プロジェクトごとにグループ化された期日切れタスク一覧
- タスク名、期日、遅延日数、進捗率、担当者
- スプレッドシートへの直接リンク

### バーンダウンチャート自動記録

毎日自動的にバーンダウンデータを記録する機能を設定できます。

#### 🚀 簡単セットアップ（推奨）

GASエディタで以下の関数を1回実行するだけでOK：

1. 関数選択ドロップダウンから `setupDailyBurndownTrigger` を選択
2. 「実行」ボタンをクリック
3. 実行ログに「✅ トリガーを作成しました！」と表示されたら完了

これで毎日午前9時に自動記録が開始されます。

**停止したい場合**: `removeDailyBurndownTrigger` を実行

#### 手動セットアップ（従来の方法）

<details>
<summary>クリックして展開</summary>

1. GASエディタ左側の「トリガー」（時計アイコン）をクリック
2. 「トリガーを追加」をクリック
3. 以下を設定:
   - 実行する関数: `recordDailyBurndownData`
   - イベントのソース: 時間主導型
   - 時間ベースのトリガーのタイプ: 日付ベースのタイマー
   - 時刻: 午前9時～10時
4. 「保存」をクリック

</details>

#### 記録内容

- プロジェクトごとの予定進捗率と実績進捗率
- 完了タスク数、残タスク数
- ベロシティ（過去7日間の平均完了タスク数/日）
- 完了予測日
- データは BurndownData シートに記録されます
- 90日より古いデータは自動削除されます

### 全シート自動同期

毎日自動的に全Ganttシートからデータを読み取り、プロジェクト一覧・全タスク・期日切れタスクを同期する機能を設定できます。

#### 🚀 簡単セットアップ（推奨）

GASエディタで以下の関数を1回実行するだけでOK：

1. 関数選択ドロップダウンから `setupSyncAllSheetsTrigger` を選択
2. 「実行」ボタンをクリック
3. 実行ログに「✅ トリガーを作成しました！」と表示されたら完了

これで毎日午前10時に自動同期が開始されます。

**停止したい場合**: `removeSyncAllSheetsTrigger` を実行

#### 同期内容

- 全Ganttシート（`XXXX_プロジェクト名`形式）からデータを読み取り
- プロジェクト一覧シートを更新
- 全プロジェクトタスクシートを更新
- 期日切れタスク一覧シートを更新
- Ganttシートのステータスから進捗率を自動更新

### Google Driveフォルダ監視（今後実装予定）

Google Driveのフォルダを監視して、新しいJSONファイルが追加されたら自動的にGanttチャートを生成する機能は今後実装予定です。

## 📝 カスタマイズ

### タグの色を変更

`Config.gs` の `TAG_COLORS` オブジェクトを編集:

```javascript
TAG_COLORS: {
  '開発': '#4A90E2',      // 青
  'デザイン': '#E24A90',  // ピンク
  'テスト': '#90E24A',    // 緑
  '調査': '#E2904A',      // オレンジ
  'レビュー': '#904AE2',  // 紫
  'デフォルト': '#999999' // グレー
}
```

### 優先度の色を変更

`Config.gs` の `PRIORITY_COLORS` オブジェクトを編集:

```javascript
PRIORITY_COLORS: {
  '高': '#FF6B6B',   // 赤
  '中': '#FFA726',   // オレンジ
  '低': '#66BB6A'    // 緑
}
```

### Ganttチャートの表示期間を変更

`Config.gs` の `GANTT.DAYS_TO_SHOW` を編集:

```javascript
GANTT: {
  START_COLUMN: 10,
  HEADER_ROW: 1,
  DATA_START_ROW: 2,
  DAYS_TO_SHOW: 180,  // この数値を変更（日数）
  CELL_WIDTH: 30
}
```

## 🐛 トラブルシューティング

### エラー: "ReferenceError: CONFIG is not defined"

- `Config.gs` ファイルが正しく作成されているか確認
- ファイル名が正確に `Config.gs` になっているか確認
- GASエディタで「保存」ボタンをクリック

### エラー: "Exception: スプレッドシートが見つかりません"

- `Config.gs` の `SPREADSHEET_ID` が正しいか確認
- スプレッドシートへのアクセス権限があるか確認

### Ganttバーが表示されない

- タスクに `start_date` と `end_date` が設定されているか確認
- 日付形式が `YYYY-MM-DD` になっているか確認

### Discord通知が届かない

- `Config.gs` の `DISCORD_WEBHOOK_URL` が正しいか確認
- Webhook URLが有効か確認

## 📚 ファイル構成

```
gas/
├── Config.gs           # 設定ファイル（スプレッドシートID、Discord、色設定など）
├── SheetManager.gs     # シート操作クラス（作成、更新、データ投入）
├── GanttRenderer.gs    # Ganttチャート描画クラス（バー、ヘッダー、スタイル）
├── Code.gs             # メインスクリプト（実行エントリーポイント）
└── README.md           # このファイル
```

## 🎯 次のステップ

1. ✅ GASスクリプトのセットアップ完了
2. ✅ テストデータでGanttチャート生成成功
3. 🔲 実際のプロジェクトデータで動作確認
4. 🔲 Google Drive監視機能の実装
5. 🔲 エラーハンドリングの強化
6. 🔲 パフォーマンス最適化

## 💡 ヒント

- **スプレッドシートURL**: https://docs.google.com/spreadsheets/d/1y7U-3hVfdubQPh-H39k-5bvuHF7cOkHkOuA121GtxCI
- **実行ログの確認**: GASエディタ → 「実行ログ」をクリック
- **スクリプトの保存**: `Ctrl+S` (Windows) / `Cmd+S` (Mac)
- **関数の実行**: `Ctrl+R` (Windows) / `Cmd+R` (Mac)

---

## 🔧 clasp統合（開発者向け）

### 概要

claspを使用すると、ローカル環境からGoogle Apps Scriptコードを管理できます。Git経由でのバージョン管理や、エディタでの快適な開発が可能になります。

### 初回セットアップ

#### 1. clasp認証

```bash
npm run gas:login
```

- ブラウザが開き、Googleアカウントでの認証が求められます
- Apps Script APIへのアクセス許可を承認します
- 認証情報は `~/.clasprc.json` に保存されます

**注意**: 事前に Apps Script APIを有効化してください
- URL: https://script.google.com/home/usersettings
- "Google Apps Script API" をONにする

#### 2. GASプロジェクトから既存コードをpull

```bash
npm run gas:pull
```

- GASプロジェクトからローカルにコードをダウンロード
- 既存のgas/ディレクトリのファイルと統合されます
- `appsscript.json` も自動的にダウンロードされます

### 開発ワークフロー

#### 通常の開発手順

1. **ローカルでgas/内のファイルを編集**
   ```bash
   vim gas/SheetManager.gs
   # またはVS Codeなどのエディタで編集
   ```

2. **GASにpush**
   ```bash
   npm run gas:push
   ```
   - ローカルの変更がGASプロジェクトに反映されます
   - 依存関係順（Config → SheetManager → GanttRenderer → Code）でアップロードされます

3. **GASエディタで確認**
   ```bash
   npm run gas:open
   ```
   - ブラウザでGASエディタが開きます
   - アップロードされたコードを確認できます

4. **ログ確認**
   ```bash
   npm run gas:logs
   ```
   - GASの実行ログがターミナルに表示されます

#### GASエディタで編集した場合

GASエディタで直接編集した内容をローカルに反映する場合:

```bash
# GASからローカルにpull
npm run gas:pull

# 差分確認
git diff gas/

# 問題なければコミット
git add gas/
git commit -m "chore: sync from GAS editor"
```

### 利用可能なコマンド

#### 基本コマンド

- `npm run gas:push` - ローカル → GAS（コードをアップロード）
- `npm run gas:pull` - GAS → ローカル（コードをダウンロード）
- `npm run gas:open` - ブラウザでGASエディタを開く
- `npm run gas:logs` - 実行ログ表示

#### 開発支援コマンド

- `npm run gas:push:watch` - ファイル変更を監視して自動push（開発時のみ）
- `npm run gas:logs:watch` - ログをリアルタイム監視

#### 認証管理

- `npm run gas:login` - clasp認証
- `npm run gas:logout` - clasp認証解除

#### バージョン管理

- `npm run gas:version "v1.1.0 - New feature"` - GASバージョン作成
- `npm run gas:versions` - バージョン一覧表示

### 推奨ワークフロー

```bash
# 1. ローカルでファイル編集
vim gas/Code.gs

# 2. GASにpush
npm run gas:push

# 3. GASエディタで動作確認
npm run gas:open

# 4. ログ確認（関数実行後）
npm run gas:logs

# 5. Gitにコミット
git add gas/Code.gs
git commit -m "feat: add new feature"
git push origin main
```

### 自動同期モード（開発時）

開発中にファイル変更を自動的にGASに反映する場合:

```bash
# ターミナル1: ファイル監視 + 自動push
npm run gas:push:watch

# ターミナル2: ログ監視
npm run gas:logs:watch

# Ctrl+C で監視を停止
```

### トラブルシューティング

#### エラー: `User has not enabled the Apps Script API`

**原因**: Apps Script APIが有効化されていない

**解決方法**:
1. https://script.google.com/home/usersettings を開く
2. "Google Apps Script API" をONにする
3. `npm run gas:login` を再実行

#### エラー: `Could not find script`

**原因**: `.clasp.json` のスクリプトIDが間違っている

**解決方法**:
1. Apps Scriptエディタ → プロジェクト設定 → スクリプトIDを確認
2. `.clasp.json` のスクリプトIDを修正
3. `npm run gas:pull` を再実行

#### 認証リセット

認証に問題がある場合:

```bash
# logout
npm run gas:logout

# 認証ファイル削除（念のため）
rm ~/.clasprc.json

# 再login
npm run gas:login
```

#### pushが失敗する

**原因**: GASファイルに構文エラーがある

**解決方法**:
1. エラーメッセージでファイル名と行番号を確認
2. 該当ファイルを修正
3. 再度push

```bash
# 修正後
npm run gas:push
```

### プロジェクト構成

```
185-automatic-gantt-chart-creation/
├── .clasp.json              # clasp設定（スクリプトID、rootDir）
├── package.json             # npm scripts（gas:*コマンド）
├── gas/                     # GASファイル
│   ├── appsscript.json     # GASプロジェクトメタデータ
│   ├── .claspignore        # clasp除外ファイル設定
│   ├── Config.gs
│   ├── SheetManager.gs
│   ├── GanttRenderer.gs
│   ├── Code.gs
│   └── README.md           # このファイル
└── src/                     # TypeScriptファイル（ローカル実行用）
```

### セキュリティ考慮事項

- `.clasp.json` はGitにコミット済み（スクリプトIDは秘匿情報ではない）
- `~/.clasprc.json` はホームディレクトリに保存（プロジェクト外）
- `Config.gs` に DISCORD_WEBHOOK_URL と TODOIST_API_KEY がハードコードされています
  - 今後の改善: PropertiesService で環境変数化を検討

### Tips

- **ファイルのpush順序**: Config → SheetManager → GanttRenderer → Code（依存関係順）
- **初回pullは慎重に**: 既存のgas/ディレクトリを事前にバックアップ推奨
- **GASバージョン**: リリース時に `npm run gas:version` でバージョン作成を推奨
- **ログ確認**: `npm run gas:logs` で直近15件のログを確認可能
