/**
 * 設定管理
 */

const CONFIG = {
  // スプレッドシートID
  SPREADSHEET_ID: '1y7U-3hVfdubQPh-H39k-5bvuHF7cOkHkOuA121GtxCI',

  // Google DriveフォルダID（JSONファイルアップロード先）
  // .envのGOOGLE_DRIVE_FOLDER_IDと同じ値を設定してください
  DRIVE_FOLDER_ID: '1UEOEdzxO6yLlwgDS7NF4EcdicHTTAgTT',

  // Discord Webhook URL
  DISCORD_WEBHOOK_URL: 'https://discord.com/api/webhooks/1450654318612185221/0wfVPnUJJ8uvfD_Y4nZ901o9ezqx2H2TGDwN27s9-Yge4nOAQ4gr96UBAlH7PyMCAhcS',

  // Todoist API Token
  TODOIST_API_KEY: 'dcbc9cfc8aa52fce62a50a09e3ccecd0200faf15',

  // シート名
  SHEET_NAMES: {
    PROJECT_LIST: 'プロジェクト一覧',
    ALL_TASKS: '全プロジェクトタスク',
    OVERDUE_TASKS: '期日切れタスク',
    TODOIST_TASKS: 'Todoistタスク',
    GANTT_PREFIX: 'Gantt_' // プロジェクトIDが後ろに付く
  },
  
  // タグごとの色設定
  TAG_COLORS: {
    '企画・計画': '#4A90E2',
    '設計': '#66BB6A',
    '開発・実装': '#FDD835',
    'テスト・検証': '#FF9800',
    'リリース・デプロイ': '#E53935',
    '運用・保守': '#AB47BC',
    'マーケティング': '#26C6DA',
    '営業・商談': '#26C6DA',
    '事務・管理': '#78909C',
    'その他': '#BDBDBD',
    'デフォルト': '#999999'
  },

  // バーンダウンチャートの色設定
  BURNDOWN_COLORS: {
    EXPECTED_LINE: '#4A90E2',      // 予定進捗ライン（青）
    ACTUAL_LINE: '#66BB6A',        // 実績進捗ライン（緑）
    WARNING_LINE: '#FF9800',       // 警告ライン（オレンジ）
    DANGER_ZONE: '#FFCDD2',        // 遅延ゾーン背景（薄い赤）
    SAFE_ZONE: '#C8E6C9'           // 順調ゾーン背景（薄い緑）
  },

  // 優先度の色設定
  PRIORITY_COLORS: {
    '高': '#FF6B6B',
    '中': '#FFA726',
    '低': '#66BB6A'
  },
  
  // Ganttチャートの設定
  GANTT: {
    START_COLUMN: 13, // M列から開始（A=1, M=13） - L列に実績工数を追加
    HEADER_ROW: 1,
    DATA_START_ROW: 2,
    DAYS_TO_SHOW: 180, // デフォルト表示日数
    CELL_WIDTH: 30 // セルの幅（ピクセル）
  }
};

/**
 * Discord通知を送信
 */
function sendDiscordNotification(message, isError = false) {
  const payload = {
    embeds: [{
      title: isError ? '❌ GAS実行エラー' : '✅ GAS実行完了',
      description: message,
      color: isError ? 0xFF0000 : 0x00FF00,
      timestamp: new Date().toISOString(),
      footer: {
        text: 'Gantt Chart Generator (GAS)'
      }
    }]
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    UrlFetchApp.fetch(CONFIG.DISCORD_WEBHOOK_URL, options);
  } catch (error) {
    Logger.log('Discord通知エラー: ' + error);
  }
}
