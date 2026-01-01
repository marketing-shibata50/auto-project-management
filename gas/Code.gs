/**
 * Ganttチャート自動生成システム - メインスクリプト
 *
 * このスクリプトはJSONファイルからGanttチャートを自動生成します
 */

/**
 * スプレッドシートを開いたときにカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 ガントチャート')
    .addItem('🔄 すべて更新', 'menuUpdateAll')
    .addItem('📂 Driveから再読み込み', 'menuReloadFromDrive')
    .addSeparator()
    .addItem('📊 ガントチャート更新', 'menuUpdateGantt')
    .addItem('📋 全タスク一覧更新', 'menuUpdateAllTasks')
    .addItem('⚠️ 期日切れ一覧更新', 'menuUpdateOverdue')
    .addSeparator()
    .addItem('🔄 バッチ同期 (次の5シート)', 'menuSyncNextBatch')
    .addItem('🎨 期日切れ色付け', 'menuUpdateOverdueColors')
    .addItem('📝 Todoist同期', 'menuSyncTodoist')
    .addToUi();
}

/**
 * メニュー: すべて更新
 */
function menuUpdateAll() {
  const ui = SpreadsheetApp.getUi();
  try {
    menuUpdateGantt();
    menuUpdateAllTasks();
    menuUpdateOverdue();
    ui.alert('✅ 更新完了', 'すべてのシートを更新しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ エラー', `更新中にエラーが発生しました:\n${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * メニュー: Driveから再読み込み
 * Google DriveのJSONファイルを手動で読み込み、スプレッドシートに反映
 */
function menuReloadFromDrive() {
  const ui = SpreadsheetApp.getUi();

  // 確認ダイアログ
  const response = ui.alert(
    '📂 Driveから再読み込み',
    'Google Drive上のJSONファイルを読み込み、\nスプレッドシートに反映します。\n\n処理を実行しますか？',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  try {
    // checkForNewJsonFiles を手動実行
    checkForNewJsonFiles();

    ui.alert(
      '✅ 読み込み完了',
      'JSONファイルの読み込みが完了しました。\n\n詳細はログで確認できます。',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert(
      '❌ エラー',
      `読み込み中にエラーが発生しました:\n${e.message}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * メニュー: ガントチャート更新
 * 現在アクティブなシートがガントチャートシートの場合、そのシートを更新
 */
function menuUpdateGantt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();

  // プロジェクト一覧シートからプロジェクトIDを取得
  const projectListSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PROJECT_LIST);
  if (!projectListSheet) {
    throw new Error('プロジェクト一覧シートが見つかりません。');
  }

  // ガントチャートシートかどうか判定（プロジェクト一覧からリンクを確認）
  const data = projectListSheet.getDataRange().getValues();
  let projectId = null;

  for (let i = 1; i < data.length; i++) {
    const rowProjectId = data[i][0];
    const expectedSheetName = `${rowProjectId}_`;
    if (sheetName.startsWith(expectedSheetName)) {
      projectId = rowProjectId;
      break;
    }
  }

  if (!projectId) {
    // アクティブシートがガントチャートでない場合、選択ダイアログを表示するか最初のプロジェクトを更新
    SpreadsheetApp.getUi().alert('⚠️ 注意', '現在のシートはガントチャートシートではありません。\nガントチャートシートを選択してから実行してください。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // DriveからJSONを再読み込みしてガントチャートを更新
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (fileName.includes(projectId) && fileName.endsWith('.json')) {
      const content = file.getBlob().getDataAsString();
      const projectData = JSON.parse(content);

      const renderer = new GanttRenderer(activeSheet, projectData);
      renderer.render();
      Logger.log(`✓ ガントチャート「${sheetName}」を更新しました`);
      return;
    }
  }

  throw new Error(`プロジェクトID ${projectId} のJSONファイルが見つかりません。`);
}

/**
 * メニュー: 全タスク一覧更新
 */
function menuUpdateAllTasks() {
  const sheetManager = new SheetManager();
  const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (fileName.endsWith('.json')) {
      const content = file.getBlob().getDataAsString();
      const projectData = JSON.parse(content);
      sheetManager.updateAllTasks(projectData);
    }
  }

  Logger.log('✓ 全タスク一覧を更新しました');
}

/**
 * メニュー: 期日切れ一覧更新
 */
function menuUpdateOverdue() {
  const sheetManager = new SheetManager();
  sheetManager.updateOverdueTasks();
  Logger.log('✓ 期日切れ一覧を更新しました');
}

/**
 * メニュー: バッチ同期 (次の5シート)
 */
function menuSyncNextBatch() {
  const ui = SpreadsheetApp.getUi();
  try {
    syncNextBatch();
    ui.alert('✅ 同期完了', 'バッチ同期が完了しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ エラー', `同期中にエラーが発生しました:\n${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * メニュー: 期日切れ色付け
 */
function menuUpdateOverdueColors() {
  const ui = SpreadsheetApp.getUi();
  try {
    updateOverdueColors();
    ui.alert('✅ 色付け完了', '期日切れタスクの色付けが完了しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ エラー', `色付け中にエラーが発生しました:\n${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * メニュー: Todoist同期
 */
function menuSyncTodoist() {
  const ui = SpreadsheetApp.getUi();
  try {
    syncTodoistTasks();
    ui.alert('✅ 同期完了', 'Todoistタスクの同期が完了しました。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ エラー', `同期中にエラーが発生しました:\n${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * プロジェクトデータからGanttチャートを生成
 *
 * @param {Object} projectData - プロジェクトデータ（JSON）
 */
function generateGanttFromData(projectData) {
  Logger.log(`プロジェクト「${projectData.project_name}」のGanttチャート生成を開始...`);

  const sheetManager = new SheetManager();

  // 1. プロジェクトGanttシートを作成＆レンダリング
  Logger.log('Ganttチャートを作成中...');
  const ganttSheet = sheetManager.getOrCreateGanttSheet(
    String(projectData.project_id),
    projectData.project_name
  );

  const renderer = new GanttRenderer(ganttSheet, projectData);
  renderer.render();

  // 2. 全タスクリストを更新
  Logger.log('全タスクリストを更新中...');
  sheetManager.updateAllTasks(projectData);

  // 3. プロジェクト一覧を更新（最後に更新してシートリンクを取得）
  Logger.log('プロジェクト一覧を更新中...');
  sheetManager.updateProjectList(projectData);

  Logger.log('✓ すべてのシートを更新完了');
}

/**
 * ファイル名からproject_idを抽出
 * @param {string} fileName - ファイル名（例: "new_192_ポートフォリオ.json" or "update_192_ポートフォリオ.json" or "processed_192_ポートフォリオ.json" or "updated_192_ポートフォリオ.json"）
 * @return {string|null} - プロジェクトID（例: "192"）
 */
function extractProjectId(fileName) {
  // new_, update_, processed_, updated_ を除去
  const nameWithoutPrefix = fileName.replace(/^(new_|update_|processed_|updated_)/, '');

  // 最初の "_" より前を取得
  const match = nameWithoutPrefix.match(/^(\d+)_/);
  return match ? match[1] : null;
}

/**
 * プロジェクト一覧シートに既存プロジェクトがあるか判定
 * @param {string} projectId - プロジェクトID
 * @return {boolean} - 既存プロジェクトならtrue
 */
function existsInProjectList(projectId) {
  const sheetManager = new SheetManager();
  const projectListSheet = sheetManager.ss.getSheetByName(CONFIG.SHEET_NAMES.PROJECT_LIST);

  if (!projectListSheet) {
    return false;
  }

  const data = projectListSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(projectId)) {
      return true;
    }
  }

  return false;
}

/**
 * Google Driveフォルダを監視して未処理のJSONファイルを処理
 *
 * この関数は時間ベーストリガー（1分ごと）で自動実行されます。
 *
 * 処理フロー:
 * 1. CONFIG.DRIVE_FOLDER_IDで指定されたフォルダ内のJSONファイルを検索
 * 2. 未処理ファイル（processed_/updated_で始まらない）を処理
 * 3. 既存プロジェクトの場合は processed_/updated_ 付きでも更新
 * 4. 処理後、ファイル名を変更:
 *    - 新規作成: "processed_XXXX.json"
 *    - 既存更新: "updated_XXXX.json"
 *
 * トリガー設定方法:
 * 1. Apps Scriptエディタで「トリガー」アイコンをクリック
 * 2. 「トリガーを追加」をクリック
 * 3. 実行する関数: checkForNewJsonFiles
 * 4. イベントのソース: 時間主導型
 * 5. 時間ベースのトリガーのタイプ: 分ベースのタイマー
 * 6. 時間の間隔: 1分おき
 */
function checkForNewJsonFiles() {
  try {
    const folderId = CONFIG.DRIVE_FOLDER_ID;
    Logger.log(`[DEBUG] フォルダID: ${folderId}`);

    if (!folderId) {
      Logger.log('⚠ CONFIG.DRIVE_FOLDER_ID が設定されていません。');
      return;
    }

    if (folderId === 'your_folder_id_here') {
      Logger.log('⚠ CONFIG.DRIVE_FOLDER_ID がデフォルト値のままです。実際のフォルダIDを設定してください。');
      return;
    }

    const folder = DriveApp.getFolderById(folderId);
    Logger.log(`[DEBUG] フォルダ名: ${folder.getName()}`);

    const files = folder.getFiles();  // 全ファイルを取得

    let fileCount = 0;
    let processedCount = 0;

    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const mimeType = file.getMimeType();
      fileCount++;

      Logger.log(`[DEBUG] ファイル${fileCount}: ${fileName}`);
      Logger.log(`[DEBUG]   - MIMEタイプ: ${mimeType}`);
      Logger.log(`[DEBUG]   - .jsonで終わる? ${fileName.endsWith('.json')}`);
      Logger.log(`[DEBUG]   - processed/updated済み? ${fileName.startsWith('processed_') || fileName.startsWith('updated_')}`);

      // JSONファイルかチェック
      if (!fileName.endsWith('.json')) {
        Logger.log(`⏭  スキップ（JSONではない）: ${fileName}`);
        continue;
      }

      // project_idを抽出
      const projectId = extractProjectId(fileName);
      Logger.log(`[DEBUG]   - プロジェクトID: ${projectId}`);

      // 処理条件の判定
      const isNewFile = fileName.startsWith('new_');
      const isUpdateFile = fileName.startsWith('update_');
      const isProcessedFile = fileName.startsWith('processed_') || fileName.startsWith('updated_');
      const isExistingProject = projectId && existsInProjectList(projectId);

      // 処理すべきかチェック
      const shouldProcess = (isNewFile || isUpdateFile) && !isProcessedFile;

      if (shouldProcess) {
        // new_ で始まる場合は新規作成
        if (isNewFile) {
          Logger.log(`✓ 新規プロジェクト作成を検出: ${fileName}`);
        }
        // update_ で始まる場合は既存プロジェクト更新
        else if (isUpdateFile) {
          if (!isExistingProject) {
            Logger.log(`⚠ 警告: ${fileName} は update_ プレフィックスですが、プロジェクトID ${projectId} が存在しません。新規作成として処理します。`);
          } else {
            Logger.log(`✓ 既存プロジェクトの更新を検出: ${fileName}`);
          }
        }

        try {
          // JSONファイルを処理
          processJsonFile(file.getId());

          // 処理成功後、ファイル名を変更
          const cleanFileName = fileName.replace(/^(new_|update_)/, '');
          let newFileName;

          if (isUpdateFile && isExistingProject) {
            // update_ で始まり、かつ既存プロジェクトの場合は updated_ を付ける
            newFileName = 'updated_' + cleanFileName;
            Logger.log(`✓ 更新完了。ファイル名を変更: ${newFileName}`);
          } else {
            // new_ で始まる、または update_ でも既存プロジェクトが見つからない場合は processed_ を付ける
            newFileName = 'processed_' + cleanFileName;
            Logger.log(`✓ 新規作成完了。ファイル名を変更: ${newFileName}`);
          }

          file.setName(newFileName);
          processedCount++;
        } catch (error) {
          Logger.log(`✗ ファイル処理エラー (${fileName}): ${error.message}`);
          Logger.log(`✗ スタックトレース: ${error.stack}`);
          // エラーが発生してもループは続ける
        }
      } else {
        Logger.log(`⏭  スキップ: ${fileName}`);
      }
    }

    Logger.log(`[DEBUG] 検出ファイル数: ${fileCount}, 処理数: ${processedCount}`);

    if (processedCount > 0) {
      Logger.log(`✓ ${processedCount}個のファイルを処理しました。`);
    } else {
      Logger.log('未処理のJSONファイルはありません。');
    }

  } catch (error) {
    Logger.log('✗ フォルダ監視エラー: ' + error.message);
    Logger.log('✗ スタックトレース: ' + error.stack);
    sendDiscordNotification(
      `Google Driveフォルダの監視中にエラーが発生しました。\n\nエラー内容: ${error.message}`,
      true
    );
  }
}

/**
 * 手動でJSONファイルを指定して処理
 *
 * @param {string} fileId - Google DriveのファイルID
 */
function processJsonFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    const content = file.getBlob().getDataAsString();
    const projectData = JSON.parse(content);

    Logger.log(`JSONファイル「${file.getName()}」を処理中...`);
    generateGanttFromData(projectData);

    sendDiscordNotification(
      `プロジェクト「${projectData.project_name}」のGanttチャートを生成しました。\n` +
      `ファイル: ${file.getName()}\n` +
      `スプレッドシート: https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}`
    );

  } catch (error) {
    Logger.log('✗ エラー: ' + error);
    sendDiscordNotification(
      `JSONファイル処理中にエラーが発生しました。\n\nエラー内容: ${error.message}`,
      true
    );
    throw error;
  }
}

/**
 * 実際のJSONファイルから生成（Google Driveのファイル使用）
 *
 * 【自動実行】
 * - npm run gantt:save でアップロード時に自動的にprocessJsonFile()が実行されます
 * - Apps Script APIを使用してNode.jsから直接呼び出されます
 *
 * 【手動実行】
 * 1. Google DriveにJSONファイルをアップロード
 * 2. ファイルを右クリック → 「リンクを取得」 → ファイルIDをコピー
 * 3. 下記の FILE_ID を実際のファイルIDに書き換え
 * 4. この関数を実行
 *
 * ファイルIDの例: "1a2b3c4d5e6f7g8h9i0j" （長い英数字の文字列）
 */
function generateGanttFromFile() {
  const FILE_ID = 'your_file_id_here'; // ここに実際のファイルIDを入力

  if (FILE_ID === 'your_file_id_here') {
    Logger.log('⚠ FILE_ID を実際のファイルIDに書き換えてください。');
    return;
  }

  processJsonFile(FILE_ID);
}

/**
 * 全Ganttシートを一括再同期
 *
 * - 全Ganttシートからプロジェクトデータを読み取り
 * - プロジェクト一覧、全タスクリスト、期日切れタスクを更新
 */
function syncNextBatch() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const BATCH_SIZE = 5;

    // 現在のインデックスを取得（初回は0）
    let currentIndex = parseInt(scriptProperties.getProperty('SYNC_CURRENT_INDEX') || '0');

    Logger.log(`=== 自動継続バッチ同期開始 (開始インデックス: ${currentIndex}) ===`);

    const sheetManager = new SheetManager();
    const ss = sheetManager.ss;
    const allSheets = ss.getSheets();

    // Ganttシートを検出
    const ganttSheets = allSheets.filter(sheet => {
      const name = sheet.getName();
      return name.match(/^\d{1,4}_.+$/);
    });

    const totalSheets = ganttSheets.length;
    Logger.log(`✓ 総Ganttシート数: ${totalSheets}件`);

    if (totalSheets === 0) {
      Logger.log('⚠ Ganttシートが見つかりませんでした。');
      scriptProperties.setProperty('SYNC_CURRENT_INDEX', '0');
      return;
    }

    // インデックスが範囲外の場合はリセット
    if (currentIndex >= totalSheets) {
      Logger.log('✓ 全シート処理完了。最初に戻ります。');
      currentIndex = 0;
    }

    // バッチ範囲を計算
    const endIndex = Math.min(currentIndex + BATCH_SIZE, totalSheets);
    const batchSheets = ganttSheets.slice(currentIndex, endIndex);
    const isLastBatch = endIndex >= totalSheets;

    Logger.log(`✓ バッチ処理: ${batchSheets.length}件 (${currentIndex}~${endIndex - 1}) ${isLastBatch ? '[最終バッチ]' : ''}`);

    // 各シートを同期
    let successCount = 0;
    let errorCount = 0;

    for (const ganttSheet of batchSheets) {
      try {
        Logger.log(`--- ${ganttSheet.getName()} を同期中 ---`);

        const projectData = sheetManager.readProjectDataFromSheet(ganttSheet);

        if (!projectData) {
          Logger.log(`⚠ データ読み取り失敗: ${ganttSheet.getName()}`);
          errorCount++;
          continue;
        }

        sheetManager.updateProjectList(projectData);
        sheetManager.updateAllTasks(projectData);
        sheetManager.updateGanttProgressFromStatus(ganttSheet);

        Logger.log(`✓ 同期完了: ${ganttSheet.getName()}`);
        successCount++;

        Utilities.sleep(1000);

      } catch (error) {
        Logger.log(`✗ 同期エラー (${ganttSheet.getName()}): ${error.message}`);
        errorCount++;
      }
    }

    // 最終バッチの場合は期日切れタスク更新
    if (isLastBatch) {
      Logger.log('--- 期日切れタスクシートを更新中 ---');
      try {
        const overdueCount = sheetManager.updateOverdueTasks();
        Logger.log(`✓ 期日切れタスク更新完了: ${overdueCount}件`);
      } catch (error) {
        Logger.log(`✗ 期日切れタスク更新エラー: ${error.message}`);
      }
    }

    // 次回のインデックスを保存
    const nextIndex = isLastBatch ? 0 : endIndex;
    scriptProperties.setProperty('SYNC_CURRENT_INDEX', nextIndex.toString());

    Logger.log('=== バッチ同期完了 ===');
    Logger.log(`成功: ${successCount}件 / エラー: ${errorCount}件`);
    Logger.log(`次回開始インデックス: ${nextIndex} ${isLastBatch ? '(最初に戻る)' : ''}`);

  } catch (error) {
    Logger.log('✗ バッチ全体エラー: ' + error.message);
    Logger.log('✗ スタックトレース: ' + error.stack);
    throw error;
  }
}

/**
 * 全Ganttシートを強制的に再同期
 * （変更検知機能を削除したため、syncAllSheets()と同じ）
 */
function updateOverdueColors() {
  try {
    Logger.log('=== 期日切れタスクの色付け開始 ===');

    const sheetManager = new SheetManager();
    const ss = sheetManager.ss;
    const allSheets = ss.getSheets();

    // Ganttシートを検出
    const ganttSheets = allSheets.filter(sheet => {
      const name = sheet.getName();
      return name.match(/^\d{1,4}_.+$/);
    });

    Logger.log(`✓ Ganttシート検出: ${ganttSheets.length}件`);

    if (ganttSheets.length === 0) {
      Logger.log('⚠ Ganttシートが見つかりませんでした。');
      return;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let totalColoredCells = 0;

    for (const ganttSheet of ganttSheets) {
      try {
        // タスクデータの開始行（GanttRendererの構造: 2行目がヘッダー、4行目からタスク）
        const lastRow = ganttSheet.getLastRow();
        if (lastRow < 4) continue;

        const taskStartRow = 4; // 固定（4行目からタスクデータ）

        // タスクデータの範囲を取得
        const taskDataRange = ganttSheet.getRange(taskStartRow, 1, lastRow - taskStartRow + 1, 9);
        const taskData = taskDataRange.getValues();

        // 全終了日セル（G列）をリセット（シート上の全タスク行）
        const endDateCol = ganttSheet.getRange(taskStartRow, 7, lastRow - taskStartRow + 1, 1);
        endDateCol.setBackground('#FFFFFF');
        endDateCol.setFontColor('#000000');
        endDateCol.setFontWeight('normal');

        let coloredCount = 0;

        // 各タスクをチェック
        for (let i = 0; i < taskData.length; i++) {
          const row = taskData[i];
          const endDateValue = row[6]; // G列（終了日）
          const progress = row[7]; // H列（進捗率）
          const status = String(row[4] || ''); // E列（ステータス）

          // 空行はスキップ
          if (!row[0]) continue;

          // ステータスが「完了」または「中断」の場合はスキップ
          if (status === '完了' || status === '中断') continue;

          // 期日切れ判定（進捗率は0.0-1.0の小数で保存されている）
          if (endDateValue && progress !== 1) {
            const endDate = new Date(endDateValue);
            endDate.setHours(0, 0, 0, 0);

            if (endDate < today) {
              // 期日切れ → 色付け
              const endDateCell = ganttSheet.getRange(taskStartRow + i, 7);
              endDateCell.setBackground('#FF0000'); // 赤背景
              endDateCell.setFontColor('#FFFF00'); // 黄色文字
              endDateCell.setFontWeight('bold'); // 太字
              coloredCount++;
            }
          }
        }

        if (coloredCount > 0) {
          Logger.log(`✓ ${ganttSheet.getName()}: ${coloredCount}件の期日切れタスクに色付け`);
        }
        totalColoredCells += coloredCount;

      } catch (error) {
        Logger.log(`✗ 色付けエラー (${ganttSheet.getName()}): ${error.message}`);
      }
    }

    Logger.log('=== 色付け完了 ===');
    Logger.log(`総計: ${totalColoredCells}件の期日切れタスクに色付け`);

  } catch (error) {
    Logger.log('✗ 全体エラー: ' + error.message);
    Logger.log('✗ スタックトレース: ' + error.stack);
    throw error;
  }
}

/**
 * 期日切れタスクをDiscordに通知
 *
 * - 期日切れタスクシートのデータを読み取り
 * - Discord WebhookでEmbed形式の通知を送信
 */
function sendOverdueTasksNotification() {
  try {
    const sheetManager = new SheetManager();
    const overdueSheet = sheetManager.getOrCreateOverdueTasksSheet();

    // 期日切れタスクを取得
    const lastRow = overdueSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log('✓ 期日切れタスクはありません。');
      return;
    }

    // データを読み取り（ヘッダー除く）
    const data = overdueSheet.getRange(2, 1, lastRow - 1, 9).getValues();

    // プロジェクトごとにグループ化
    const projectGroups = {};
    for (const row of data) {
      const projectId = String(row[0] || '');
      const projectName = String(row[1] || '');
      const taskId = String(row[2] || '');
      const parentTaskName = String(row[3] || '');
      const childTaskName = String(row[4] || '');
      const dueDate = row[5] || '';
      const assignee = String(row[6] || '');
      const daysOverdue = row[7] || 0;
      const progress = row[8] || 0;

      // タスク名を決定
      const taskName = childTaskName || parentTaskName || taskId;

      // 進捗率を正規化（0.0-1.0 → 0-100）
      let progressPercent = 0;
      if (typeof progress === 'number') {
        progressPercent = progress <= 1 ? Math.round(progress * 100) : Math.round(progress);
      }

      if (!projectGroups[projectId]) {
        projectGroups[projectId] = {
          projectName: projectName,
          tasks: []
        };
      }

      projectGroups[projectId].tasks.push({
        taskName: taskName,
        dueDate: dueDate,
        assignee: assignee,
        daysOverdue: daysOverdue,
        progress: progressPercent
      });
    }

    // Embed形式でDiscord通知を作成
    const embeds = [];

    for (const projectId in projectGroups) {
      const group = projectGroups[projectId];

      // タスクリストをMarkdown形式で整形
      const taskList = group.tasks
        .slice(0, 10) // 最大10件まで
        .map(task => {
          const dueDateStr = task.dueDate instanceof Date
            ? `${task.dueDate.getMonth() + 1}/${task.dueDate.getDate()}`
            : task.dueDate;
          return `• **${task.taskName}** (期日: ${dueDateStr}, 遅延: ${task.daysOverdue}日, 進捗: ${task.progress}%)`;
        })
        .join('\n');

      const moreCount = group.tasks.length > 10 ? `\n... 他${group.tasks.length - 10}件` : '';

      embeds.push({
        title: `⚠️ 期日切れタスク: ${group.projectName}`,
        description: taskList + moreCount,
        color: 0xFF6B6B, // 赤色
        footer: {
          text: `合計 ${group.tasks.length}件の期日切れタスク`
        }
      });
    }

    if (embeds.length === 0) {
      Logger.log('✓ 期日切れタスクはありません。');
      return;
    }

    // Discord Webhook送信
    const payload = { embeds: embeds };
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(CONFIG.DISCORD_WEBHOOK_URL, options);
    Logger.log(`✓ Discord通知送信完了: ${response.getResponseCode()}`);

  } catch (error) {
    Logger.log(`✗ Discord通知エラー: ${error.message}`);
    Logger.log(`✗ スタックトレース: ${error.stack}`);
    throw error;
  }
}

/**
 * TodoistのInboxタスクを同期
 *
 * - Todoist REST API v2を使用してInboxタスクを取得
 * - Todoistタスクシートに書き込み
 */
function syncTodoistTasks() {
  try {
    Logger.log('=== Todoistタスク同期開始 ===');

    // 1. Todoist APIからタスクを取得
    const apiKey = CONFIG.TODOIST_API_KEY;
    if (!apiKey) {
      Logger.log('⚠ TODOIST_API_KEY が設定されていません。');
      return 0;
    }

    const url = 'https://api.todoist.com/rest/v2/tasks';
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${apiKey}`
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log(`✗ Todoist API エラー: ${responseCode}`);
      Logger.log(`✗ レスポンス: ${response.getContentText()}`);
      return 0;
    }

    const allTasks = JSON.parse(response.getContentText());
    Logger.log(`✓ Todoistタスク取得完了: ${allTasks.length}件`);

    // 2. プロジェクト一覧を取得してInboxプロジェクトIDを特定
    const projectsUrl = 'https://api.todoist.com/rest/v2/projects';
    const projectsResponse = UrlFetchApp.fetch(projectsUrl, options);
    const projects = JSON.parse(projectsResponse.getContentText());

    let inboxProjectId = null;
    for (const project of projects) {
      if (project.name === 'Inbox' || project.is_inbox_project) {
        inboxProjectId = String(project.id);
        Logger.log(`✓ Inboxプロジェクト検出: ${inboxProjectId}`);
        break;
      }
    }

    if (!inboxProjectId) {
      Logger.log('⚠ InboxプロジェクトIDが見つかりませんでした。');
      return 0;
    }

    // 3. Inboxプロジェクトのタスクのみをフィルタリング
    const inboxTasks = allTasks.filter(task => String(task.project_id) === inboxProjectId);
    Logger.log(`✓ Inboxタスク抽出完了: ${inboxTasks.length}件`);

    // 4. シートマネージャーを使用してタスクを更新
    const sheetManager = new SheetManager();
    const count = sheetManager.updateTodoistTasks(inboxTasks);

    Logger.log(`✓ Todoistタスク同期完了: ${count}件`);
    return count;

  } catch (error) {
    Logger.log(`✗ Todoistタスク同期エラー: ${error.message}`);
    Logger.log(`✗ スタックトレース: ${error.stack}`);
    throw error;
  }
}

/**
 * Todoist連携機能のテスト（開発用）
 *
 * 使用方法:
 * 1. GASエディタでこの関数を選択
 * 2. 実行ボタンをクリック
 * 3. 以下がテストされます:
 *    - Todoist APIからInboxタスクを取得
 *    - Todoistタスクシートへの書き込み
 *    - Discord通知の送信
 */
/**
 * 既存シートに実績工数列を追加するマイグレーション関数
 *
 * 【重要】この関数は一度だけ実行してください。
 *
 * 処理内容:
 * 1. 全Ganttシート（X_プロジェクト名形式）の12列目（L列）に「実績工数（h）」列を挿入
 * 2. 全プロジェクトタスクシートの14列目（N列）に「実績工数（h）」列を挿入
 *
 * 使用方法:
 * 1. GASエディタでこの関数を選択
 * 2. 実行ボタンをクリック
 * 3. 実行後、各シートに「実績工数（h）」列が追加されたことを確認
 */
function migrateExistingSheetsForActualHours() {
  try {
    Logger.log('=== 実績工数列マイグレーション開始 ===');

    const sheetManager = new SheetManager();
    const ss = sheetManager.ss;
    const allSheets = ss.getSheets();

    let ganttCount = 0;
    let errorCount = 0;

    // 1. 全Ganttシートを処理
    Logger.log('--- Ganttシートを処理中 ---');

    for (const sheet of allSheets) {
      const sheetName = sheet.getName();

      // Ganttシート（X_プロジェクト名 または XXXX_プロジェクト名形式）を検出
      if (sheetName.match(/^\d{1,4}_.+$/)) {
        try {
          Logger.log(`処理中: ${sheetName}`);

          // 2行目（ヘッダー行）の12列目（L列）の値を確認
          const currentHeader = sheet.getRange(2, 12).getValue();

          // すでに「実績工数（h）」がある場合はスキップ
          if (currentHeader === '実績工数（h）') {
            Logger.log(`  ⏭  スキップ（すでに実績工数列あり）: ${sheetName}`);
            continue;
          }

          // 12列目（L列）に新しい列を挿入
          sheet.insertColumnAfter(11); // K列の後ろに挿入

          // ヘッダー行（2行目）の12列目に「実績工数（h）」を設定
          sheet.getRange(2, 12).setValue('実績工数（h）');

          // ヘッダーのスタイルを適用（他のヘッダーと同じスタイル）
          const headerCell = sheet.getRange(2, 12);
          headerCell.setBackground('#E8F4FD');
          headerCell.setHorizontalAlignment('center');
          headerCell.setVerticalAlignment('middle');

          // 1行目の12列目も同じ色で塗る
          sheet.getRange(1, 12).setBackground('#E8F4FD');
          sheet.getRange(1, 12).setHorizontalAlignment('center');
          sheet.getRange(1, 12).setVerticalAlignment('middle');

          Logger.log(`  ✓ 完了: ${sheetName}`);
          ganttCount++;

        } catch (error) {
          Logger.log(`  ✗ エラー (${sheetName}): ${error.message}`);
          errorCount++;
        }
      }
    }

    Logger.log(`✓ Ganttシート処理完了: ${ganttCount}件`);

    // 2. 全プロジェクトタスクシートを処理
    Logger.log('--- 全プロジェクトタスクシートを処理中 ---');

    const allTasksSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ALL_TASKS);

    if (allTasksSheet) {
      try {
        // 1行目（ヘッダー行）の14列目（N列）の値を確認
        const currentHeader = allTasksSheet.getRange(1, 14).getValue();

        // すでに「実績工数（h）」がある場合はスキップ
        if (currentHeader === '実績工数（h）') {
          Logger.log('  ⏭  スキップ（すでに実績工数列あり）: 全プロジェクトタスク');
        } else {
          // 14列目（N列）に新しい列を挿入
          allTasksSheet.insertColumnAfter(13); // M列の後ろに挿入

          // ヘッダー行（1行目）の14列目に「実績工数（h）」を設定
          allTasksSheet.getRange(1, 14).setValue('実績工数（h）');

          // ヘッダーのスタイルを適用
          const headerCell = allTasksSheet.getRange(1, 14);
          headerCell.setBackground('#4A90E2');
          headerCell.setFontColor('#FFFFFF');
          headerCell.setFontWeight('bold');
          headerCell.setHorizontalAlignment('center');

          // 列幅を設定
          allTasksSheet.setColumnWidth(14, 75);

          Logger.log('  ✓ 完了: 全プロジェクトタスク');
        }
      } catch (error) {
        Logger.log(`  ✗ エラー (全プロジェクトタスク): ${error.message}`);
        errorCount++;
      }
    } else {
      Logger.log('  ⚠ 全プロジェクトタスクシートが見つかりません。');
    }

    Logger.log('=== マイグレーション完了 ===');
    Logger.log(`成功: ${ganttCount}件のGanttシート + 全タスクシート / エラー: ${errorCount}件`);

    // Discord通知
    sendDiscordNotification(
      `✅ **実績工数列マイグレーション完了**\n\n` +
      `Ganttシート: ${ganttCount}件\n` +
      `全プロジェクトタスクシート: 処理済み\n` +
      `エラー: ${errorCount}件`,
      false
    );

  } catch (error) {
    Logger.log('✗ マイグレーションエラー: ' + error.message);
    Logger.log('✗ スタックトレース: ' + error.stack);
    sendDiscordNotification(
      `実績工数列マイグレーション中にエラーが発生しました。\n\nエラー内容: ${error.message}`,
      true
    );
    throw error;
  }
}

/**
 * 依存タスクブロッカー検知: 遅延タスクの影響分析
 *
 * この関数は遅延しているタスクを検出し、それらが他のタスクに与える影響を分析します。
 *
 * @param {Object} projectData - プロジェクトデータ
 * @return {Object} 分析結果
 *   {
 *     delayedTasks: [遅延タスクの配列],
 *     impactedTasks: {
 *       delayed_task_id: {
 *         impactedTaskIds: [影響を受けるタスクIDの配列],
 *         criticalPathLength: クリティカルパスの長さ
 *       }
 *     },
 *     circularDependencies: [[循環依存タスクIDの配列], ...],
 *     taskMap: { task_id: task_object }
 *   }
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// ============================================
// 変更検知機能（パフォーマンス最適化）
// ============================================

/**
 * Ganttシートの最終同期日時を取得
 * @param {Sheet} sheet - Ganttシート
 * @returns {Date|null} - 最終同期日時（未同期の場合はnull）
 */
function getLastSyncTime(sheet) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const key = `lastSync_${sheet.getName()}`;
    const value = props.getProperty(key);
    return value ? new Date(value) : null;
  } catch (error) {
    return null;
  }
}

/**
 * Ganttシートの最終同期日時を記録
 * @param {Sheet} sheet - Ganttシート
 */
function setLastSyncTime(sheet) {
  try {
    const props = PropertiesService.getDocumentProperties();
    const key = `lastSync_${sheet.getName()}`;
    const now = new Date().toISOString();
    props.setProperty(key, now);
  } catch (error) {
    Logger.log(`⚠ 最終同期日時の記録失敗: ${sheet.getName()} - ${error.message}`);
  }
}

/**
 * Ganttシートが同期が必要かどうか判定
 * @param {Sheet} sheet - Ganttシート
 * @param {number} minIntervalMinutes - 最小同期間隔（分）デフォルト5分
 * @returns {boolean} - 同期が必要ならtrue
 */
function shouldSyncSheet(sheet, minIntervalMinutes = 5) {
  const lastSyncTime = getLastSyncTime(sheet);

  // 初回同期（最終同期日時が未設定）
  if (!lastSyncTime) {
    return true;
  }

  // 現在時刻との差分を計算（分単位）
  const now = new Date();
  const diffMinutes = (now - lastSyncTime) / (1000 * 60);

  // 最小間隔を経過していれば同期が必要
  return diffMinutes >= minIntervalMinutes;
}

