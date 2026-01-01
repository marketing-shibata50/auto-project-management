/**
 * シート管理クラス
 */

class SheetManager {
  constructor() {
    // スプレッドシートにバインドされたGASの場合はgetActiveSpreadsheet()を使用
    // スタンドアロンGASの場合はopenById()を使用
    try {
      this.ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!this.ss) {
        this.ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      }
    } catch (e) {
      this.ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    }
  }
  
  /**
   * プロジェクト一覧シートの取得または作成
   */
  getOrCreateProjectListSheet() {
    let sheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.PROJECT_LIST);
    
    if (!sheet) {
      sheet = this.ss.insertSheet(CONFIG.SHEET_NAMES.PROJECT_LIST);
      this._initializeProjectListSheet(sheet);
    }
    
    return sheet;
  }
  
  /**
   * 全プロジェクトタスクシートの取得または作成
   */
  getOrCreateAllTasksSheet() {
    let sheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.ALL_TASKS);

    if (!sheet) {
      sheet = this.ss.insertSheet(CONFIG.SHEET_NAMES.ALL_TASKS);
      this._initializeAllTasksSheet(sheet);
    }

    return sheet;
  }

  /**
   * 期日切れタスクシートの取得または作成
   */
  getOrCreateOverdueTasksSheet() {
    let sheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.OVERDUE_TASKS);

    if (!sheet) {
      sheet = this.ss.insertSheet(CONFIG.SHEET_NAMES.OVERDUE_TASKS);
      this._initializeOverdueTasksSheet(sheet);
    }

    return sheet;
  }

  /**
   * Todoistタスクシートの取得または作成
   */
  getOrCreateTodoistTasksSheet() {
    let sheet = this.ss.getSheetByName(CONFIG.SHEET_NAMES.TODOIST_TASKS);

    if (!sheet) {
      sheet = this.ss.insertSheet(CONFIG.SHEET_NAMES.TODOIST_TASKS);
      this._initializeTodoistTasksSheet(sheet);
    }

    return sheet;
  }

  /**
   * プロジェクト別Ganttシートの取得または作成
   */
  getOrCreateGanttSheet(projectId, projectName) {
    const sheetName = `${projectId}_${projectName}`;
    let sheet = this.ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = this.ss.insertSheet(sheetName);
      this._initializeGanttSheet(sheet);
    }

    return sheet;
  }

  /**
   * バーンダウンデータシートの取得または作成
   */
  getOrCreateBurndownSheet() {
    const sheetName = 'BurndownData';
    let sheet = this.ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = this.ss.insertSheet(sheetName);
      this._initializeBurndownSheet(sheet);
      Logger.log('[INFO] Created new burndown sheet: ' + sheetName);
    }

    return sheet;
  }

  /**
   * プロジェクト一覧シートの初期化
   */
  _initializeProjectListSheet(sheet) {
    const headers = [
      'プロジェクトID',
      'プロジェクト名',
      'シートリンク',
      'GitHubリンク',
      'プロジェクト目的',
      'プロジェクトジャンル',
      'プロジェクト期日',
      '総タスク数',
      '完了タスク数',
      '進捗率',
      '総工数見積もり',
      '最終更新日時'
    ];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    sheet.setFrozenRows(1);

    // 列幅を75に設定
    for (let i = 1; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 75);
    }
  }
  
  /**
   * 全プロジェクトタスクシートの初期化
   */
  _initializeAllTasksSheet(sheet) {
    const headers = [
      'プロジェクトID',
      'プロジェクト名',
      'ID',
      '親タスク名',
      '子タスク名',
      'タグ',
      'ステータス',
      '開始日',
      '終了日',
      '進捗率',
      '担当者',
      '優先度',
      '工数見積（h）',
      '実績工数（h）'
    ];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    sheet.setFrozenRows(1);

    // 列幅を75に設定
    for (let i = 1; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 75);
    }
  }

  /**
   * 期日切れタスクシートの初期化
   */
  _initializeOverdueTasksSheet(sheet) {
    const headers = [
      'PID',
      'プロジェクト名',
      'TID',
      '親タスク',
      '子タスク',
      '期日',
      '担当',
      '遅延日数',
      '進捗率'
    ];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#FF6B6B');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    sheet.setFrozenRows(1);

    // 列幅を個別に設定
    const columnWidths = [
      60,   // PID
      200,  // プロジェクト名
      60,   // TID
      300,  // 親タスク
      300,  // 子タスク
      80,   // 期日
      80,   // 担当
      80,   // 遅延日数
      80    // 進捗率
    ];

    for (let i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i]);
    }
  }

  /**
   * Todoistタスクシートの初期化
   */
  _initializeTodoistTasksSheet(sheet) {
    const headers = [
      'タスクID',
      'タスク名',
      '説明',
      '期日',
      '優先度',
      'ラベル',
      '完了状態',
      'Todoistリンク',
      '最終更新日時'
    ];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#E37400');
    headerRange.setFontColor('#FFFFFF');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');

    sheet.setFrozenRows(1);

    // 列幅を設定
    sheet.setColumnWidth(1, 80);  // タスクID
    sheet.setColumnWidth(2, 200); // タスク名
    sheet.setColumnWidth(3, 250); // 説明
    sheet.setColumnWidth(4, 100); // 期日
    sheet.setColumnWidth(5, 60);  // 優先度
    sheet.setColumnWidth(6, 150); // ラベル
    sheet.setColumnWidth(7, 80);  // 完了状態
    sheet.setColumnWidth(8, 150); // Todoistリンク
    sheet.setColumnWidth(9, 150); // 最終更新日時
  }


  /**
   * Ganttシートの初期化
   *
   * 注: GanttRendererが全体のレイアウトを管理するため、
   * ここでは最小限のシート設定のみを行う
   */
  _initializeGanttSheet(sheet) {
    // GanttRendererが以下を設定:
    // - 1行目: 月（タイムライン、2色交互）
    // - 2行目: 列名（A~K） + 日（タイムライン、2色交互）
    // - 3行目: 空白行
    // - 4行目以降: タスクデータ（親タスクは1行全体が太字）
    // - 固定行: 2行（月1 + 列名/日1）
    // - 固定列: 3列（ID + 親タスク名 + 子タスク名）

    // シートの基本設定のみ実行
    sheet.setFrozenColumns(3); // ID、親タスク名、子タスク名を固定
  }

  /**
   * バーンダウンシートの初期化
   */
  _initializeBurndownSheet(sheet) {
    // ヘッダー行（1行目）
    const headers = [
      '記録日',           // A: Date
      'プロジェクト名',   // B: Project Name
      '予定進捗率',       // C: Expected Progress (%)
      '実績進捗率',       // D: Actual Progress (%)
      '残タスク数',       // E: Remaining Tasks
      '完了タスク数',     // F: Completed Tasks
      '総タスク数',       // G: Total Tasks
      'ベロシティ',       // H: Velocity (tasks/day)
      '完了予測日'        // I: Predicted Completion Date
    ];

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4A90E2');
    headerRange.setFontColor('#FFFFFF');

    // 列幅設定
    sheet.setColumnWidth(1, 120);  // 記録日
    sheet.setColumnWidth(2, 150);  // プロジェクト名
    sheet.setColumnWidth(3, 100);  // 予定進捗率
    sheet.setColumnWidth(4, 100);  // 実績進捗率
    sheet.setColumnWidth(5, 100);  // 残タスク数
    sheet.setColumnWidth(6, 100);  // 完了タスク数
    sheet.setColumnWidth(7, 100);  // 総タスク数
    sheet.setColumnWidth(8, 120);  // ベロシティ
    sheet.setColumnWidth(9, 150);  // 完了予測日

    // 条件付き書式: 実績が予定より遅れている場合は赤色
    const dataRange = sheet.getRange('D2:D1000');
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=D2<C2')
      .setBackground('#FFCDD2')  // Light red
      .setRanges([dataRange])
      .build();

    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);

    // 先頭行を固定
    sheet.setFrozenRows(1);

    Logger.log('[INFO] Initialized burndown sheet with headers and formatting');
  }

  /**
   * プロジェクト一覧を更新
   */
  updateProjectList(projectData) {
    const sheet = this.getOrCreateProjectListSheet();
    const projectId = String(projectData.project_id); // 文字列に統一
    const projectName = projectData.project_name;

    // 既存のプロジェクトを検索
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    // 検索用に数値化（"0009" → 9, "9" → 9）
    const projectIdNum = parseInt(projectId, 10);

    for (let i = 1; i < data.length; i++) {
      const existingId = data[i][0];
      const existingIdNum = parseInt(String(existingId), 10);

      if (existingIdNum === projectIdNum) { // 数値として比較
        rowIndex = i + 1;
        break;
      }
    }

    // タスク数と進捗を計算
    const totalTasks = projectData.tasks.length;
    const completedTasks = projectData.tasks.filter(t => t.progress === 100).length;

    // 全タスクの進捗率の平均を計算
    const totalProgress = projectData.tasks.reduce((sum, t) => sum + (t.progress || 0), 0);
    const progressRate = totalTasks > 0 ? Math.round(totalProgress / totalTasks) : 0;

    const totalHours = projectData.tasks.reduce((sum, t) => sum + (parseFloat(t.estimated_hours) || 0), 0);

    // Ganttシートのリンクを生成（SheetIDマップを使用して確実に取得）
    const projectIdPadded = String(projectId).padStart(4, '0');
    let sheetId = '';

    // 全シートを検索してプロジェクトIDが一致するGanttシートを探す
    const allSheets = this.ss.getSheets();
    for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      if (sheetName.match(/^\d{1,4}_.+$/)) {
        const match = sheetName.match(/^(\d{1,4})_/);
        if (match) {
          const sheetProjectId = match[1].padStart(4, '0');
          if (sheetProjectId === projectIdPadded) {
            sheetId = sheet.getSheetId();
            break;
          }
        }
      }
    }

    const sheetLink = sheetId ? `=HYPERLINK("#gid=${sheetId}", "シート")` : '';

    // GitHubリンクを生成（JSONにgithub_urlがあればそれを使用）
    const githubUrl = projectData.github_url || '';
    const githubLink = githubUrl ? `=HYPERLINK("${githubUrl}", "リンク")` : '';

    const rowData = [
      projectId,
      projectName,
      sheetLink,
      githubLink,
      projectData.project_purpose || '',
      projectData.project_type || projectData.project_genre || '',
      projectData.project_deadline || '',
      totalTasks,
      completedTasks,
      progressRate + '%',
      totalHours,
      new Date().toLocaleString('ja-JP')
    ];

    if (rowIndex === -1) {
      // 新規追加
      sheet.appendRow(rowData);
    } else {
      // 既存更新
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    }
  }
  
  /**
   * 全タスクリストを更新
   */
  updateAllTasks(projectData) {
    const sheet = this.getOrCreateAllTasksSheet();
    const projectId = String(projectData.project_id); // 文字列に統一
    const projectName = projectData.project_name;

    // 既存のプロジェクトのタスクを削除
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];

    // 検索用に数値化（"0009" → 9, "9" → 9）
    const projectIdNum = parseInt(projectId, 10);

    for (let i = data.length - 1; i >= 1; i--) {
      const existingIdNum = parseInt(String(data[i][0]), 10);
      if (existingIdNum === projectIdNum) { // 数値として比較
        rowsToDelete.push(i + 1);
      }
    }

    if (rowsToDelete.length > 0) {
      rowsToDelete.forEach(row => {
        sheet.deleteRow(row);
      });
    }

    // 新しいタスクを追加（親タスク名・子タスク名を分離）
    const taskRows = projectData.tasks.map(task => {
      // 親タスク名と子タスク名を設定
      let parentTaskName = '';
      let childTaskName = '';

      if (task.parent_task_id) {
        // 子タスクの場合：B列に親タスク名、C列に子タスク名
        const parentTask = projectData.tasks.find(t => t.task_id === task.parent_task_id);
        parentTaskName = parentTask ? parentTask.task_name : '';
        childTaskName = task.task_name;
      } else {
        // 親タスクの場合：B列に親タスク名、C列は空白
        parentTaskName = task.task_name;
        childTaskName = '';
      }

      // タグは最初の1つだけ（配列の場合）
      let tag = '';
      if (Array.isArray(task.tags) && task.tags.length > 0) {
        tag = task.tags[0];
      } else if (task.tags) {
        tag = task.tags;
      }

      // ステータスはタスクデータから取得（デフォルトは進捗率から判定）
      const progress = task.progress || 0;
      let status = task.status || '';

      // ステータスが空の場合のみ進捗率から判定
      if (!status) {
        if (progress === 0) {
          status = '未着手';
        } else if (progress === 100) {
          status = '完了';
        } else {
          status = '進行中';
        }
      }

      return [
        projectId,
        projectName,
        task.task_id,
        parentTaskName,
        childTaskName,
        tag,
        status,
        task.start_date || '',
        task.end_date || '',
        progress / 100, // 進捗率を0.5 = 50%形式で
        task.assignee || '',
        task.priority || '',
        task.estimated_hours || '',
        task.actual_hours || ''
      ];
    });

    if (taskRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, taskRows.length, taskRows[0].length).setValues(taskRows);

      // 書式と色を適用
      this._applyAllTasksFormatting(sheet, startRow, taskRows.length);
    }
  }

  /**
   * 全タスクリストに書式と色を適用
   */
  _applyAllTasksFormatting(sheet, startRow, numRows) {
    // タグ列（F列 = 6列目）に色付き条件付き書式を設定
    const tagCol = sheet.getRange(startRow, 6, numRows, 1);
    const tagColors = CONFIG.TAG_COLORS;

    // まずタグ列全体にデフォルト色を適用
    tagCol.setBackground(tagColors['デフォルト']);
    tagCol.setFontColor('#FFFFFF');

    let rules = sheet.getConditionalFormatRules();

    Object.keys(tagColors).forEach(tag => {
      // 'デフォルト'はスキップ（特定のタグ名ではないため）
      if (tag === 'デフォルト') return;

      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(tag)
        .setBackground(tagColors[tag])
        .setFontColor('#FFFFFF')
        .setRanges([tagCol])
        .build();
      rules.push(rule);
    });

    // ステータス列（G列 = 7列目）に色付き条件付き書式を設定
    const statusCol = sheet.getRange(startRow, 7, numRows, 1);

    const ruleNotStarted = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('未着手')
      .setBackground('#E0E0E0')
      .setRanges([statusCol])
      .build();
    rules.push(ruleNotStarted);

    const ruleInProgress = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('進行中')
      .setBackground('#FFF59D')
      .setRanges([statusCol])
      .build();
    rules.push(ruleInProgress);

    const ruleCompleted = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('完了')
      .setBackground('#A5D6A7')
      .setRanges([statusCol])
      .build();
    rules.push(ruleCompleted);

    const ruleInterrupted = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('中断')
      .setBackground('#FF6B6B')
      .setRanges([statusCol])
      .build();
    rules.push(ruleInterrupted);

    // 進捗率列（J列 = 10列目）にデータバー表示
    const progressCol = sheet.getRange(startRow, 10, numRows, 1);
    progressCol.setNumberFormat('0%');

    const ruleProgress = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#4A90E2', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([progressCol])
      .build();
    rules.push(ruleProgress);

    // 日付列（H列・I列 = 8,9列目）のフォーマット
    const dateColStart = sheet.getRange(startRow, 8, numRows, 1);
    const dateColEnd = sheet.getRange(startRow, 9, numRows, 1);
    dateColStart.setNumberFormat('yy/mm/dd');
    dateColEnd.setNumberFormat('yy/mm/dd');

    sheet.setConditionalFormatRules(rules);
  }

  /**
   * Ganttシートのステータスから進捗率を更新
   */
  updateGanttProgressFromStatus(sheet) {
    try {
      const lastRow = sheet.getLastRow();
      if (lastRow < 4) return;

      // タスクデータ範囲（4行目以降、E列=ステータス、H列=進捗率）
      const statusRange = sheet.getRange(4, 5, lastRow - 3, 1); // E列
      const progressRange = sheet.getRange(4, 8, lastRow - 3, 1); // H列

      const statuses = statusRange.getValues();
      const progresses = progressRange.getValues();

      let updatedCount = 0;

      for (let i = 0; i < statuses.length; i++) {
        const status = String(statuses[i][0] || '');
        let currentProgress = progresses[i][0] || 0;

        // 数値に変換（0.5 → 0.5, 50% → 0.5, 50 → 50）
        if (typeof currentProgress === 'string') {
          currentProgress = parseFloat(currentProgress.replace('%', '')) / 100;
        } else if (currentProgress > 1) {
          currentProgress = currentProgress / 100;
        }

        let newProgress = currentProgress;

        // ステータスに応じて進捗率を更新
        if (status === '未着手') {
          newProgress = 0;
        } else if (status === '完了') {
          newProgress = 1; // 100%
        } else if (status === '進行中' && currentProgress === 0) {
          newProgress = 0.5; // 進行中で0%の場合は50%に設定
        }
        // 中断や他のステータスは現在の値を保持

        // 変更があった場合のみ更新
        if (newProgress !== currentProgress) {
          progresses[i][0] = newProgress;
          updatedCount++;
        }
      }

      if (updatedCount > 0) {
        progressRange.setValues(progresses);
      }

    } catch (error) {
      Logger.log(`✗ 進捗率更新エラー: ${error.message}`);
    }
  }

  /**
   * Ganttシートからプロジェクトデータを読み取る
   */
  readProjectDataFromSheet(sheet) {
    try {
      const sheetName = sheet.getName();

      // シート名から project_id と project_name を抽出（X_プロジェクト名 または XXXX_プロジェクト名形式）
      const match = sheetName.match(/^(\d{1,4})_(.+)$/);
      if (!match) {
        Logger.log(`⚠ シート名が正しい形式ではありません: ${sheetName}`);
        return null;
      }

      // IDを4桁にゼロ埋め（例: "9" → "0009"）
      const projectId = match[1].padStart(4, '0');
      const projectName = match[2];

      // タスクデータを読み取り（4行目以降、A~J列）
      const lastRow = sheet.getLastRow();
      if (lastRow < 4) {
        Logger.log(`⚠ タスクデータが存在しません: ${sheetName}`);
        return null;
      }

      const taskDataRange = sheet.getRange(4, 1, lastRow - 3, 12); // A~L列（12列）
      const taskData = taskDataRange.getValues();

      // タスクオブジェクトを構築
      const tasks = [];
      const taskMap = {}; // task_id -> task の対応表（親タスク検索用）

      for (let i = 0; i < taskData.length; i++) {
        const row = taskData[i];

        // IDが空の行はスキップ
        if (!row[0]) continue;

        // 親タスク名（B列 = row[1]）と子タスク名（C列 = row[2]）
        const parentTaskName = row[1] || '';
        const childTaskName = row[2] || '';

        // タスク名を決定（子タスクがある場合は子タスク名、なければ親タスク名）
        const taskName = childTaskName || parentTaskName;

        // 親タスクIDの判定（子タスク名があり親タスク名が空の場合は子タスク）
        let parentTaskId = null;
        if (childTaskName && !parentTaskName) {
          // 子タスクの場合：直前の親タスク（B列に値があり、C列が空白）を探す
          for (let j = i - 1; j >= 0; j--) {
            const prevRow = taskData[j];
            const prevParentTaskName = prevRow[1] || '';
            const prevChildTaskName = prevRow[2] || '';

            // 直前の親タスク（B列に値があり、C列が空白）
            if (prevParentTaskName && !prevChildTaskName) {
              parentTaskId = String(prevRow[0] || '');
              break;
            }
          }
        }

        // タグを配列に変換（D列 = row[3]）
        const tag = row[3] || '';
        const tags = tag ? [tag] : [];

        // 進捗率（H列 = row[7]）
        // スプレッドシートでは0.5または50%のように表示される
        let progressValue = row[7] || 0;
        if (typeof progressValue === 'number') {
          // 0.5のような小数の場合は100倍
          if (progressValue <= 1) {
            progressValue = progressValue * 100;
          }
        } else if (typeof progressValue === 'string') {
          // "50%"のような文字列の場合は数値に変換
          progressValue = parseFloat(progressValue.replace('%', '')) || 0;
        }
        const progress = Math.round(progressValue);

        // 日付をフォーマット（YYYY-MM-DD形式に変換）
        const formatDate = (dateValue) => {
          if (!dateValue) return '';
          if (dateValue instanceof Date) {
            const year = dateValue.getFullYear();
            const month = String(dateValue.getMonth() + 1).padStart(2, '0');
            const day = String(dateValue.getDate()).padStart(2, '0');
            return `${year}-${month}-${day}`;
          }
          return String(dateValue);
        };

        tasks.push({
          task_id: String(row[0] || ''), // A列
          task_name: taskName, // B列または C列
          start_date: formatDate(row[5]), // F列
          end_date: formatDate(row[6]), // G列
          assignee: String(row[8] || ''), // I列
          priority: String(row[9] || ''), // J列
          tags: tags, // D列
          estimated_hours: parseFloat(row[10]) || 0, // K列
          actual_hours: parseFloat(row[11]) || 0, // L列
          progress: progress, // H列
          status: String(row[4] || '未着手'), // E列 - ステータス
          parent_task_id: parentTaskId,
          dependencies: [], // シートには保存されていない
          is_milestone: false, // シートには保存されていない
          description: '' // シートには保存されていない
        });
      }

      // プロジェクトデータを構築
      const projectData = {
        project_id: projectId,
        project_name: projectName,
        project_purpose: '',
        project_type: '',
        project_deadline: '',
        github_url: '',
        tasks: tasks
      };

      // プロジェクト一覧から追加情報を取得
      this._enrichProjectDataFromList(projectData);

      return projectData;

    } catch (error) {
      Logger.log(`✗ シートからの読み取りエラー: ${error.message}`);
      return null;
    }
  }

  /**
   * プロジェクト一覧からプロジェクトデータを補完
   */
  _enrichProjectDataFromList(projectData) {
    try {
      const listSheet = this.getOrCreateProjectListSheet();
      const data = listSheet.getDataRange().getValues();
      const projectId = String(projectData.project_id); // 文字列に統一

      // 検索用に数値化（"0009" → 9, "9" → 9）
      const projectIdNum = parseInt(projectId, 10);

      // プロジェクトIDで検索
      for (let i = 1; i < data.length; i++) {
        const existingIdNum = parseInt(String(data[i][0]), 10);
        if (existingIdNum === projectIdNum) { // 数値として比較
          // プロジェクト一覧から情報を補完
          projectData.project_purpose = data[i][4] || '';
          projectData.project_type = data[i][5] || '';
          projectData.project_deadline = data[i][6] || '';
          // GitHubリンクはHYPERLINK関数から抽出（複雑なため省略）
          break;
        }
      }
    } catch (error) {
      Logger.log(`⚠ プロジェクト一覧からの情報補完に失敗: ${error.message}`);
    }
  }

  /**
   * 期日切れタスクを更新
   */
  updateOverdueTasks() {
    try {
      const sheet = this.getOrCreateOverdueTasksSheet();

      const today = new Date();
      today.setHours(0, 0, 0, 0);

      // 既存データをクリア（ヘッダー以外）
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // 行削除ではなく範囲クリアを使用（固定行すべて削除エラー回避）
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      }

      // ヘッダーを再設定（フォーマット更新のため）
      this._initializeOverdueTasksSheet(sheet);

      // 全プロジェクトタスクシートからデータを取得
      const allTasksSheet = this.getOrCreateAllTasksSheet();
      const allTasksLastRow = allTasksSheet.getLastRow();

      if (allTasksLastRow < 2) {
        Logger.log('全プロジェクトタスクシートにデータがありません。');
        return 0;
      }

      // 全プロジェクトタスクシートから全データを取得（2行目以降、13列）
      // 列: プロジェクトID(0), プロジェクト名(1), ID(2), 親タスク名(3), 子タスク名(4), タグ(5), ステータス(6), 開始日(7), 終了日(8), 進捗率(9), 担当者(10), 優先度(11), 工数見積(12)
      const allTasksData = allTasksSheet.getRange(2, 1, allTasksLastRow - 1, 13).getValues();

      // プロジェクトIDからシートIDへのマッピングを作成
      const sheetIdMap = {};
      const allSheets = this.ss.getSheets();
      for (const ganttSheet of allSheets) {
        const sheetName = ganttSheet.getName();
        if (sheetName.match(/^\d{1,4}_.+$/)) {
          const match = sheetName.match(/^(\d{1,4})_/);
          if (match) {
            const projectId = match[1].padStart(4, '0');
            sheetIdMap[projectId] = ganttSheet.getSheetId();
          }
        }
      }

      // 期日切れタスクをフィルタリング
      const overdueTasksData = [];

      for (const row of allTasksData) {
        const projectId = String(row[0] || '');
        const projectName = String(row[1] || '');
        const taskId = String(row[2] || '');
        const parentTaskName = String(row[3] || '');
        const childTaskName = String(row[4] || '');
        const status = String(row[6] || '未着手');
        const endDate = row[8]; // 終了日（期日）
        let progressValue = row[9] || 0; // 進捗率
        const assignee = String(row[10] || '');

        // タスクIDが空の場合、または「T」で始まらない場合はスキップ
        if (!taskId || !taskId.startsWith('T')) continue;

        // 期日が設定されていない場合はスキップ
        if (!endDate || !(endDate instanceof Date)) continue;

        // ステータスが「完了」または「中断」の場合はスキップ
        if (status === '完了' || status === '中断') continue;

        // 期日切れかどうかを判定
        const endDateObj = new Date(endDate);
        endDateObj.setHours(0, 0, 0, 0);

        if (endDateObj < today) {
          // 遅延日数を計算
          const daysOverdue = Math.floor((today - endDateObj) / (1000 * 60 * 60 * 24));

          // 進捗率を正規化
          if (typeof progressValue === 'number') {
            if (progressValue <= 1) {
              progressValue = progressValue * 100;
            }
          } else if (typeof progressValue === 'string') {
            progressValue = parseFloat(progressValue.replace('%', '')) || 0;
          }
          const progress = Math.round(progressValue);

          // シートIDを取得（プロジェクトIDを4桁にパディングしてマッチング）
          const projectIdPadded = String(projectId).padStart(4, '0');
          const sheetId = sheetIdMap[projectIdPadded] || 0;

          // 期日切れタスクデータを追加
          overdueTasksData.push([
            projectId,
            projectName,
            taskId,
            parentTaskName,
            childTaskName,
            endDate,
            assignee,
            daysOverdue,
            progress / 100, // 0.0-1.0形式に正規化
            sheetId // HYPERLINKのため
          ]);
        }
      }

      // 期日切れタスクをシートに書き込み
      if (overdueTasksData.length > 0) {
        // 遅延日数の降順でソート
        overdueTasksData.sort((a, b) => b[7] - a[7]);

        const startRow = 2;

        // データを書き込む（シートIDを除く9列分）
        const dataToWrite = overdueTasksData.map(row => row.slice(0, 9));
        sheet.getRange(startRow, 1, dataToWrite.length, dataToWrite[0].length).setValues(dataToWrite);

        // プロジェクトIDセルにHYPERLINKを設定
        for (let i = 0; i < overdueTasksData.length; i++) {
          const projectId = overdueTasksData[i][0];
          const sheetId = overdueTasksData[i][9]; // シートID
          const spreadsheetId = this.ss.getId();
          const hyperlink = `=HYPERLINK("https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheetId}", "${projectId}")`;
          sheet.getRange(startRow + i, 1).setFormula(hyperlink);
        }

        // 書式設定
        this._applyOverdueTasksFormatting(sheet, startRow, dataToWrite.length);
      }

      Logger.log(`✓ 期日切れタスク更新完了: ${overdueTasksData.length}件`);
      return overdueTasksData.length;

    } catch (error) {
      Logger.log(`✗ 期日切れタスク更新エラー: ${error.message}`);
      throw error;
    }
  }

  /**
   * 期日切れタスクシートに書式を適用
   */
  _applyOverdueTasksFormatting(sheet, startRow, numRows) {
    // 遅延日数列（H列 = 8列目）を赤色で強調
    const delayCol = sheet.getRange(startRow, 8, numRows, 1);
    delayCol.setBackground('#FFCDD2');
    delayCol.setFontColor('#B71C1C');
    delayCol.setFontWeight('bold');
    delayCol.setHorizontalAlignment('center');

    // 進捗率列（I列 = 9列目）にパーセント表示とデータバー
    const progressCol = sheet.getRange(startRow, 9, numRows, 1);
    progressCol.setNumberFormat('0%');

    let rules = sheet.getConditionalFormatRules();
    const ruleProgress = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#4A90E2', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([progressCol])
      .build();
    rules.push(ruleProgress);

    // 期日列（F列 = 6列目）を日付フォーマット
    const dateCol = sheet.getRange(startRow, 6, numRows, 1);
    dateCol.setNumberFormat('yy/mm/dd');

    sheet.setConditionalFormatRules(rules);
  }

  /**
   * Todoistタスクを更新
   * @param {Array} tasks - Todoist APIから取得したタスク配列
   */
  updateTodoistTasks(tasks) {
    try {
      const sheet = this.getOrCreateTodoistTasksSheet();

      // 既存データをクリア（ヘッダー以外）
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // 行削除ではなく範囲クリアを使用（固定行すべて削除エラー回避）
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      }

      // ヘッダーを再設定
      this._initializeTodoistTasksSheet(sheet);

      if (!tasks || tasks.length === 0) {
        Logger.log('Todoistタスクがありません。');
        return 0;
      }

      // タスクデータを整形
      const taskData = [];
      const now = new Date();

      for (const task of tasks) {
        const taskId = task.id || '';
        const taskName = task.content || '';
        const description = task.description || '';

        // 期日の処理
        let dueDate = '';
        if (task.due) {
          if (task.due.datetime) {
            dueDate = new Date(task.due.datetime);
          } else if (task.due.date) {
            dueDate = new Date(task.due.date);
          }
        }

        // 優先度の変換（TodoistのAPI: 1=通常, 2=中, 3=高, 4=最高 → 表示: 低/中/高/最高）
        let priority = '通常';
        if (task.priority === 4) priority = '最高';
        else if (task.priority === 3) priority = '高';
        else if (task.priority === 2) priority = '中';

        // ラベルの結合
        const labels = task.labels ? task.labels.join(', ') : '';

        // 完了状態
        const isCompleted = task.is_completed ? '完了' : '未完了';

        // Todoistリンク
        const todoistUrl = task.url || `https://todoist.com/app/task/${taskId}`;

        taskData.push([
          String(taskId),
          taskName,
          description,
          dueDate || '',
          priority,
          labels,
          isCompleted,
          todoistUrl,
          now
        ]);
      }

      // 期日で昇順ソート（期日が近いものが上）
      taskData.sort((a, b) => {
        const dateA = a[3]; // 期日列
        const dateB = b[3];

        // 期日が空の場合は最後に配置
        if (!dateA && !dateB) return 0;
        if (!dateA) return 1;
        if (!dateB) return -1;

        // 日付を比較
        return new Date(dateA) - new Date(dateB);
      });

      // シートに書き込む
      if (taskData.length > 0) {
        const startRow = 2;
        sheet.getRange(startRow, 1, taskData.length, 9).setValues(taskData);

        // TodoistリンクにHYPERLINK式を設定
        for (let i = 0; i < taskData.length; i++) {
          const url = taskData[i][7];
          const taskName = taskData[i][1];
          const hyperlink = `=HYPERLINK("${url}", "${taskName}")`;
          sheet.getRange(startRow + i, 8).setFormula(hyperlink);
        }

        // 書式設定
        this._applyTodoistTasksFormatting(sheet, startRow, taskData.length);
      }

      Logger.log(`✓ Todoistタスク更新完了: ${taskData.length}件`);
      return taskData.length;

    } catch (error) {
      Logger.log(`✗ Todoistタスク更新エラー: ${error.message}`);
      throw error;
    }
  }

  /**
   * Todoistタスクシートに書式を適用
   */
  _applyTodoistTasksFormatting(sheet, startRow, numRows) {
    // 期日列（D列 = 4列目）を日付フォーマット
    const dateCol = sheet.getRange(startRow, 4, numRows, 1);
    dateCol.setNumberFormat('yy/mm/dd');

    // 優先度列（E列 = 5列目）に条件付き書式
    const priorityCol = sheet.getRange(startRow, 5, numRows, 1);
    priorityCol.setHorizontalAlignment('center');

    let rules = sheet.getConditionalFormatRules();

    // 最高優先度（赤）
    const ruleHighest = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('最高')
      .setBackground('#FFCDD2')
      .setFontColor('#B71C1C')
      .setRanges([priorityCol])
      .build();
    rules.push(ruleHighest);

    // 高優先度（オレンジ）
    const ruleHigh = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('高')
      .setBackground('#FFE0B2')
      .setFontColor('#E65100')
      .setRanges([priorityCol])
      .build();
    rules.push(ruleHigh);

    // 中優先度（黄色）
    const ruleMedium = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('中')
      .setBackground('#FFF9C4')
      .setFontColor('#F57F17')
      .setRanges([priorityCol])
      .build();
    rules.push(ruleMedium);

    // 完了状態列（G列 = 7列目）に条件付き書式
    const statusCol = sheet.getRange(startRow, 7, numRows, 1);
    statusCol.setHorizontalAlignment('center');

    const ruleCompleted = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('完了')
      .setBackground('#C8E6C9')
      .setFontColor('#2E7D32')
      .setRanges([statusCol])
      .build();
    rules.push(ruleCompleted);

    // 最終更新日時列（I列 = 9列目）を日時フォーマット
    const timestampCol = sheet.getRange(startRow, 9, numRows, 1);
    timestampCol.setNumberFormat('yy/mm/dd hh:mm');

    sheet.setConditionalFormatRules(rules);
  }
}
