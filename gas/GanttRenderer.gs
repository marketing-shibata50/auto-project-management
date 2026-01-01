/**
 * Ganttチャート描画クラス
 */

class GanttRenderer {
  constructor(sheet, projectData) {
    this.sheet = sheet;
    this.projectData = projectData;
    this.tasks = projectData.tasks;
  }
  
  /**
   * Ganttチャート全体を描画
   */
  render() {
    // プロジェクト期間を計算
    const dateRange = this._calculateDateRange();

    // 日付ヘッダーを描画（1行目）
    this._renderDateHeaders(dateRange);

    // タスク情報列のヘッダーを描画（2行目）
    this._renderTaskHeaders();

    // タスク情報を描画（4行目から）
    this._renderTaskInfo();

    // 入力規則とスタイルを設定（タスク行作成後に適用）
    this._applyStyles();

    // 行の破線を設定（タイムライン含む全列）
    this._applyRowBorders(dateRange);

    // Ganttバーを描画（実線枠線で破線を上書き）
    this._renderGanttBars(dateRange);

    // 不要な行を削除
    this._cleanupUnusedRows();
  }
  
  /**
   * プロジェクト期間を計算
   */
  _calculateDateRange() {
    const dates = this.tasks
      .filter(task => task.start_date && task.end_date)
      .flatMap(task => [new Date(task.start_date), new Date(task.end_date)]);
    
    if (dates.length === 0) {
      const today = new Date();
      return {
        start: today,
        end: new Date(today.getTime() + 90 * 24 * 60 * 60 * 1000) // 90日後
      };
    }
    
    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));
    
    // 前後に余裕を持たせる
    minDate.setDate(minDate.getDate() - 7);
    maxDate.setDate(maxDate.getDate() + 7);
    
    return { start: minDate, end: maxDate };
  }
  
  /**
   * 日付ヘッダーを描画（月と日）
   */
  _renderDateHeaders(dateRange) {
    const startCol = CONFIG.GANTT.START_COLUMN;

    // 日付を正規化（時刻部分を除去）して日数を計算
    const startDate = new Date(dateRange.start);
    startDate.setHours(0, 0, 0, 0);
    const endDate = new Date(dateRange.end);
    endDate.setHours(0, 0, 0, 0);

    const daysDiff = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;

    Logger.log(`日付範囲: ${dateRange.start} 〜 ${dateRange.end}`);
    Logger.log(`正規化後: ${startDate} 〜 ${endDate}`);
    Logger.log(`日数: ${daysDiff}日`);

    // Google Sheetsの列数制限チェック（18,278列）
    if (startCol + daysDiff > 18278) {
      throw new Error(`列数が上限を超えています: ${startCol + daysDiff} > 18278`);
    }

    const monthRow = [];
    const dayRow = [];
    let currentDate = new Date(startDate); // 正規化済みの開始日を使用
    const monthRanges = []; // 結合する範囲を記録

    let currentMonth = null;
    let monthStartCol = 0;

    // 月と日を分けて2行に（月は各月の最初のセルのみ表示、それ以降は空白）
    for (let i = 0; i < daysDiff; i++) {
      const month = currentDate.getMonth() + 1;
      const day = currentDate.getDate();

      // 月が変わったときのみ月を表示、それ以外は空白
      if (month !== currentMonth) {
        monthRow.push(month + '月');
        if (currentMonth !== null) {
          monthRanges.push({ start: monthStartCol, length: i - monthStartCol });
        }
        currentMonth = month;
        monthStartCol = i;
      } else {
        monthRow.push(''); // 空白
      }

      dayRow.push(day);
      currentDate.setDate(currentDate.getDate() + 1);
    }

    // 最後の月の結合範囲を記録
    monthRanges.push({ start: monthStartCol, length: daysDiff - monthStartCol });

    Logger.log(`日付行配列サイズ: 月=${monthRow.length}, 日=${dayRow.length}`);

    // ヘッダー設定（1行目：月、2行目：日）
    this.sheet.getRange(1, startCol, 1, daysDiff).setValues([monthRow]);
    this.sheet.getRange(2, startCol, 1, daysDiff).setValues([dayRow]);

    // 2色交互の色付け（セル結合は行わない）
    monthRanges.forEach((range, index) => {
      const rangeObj = this.sheet.getRange(1, startCol + range.start, 1, range.length);

      // 2色交互に色付け（#E8F4FD と #D6EAF8）
      const bgColor = index % 2 === 0 ? '#E8F4FD' : '#D6EAF8';
      rangeObj.setBackground(bgColor);
      rangeObj.setHorizontalAlignment('center');
      rangeObj.setVerticalAlignment('middle');

      // 日の行も同じ色で塗る
      this.sheet.getRange(2, startCol + range.start, 1, range.length)
        .setBackground(bgColor)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle');
    });

    // 列幅を調整
    for (let i = 0; i < daysDiff; i++) {
      this.sheet.setColumnWidth(startCol + i, CONFIG.GANTT.CELL_WIDTH);
    }

    // 固定行を設定（月1行 + 日1行 = 2行）
    this.sheet.setFrozenRows(2);
  }

  /**
   * タスク情報列のヘッダーを描画（2行目）
   */
  _renderTaskHeaders() {
    const headers = [
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

    const headerRange = this.sheet.getRange(2, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#E8F4FD');
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');

    // 1行目のA~L列にも同じ青色を適用
    const firstRowRange = this.sheet.getRange(1, 1, 1, headers.length);
    firstRowRange.setBackground('#E8F4FD');
    firstRowRange.setHorizontalAlignment('center');
    firstRowRange.setVerticalAlignment('middle');
  }

  /**
   * タスクが期日切れかどうかを判定
   */
  _isOverdue(task) {
    if (!task.end_date || task.progress === 100) return false;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const endDate = new Date(task.end_date);
    endDate.setHours(0, 0, 0, 0);

    const isOverdue = endDate < today;
    return isOverdue;
  }

  /**
   * タスク情報を描画
   */
  _renderTaskInfo() {
    const startRow = 4; // 月1行 + 日(列名)1行 + 空白行1行 + データ開始

    const taskData = this.tasks.map(task => {
      // 親タスク名と子タスク名を設定
      let parentTaskName = '';
      let childTaskName = '';

      if (task.parent_task_id) {
        // 子タスクの場合：B列は空白、C列に子タスク名
        parentTaskName = '';
        childTaskName = task.task_name;
      } else {
        // 親タスクの場合：B列に親タスク名、C列は空白
        parentTaskName = task.task_name;
        childTaskName = '';
      }

      // タグは配列の場合、「フェーズ」以外の最初の1つ
      let tag = '';
      if (Array.isArray(task.tags) && task.tags.length > 0) {
        // 「フェーズ」以外のタグを探す
        const validTags = task.tags.filter(t => t !== 'フェーズ');
        tag = validTags.length > 0 ? validTags[0] : '';
      } else if (task.tags && task.tags !== 'フェーズ') {
        tag = task.tags;
      }

      // ステータスを進捗率から判定
      const progress = task.progress || 0;
      let status = '';
      if (progress === 0) {
        status = '未着手';
      } else if (progress === 100) {
        status = '完了';
      } else {
        status = '進行中';
      }

      return [
        task.task_id,
        parentTaskName,
        childTaskName,
        tag,
        status,
        task.start_date || '',
        task.end_date || '',
        progress / 100, // 0.5 = 50% （条件付き書式用）
        task.assignee || '',
        task.priority || '',
        task.estimated_hours || '', // 数値のみ
        task.actual_hours || '' // 実績工数
      ];
    });

    if (taskData.length > 0) {
      this.sheet.getRange(startRow, 1, taskData.length, taskData[0].length).setValues(taskData);

      // 親タスク（parent_task_idがnull）を太字にする（1行全体）
      // タイムライン最後の列を計算
      const dateRange = this._calculateDateRange();

      // 日付を正規化して日数を計算
      const startDate = new Date(dateRange.start);
      startDate.setHours(0, 0, 0, 0);
      const endDate = new Date(dateRange.end);
      endDate.setHours(0, 0, 0, 0);
      const daysDiff = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;

      const lastCol = CONFIG.GANTT.START_COLUMN + daysDiff;

      // 全タスク行の背景色をクリア（過去の色付けをリセット）
      const allTasksRange = this.sheet.getRange(startRow, 1, this.tasks.length, lastCol);
      allTasksRange.setBackground(null);

      // 全タスクの終了日セル（G列）の背景色と文字色をリセット
      const endDateCol = this.sheet.getRange(startRow, 7, this.tasks.length, 1);
      endDateCol.setBackground('#FFFFFF'); // 白背景
      endDateCol.setFontColor('#000000'); // 黒文字
      endDateCol.setFontWeight('normal'); // 通常の太さ

      for (let i = 0; i < this.tasks.length; i++) {
        if (!this.tasks[i].parent_task_id) {
          // タスク行全体を太字に（A列からタイムライン最後まで）
          const rowRange = this.sheet.getRange(startRow + i, 1, 1, lastCol);
          rowRange.setFontWeight('bold');
        }

        // 期日切れタスクの終了日セル（G列）を強調
        if (this._isOverdue(this.tasks[i])) {
          const endDateCell = this.sheet.getRange(startRow + i, 7);
          endDateCell.setBackground('#FF0000'); // 赤背景
          endDateCell.setFontColor('#FFFF00'); // 黄色文字
          endDateCell.setFontWeight('bold'); // 太字
        }
      }

      // ステータス列（E列）に条件付き書式で色分け
      const statusCol = this.sheet.getRange(startRow, 5, taskData.length, 1);
      let rules = this.sheet.getConditionalFormatRules();

      // 未着手 = グレー
      const ruleNotStarted = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('未着手')
        .setBackground('#E0E0E0')
        .setRanges([statusCol])
        .build();
      rules.push(ruleNotStarted);

      // 進行中 = 黄色
      const ruleInProgress = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('進行中')
        .setBackground('#FFF59D')
        .setRanges([statusCol])
        .build();
      rules.push(ruleInProgress);

      // 完了 = 緑（E列のみ）
      const ruleCompleted = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('完了')
        .setBackground('#A5D6A7')
        .setRanges([statusCol])
        .build();
      rules.push(ruleCompleted);

      // 中断 = 赤（E列のみ）
      const ruleInterrupted = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('中断')
        .setBackground('#FF6B6B')
        .setRanges([statusCol])
        .build();
      rules.push(ruleInterrupted);

      // 完了・中断の行全体をグレーにする（全列対象）
      const allRowsRange = this.sheet.getRange(startRow, 1, taskData.length, lastCol);

      // 完了行全体 = グレー
      const ruleCompletedRow = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$E${startRow}="完了"`)
        .setBackground('#D3D3D3')
        .setRanges([allRowsRange])
        .build();
      rules.push(ruleCompletedRow);

      // 中断行全体 = グレー
      const ruleInterruptedRow = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=$E${startRow}="中断"`)
        .setBackground('#D3D3D3')
        .setRanges([allRowsRange])
        .build();
      rules.push(ruleInterruptedRow);

      // 日付列（F列・G列）のフォーマットを「yy/mm/dd」に設定
      const dateColStart = this.sheet.getRange(startRow, 6, taskData.length, 1); // F列（開始日）
      const dateColEnd = this.sheet.getRange(startRow, 7, taskData.length, 1); // G列（終了日）
      dateColStart.setNumberFormat('yy/mm/dd');
      dateColEnd.setNumberFormat('yy/mm/dd');

      // 進捗列（H列）に条件付き書式でバー表示
      const progressCol = this.sheet.getRange(startRow, 8, taskData.length, 1);
      progressCol.setNumberFormat('0%'); // パーセント表示

      // データバーの条件付き書式を追加
      const ruleProgress = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMaxpointWithValue('#4A90E2', SpreadsheetApp.InterpolationType.NUMBER, '1')
        .setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setRanges([progressCol])
        .build();
      rules.push(ruleProgress);

      this.sheet.setConditionalFormatRules(rules);
    }
  }
  
  /**
   * Ganttバーを描画
   */
  _renderGanttBars(dateRange) {
    const startCol = CONFIG.GANTT.START_COLUMN;
    const startRow = 4; // タスクデータ開始行に合わせる

    this.tasks.forEach((task, index) => {
      if (!task.start_date || !task.end_date) return;

      const taskStart = new Date(task.start_date);
      const taskEnd = new Date(task.end_date);

      // タスクの開始位置と期間を計算
      const daysFromStart = Math.ceil((taskStart - dateRange.start) / (1000 * 60 * 60 * 24));
      const duration = Math.ceil((taskEnd - taskStart) / (1000 * 60 * 60 * 24)) + 1;

      if (daysFromStart >= 0 && duration > 0) {
        const barRange = this.sheet.getRange(
          startRow + index,
          startCol + daysFromStart,
          1,
          duration
        );

        // 親=赤、子=青で色分け
        const color = task.parent_task_id ? '#4A90E2' : '#FF6B6B'; // 子=青、親=赤
        const progress = task.progress || 0;
        
        // バーの背景色を設定
        barRange.setBackground(color);
        
        // 進捗バーを描画（グラデーション風）
        if (progress < 100) {
          const progressCols = Math.ceil(duration * progress / 100);
          if (progressCols > 0) {
            const progressRange = this.sheet.getRange(
              startRow + index,
              startCol + daysFromStart,
              1,
              progressCols
            );
            progressRange.setBackground(this._darkenColor(color));
          }
        }
        
        // マイルストーンの場合はダイヤモンド記号を追加
        if (task.is_milestone) {
          this.sheet.getRange(startRow + index, startCol + daysFromStart).setValue('◆');
        }
        
        // セルに罫線を追加
        barRange.setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    });
  }
  
  /**
   * タスクの色を取得
   */
  _getTaskColor(task) {
    const tags = Array.isArray(task.tags) ? task.tags : [task.tags];
    const firstTag = tags[0];
    
    return CONFIG.TAG_COLORS[firstTag] || CONFIG.TAG_COLORS['デフォルト'];
  }
  
  /**
   * 色を暗くする（進捗表示用）
   */
  _darkenColor(hexColor) {
    // HEXからRGBに変換
    const r = parseInt(hexColor.slice(1, 3), 16);
    const g = parseInt(hexColor.slice(3, 5), 16);
    const b = parseInt(hexColor.slice(5, 7), 16);
    
    // 70%の明度に
    const newR = Math.floor(r * 0.7);
    const newG = Math.floor(g * 0.7);
    const newB = Math.floor(b * 0.7);
    
    // RGBからHEXに変換
    return '#' + [newR, newG, newB].map(x => x.toString(16).padStart(2, '0')).join('');
  }
  
  /**
   * 期間を計算
   */
  _calculateDuration(startDate, endDate) {
    if (!startDate || !endDate) return '';

    const start = new Date(startDate);
    const end = new Date(endDate);
    const days = Math.ceil((end - start) / (1000 * 60 * 60 * 24)) + 1;

    return days; // 数値のみ返す
  }
  
  /**
   * スタイルを適用
   */
  _applyStyles() {
    const dataRange = this.sheet.getDataRange();
    dataRange.setVerticalAlignment('middle');

    // グリッドを非表示
    this.sheet.setHiddenGridlines(true);

    // ID（A列）・親タスク名（B列）・子タスク名（C列）の列幅を自動調整
    this.sheet.autoResizeColumn(1); // A列（ID）
    this.sheet.autoResizeColumn(2); // B列（親タスク名）
    this.sheet.autoResizeColumn(3); // C列（子タスク名）

    // 行の高さを25に設定（4行目以降のタスクデータ行）
    if (this.tasks.length > 0) {
      this.sheet.setRowHeights(4, this.tasks.length, 25);

      // 親タスク名・子タスク名の列を左寄せ
      const parentTaskNameCol = this.sheet.getRange(4, 2, this.tasks.length, 1);
      parentTaskNameCol.setHorizontalAlignment('left');
      const childTaskNameCol = this.sheet.getRange(4, 3, this.tasks.length, 1);
      childTaskNameCol.setHorizontalAlignment('left');

      // タグ列に入力規則を設定
      const tagOptions = [
        '企画・計画',
        '設計',
        '開発・実装',
        'テスト・検証',
        'リリース・デプロイ',
        '運用・保守',
        'マーケティング',
        '営業・商談',
        '事務・管理',
        'その他'
      ];
      const tagCol = this.sheet.getRange(4, 4, this.tasks.length, 1); // D列（タグ）
      const tagRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(tagOptions, true) // true = ドロップダウン表示
        .setAllowInvalid(true) // 選択肢以外の値も許可
        .build();
      tagCol.setDataValidation(tagRule);

      // ステータス列に入力規則を設定
      const statusOptions = [
        '未着手',
        '進行中',
        '完了',
        '中断'
      ];
      const statusCol = this.sheet.getRange(4, 5, this.tasks.length, 1); // E列（ステータス）
      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(statusOptions, true) // true = ドロップダウン表示
        .setAllowInvalid(false)
        .build();
      statusCol.setDataValidation(statusRule);

      // タグ列に色付き条件付き書式を設定
      const tagColors = CONFIG.TAG_COLORS;

      // まずタグ列全体にデフォルト色を適用
      tagCol.setBackground(tagColors['デフォルト']);
      tagCol.setFontColor('#FFFFFF');

      let rules = this.sheet.getConditionalFormatRules();
      Object.keys(tagColors).forEach(tag => {
        // 'デフォルト'はスキップ（特定のタグ名ではないため）
        if (tag === 'デフォルト') return;

        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(tag)
          .setBackground(tagColors[tag])
          .setFontColor('#FFFFFF') // 白文字
          .setRanges([tagCol])
          .build();
        rules.push(rule);
      });
      this.sheet.setConditionalFormatRules(rules);
    }
  }

  /**
   * 行の破線を適用（全列）
   */
  _applyRowBorders(dateRange) {
    if (this.tasks.length === 0) return;

    // タイムライン列の範囲を計算
    const daysDiff = Math.ceil((dateRange.end - dateRange.start) / (1000 * 60 * 60 * 24));
    const lastCol = CONFIG.GANTT.START_COLUMN + daysDiff;

    // 最初のタスク行（4行目）の上に破線を追加
    const firstRowRange = this.sheet.getRange(4, 1, 1, lastCol);
    firstRowRange.setBorder(
      true, null, null, null, // top, left, bottom, right
      null, null, // vertical, horizontal
      '#CCCCCC', // グレー
      SpreadsheetApp.BorderStyle.DASHED // 破線
    );

    // 各タスク行の下に破線を追加（全列）
    for (let i = 0; i < this.tasks.length; i++) {
      const rowRange = this.sheet.getRange(4 + i, 1, 1, lastCol);
      rowRange.setBorder(
        null, null, true, null, // top, left, bottom, right
        null, null, // vertical, horizontal
        '#CCCCCC', // グレー
        SpreadsheetApp.BorderStyle.DASHED // 破線
      );
    }
  }

  /**
   * 不要な行を削除（タスクデータの最後の行から1行空けてそれ以降を削除）
   */
  _cleanupUnusedRows() {
    if (this.tasks.length === 0) return;

    // タスクデータの最後の行 = 3行（ヘッダー） + タスク数
    const lastTaskRow = 3 + this.tasks.length;

    // 1行空けた次の行から削除開始
    const deleteStartRow = lastTaskRow + 2;

    // シートの最大行数を取得
    const maxRows = this.sheet.getMaxRows();

    // 削除する行数を計算
    const rowsToDelete = maxRows - deleteStartRow + 1;

    // 削除する行が存在する場合のみ削除
    if (rowsToDelete > 0 && deleteStartRow <= maxRows) {
      this.sheet.deleteRows(deleteStartRow, rowsToDelete);
    }
  }
}

/**
 * ブロックされているタスクをGanttシートでハイライト
 *
 * 遅延タスクの影響を受けているタスクを赤色背景でハイライトし、
 * セルにメモを追加します。
 *
 * @param {Sheet} sheet - Ganttシート
 * @param {Object} projectData - プロジェクトデータ
 * @param {Object} impactReport - analyzeDependencyImpact()の戻り値
 */
function highlightBlockedTasks(sheet, projectData, impactReport) {
  try {
    // 影響を受けるタスクIDをすべて収集
    const allImpactedTaskIds = new Set();

    for (const delayedTaskId in impactReport.impactedTasks) {
      const impact = impactReport.impactedTasks[delayedTaskId];
      for (const taskId of impact.impactedTaskIds) {
        allImpactedTaskIds.add(taskId);
      }
    }

    if (allImpactedTaskIds.size === 0) {
      Logger.log('✓ ブロックされているタスクがないため、ハイライトをスキップ');
      return;
    }

    Logger.log(`--- ${allImpactedTaskIds.size}件のブロックタスクをハイライト ---`);

    // タスクIDから行番号を取得（4行目から開始）
    const taskMap = {};
    for (let i = 0; i < projectData.tasks.length; i++) {
      const task = projectData.tasks[i];
      taskMap[task.task_id] = 4 + i; // 4行目からタスクデータが始まる
    }

    // ブロックされているタスクの行をハイライト
    for (const taskId of allImpactedTaskIds) {
      const rowNum = taskMap[taskId];

      if (!rowNum) {
        Logger.log(`⚠ タスクID ${taskId} の行が見つかりません`);
        continue;
      }

      // タスク情報列（A~L列の12列）を赤色背景でハイライト
      const range = sheet.getRange(rowNum, 1, 1, 12);
      range.setBackground('#FFCDD2'); // 赤色背景

      // タスク名セル（C列）にメモを追加
      const taskNameCell = sheet.getRange(rowNum, 3);
      const currentNote = taskNameCell.getNote();
      const blockerNote = '⚠️ ブロック中: 依存タスクが遅延しています';

      if (!currentNote.includes(blockerNote)) {
        const newNote = currentNote ? currentNote + '\n\n' + blockerNote : blockerNote;
        taskNameCell.setNote(newNote);
      }

      Logger.log(`  ハイライト完了: ${taskId} (行${rowNum})`);
    }

    Logger.log(`✓ ${allImpactedTaskIds.size}件のタスクをハイライトしました`);

  } catch (error) {
    Logger.log(`✗ ハイライトエラー: ${error.message}`);
    Logger.log(`✗ スタックトレース: ${error.stack}`);
    throw error;
  }
}

/**
 * クリティカルパス上のタスクをGanttシート上でハイライト表示
 *
 * @param {Sheet} sheet - Ganttシート
 * @param {Object} projectData - プロジェクトデータ
 * @param {Object} criticalPathReport - calculateCriticalPath()の結果
 */
function highlightCriticalPath(sheet, projectData, criticalPathReport) {
  if (!criticalPathReport || criticalPathReport.skipped) {
    Logger.log('[INFO] Critical path analysis skipped or unavailable');
    return;
  }

  if (criticalPathReport.error === 'circular_dependency') {
    Logger.log('[ERROR] Cannot highlight critical path due to circular dependency');
    return;
  }

  const criticalTaskIds = new Set(criticalPathReport.criticalTasks);
  const nearCriticalTaskIds = new Set(criticalPathReport.nearCriticalTasks);

  // タスクID → 行番号のマップを構築（4行目から開始）
  const taskMap = {};
  for (let i = 0; i < projectData.tasks.length; i++) {
    taskMap[projectData.tasks[i].task_id] = 4 + i;
  }

  // クリティカルパスタスクを赤色ハイライト
  for (const taskId of criticalTaskIds) {
    if (!taskMap[taskId]) continue;

    const rowNum = taskMap[taskId];
    const range = sheet.getRange(rowNum, 1, 1, 12);  // A列～L列
    range.setBackground('#FFEBEE');  // Light red background

    // タスク名セルに太字＋赤色テキスト
    const taskNameCell = sheet.getRange(rowNum, 3);  // C列（タスク名）
    taskNameCell.setFontWeight('bold');
    taskNameCell.setFontColor('#E53935');  // Red text

    // ノートを追加
    const metrics = criticalPathReport.taskMetrics[taskId];
    const noteText = '🔴 クリティカルパス\n' +
                     'スラック: 0日\n' +
                     '最早開始: Day ' + metrics.es + '\n' +
                     '最早終了: Day ' + metrics.ef;
    taskNameCell.setNote(noteText);
  }

  // Near-criticalタスクをオレンジ色ハイライト
  for (const taskId of nearCriticalTaskIds) {
    if (!taskMap[taskId]) continue;

    const rowNum = taskMap[taskId];
    const range = sheet.getRange(rowNum, 1, 1, 12);
    range.setBackground('#FFF3E0');  // Light orange background

    const taskNameCell = sheet.getRange(rowNum, 3);
    taskNameCell.setFontWeight('bold');
    taskNameCell.setFontColor('#FF9800');  // Orange text

    const metrics = criticalPathReport.taskMetrics[taskId];
    const noteText = '🟠 Near-Critical\n' +
                     'スラック: ' + metrics.slack + '日\n' +
                     '最早開始: Day ' + metrics.es + '\n' +
                     '最遅開始: Day ' + metrics.ls;
    taskNameCell.setNote(noteText);
  }

  Logger.log('[INFO] Highlighted ' + criticalTaskIds.size + ' critical and ' +
             nearCriticalTaskIds.size + ' near-critical tasks');
}
