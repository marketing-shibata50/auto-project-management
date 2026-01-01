#!/usr/bin/env ts-node

/**
 * Ganttチャート自動生成ヘルパースクリプト
 *
 * 対話の状態管理、JSON生成、Google Driveアップロード、対話履歴保存を担当
 */

import * as fs from 'fs';
import * as path from 'path';
import { google } from 'googleapis';
import * as dotenv from 'dotenv';
import * as https from 'https';
import * as http from 'http';
import * as url from 'url';

// 環境変数読み込み
dotenv.config();

// 型定義
interface Task {
  task_id: string;
  task_name: string;
  start_date: string;
  end_date: string;
  assignee: string;
  dependencies: string[];
  progress: number;
  priority: string;
  parent_task_id: string | null;
  tags: string[];
  estimated_hours: number;
  is_milestone: boolean;
}

interface ProjectData {
  project_id: string;
  project_name: string;
  project_purpose: string;
  project_type: string;
  project_deadline: string;
  github_url?: string;
  tasks: Task[];
}

interface DialogueEntry {
  timestamp: string;
  speaker: 'user' | 'assistant';
  message: string;
}

/**
 * 対話状態管理クラス
 */
class GanttDialogueManager {
  private projectData: Partial<ProjectData> = {};
  private dialogueHistory: DialogueEntry[] = [];
  private currentPhase: number = 1;
  private tasks: Task[] = [];
  private taskCounter: number = 1;

  /**
   * 対話履歴を追加
   */
  addDialogue(speaker: 'user' | 'assistant', message: string): void {
    this.dialogueHistory.push({
      timestamp: new Date().toISOString(),
      speaker,
      message
    });
  }

  /**
   * プロジェクト基本情報を設定
   */
  setProjectInfo(key: keyof ProjectData, value: any): void {
    (this.projectData as any)[key] = value;
  }

  /**
   * タスクを追加
   */
  addTask(task: Omit<Task, 'task_id'>): string {
    const taskId = `T${this.taskCounter.toString().padStart(3, '0')}`;
    this.tasks.push({
      task_id: taskId,
      ...task
    });
    this.taskCounter++;
    return taskId;
  }

  /**
   * タスクを更新
   */
  updateTask(taskId: string, updates: Partial<Task>): void {
    const taskIndex = this.tasks.findIndex(t => t.task_id === taskId);
    if (taskIndex !== -1) {
      this.tasks[taskIndex] = { ...this.tasks[taskIndex], ...updates };
    }
  }

  /**
   * 現在のフェーズを取得
   */
  getCurrentPhase(): number {
    return this.currentPhase;
  }

  /**
   * 次のフェーズに進む
   */
  nextPhase(): void {
    this.currentPhase++;
  }

  /**
   * JSONファイルを生成
   */
  generateJSON(): ProjectData {
    return {
      project_id: this.projectData.project_id || '',
      project_name: this.projectData.project_name || '',
      project_purpose: this.projectData.project_purpose || '',
      project_type: this.projectData.project_type || '',
      project_deadline: this.projectData.project_deadline || '',
      github_url: this.projectData.github_url || '',
      tasks: this.tasks
    };
  }

  /**
   * 対話履歴をMarkdown形式で生成
   */
  generateDialogueMarkdown(): string {
    let markdown = `# プロジェクト作成対話履歴\n\n`;
    markdown += `**プロジェクト名**: ${this.projectData.project_name || '未設定'}\n`;
    markdown += `**プロジェクトID**: ${this.projectData.project_id || '未設定'}\n`;
    markdown += `**作成日時**: ${new Date().toISOString()}\n\n`;
    markdown += `---\n\n`;

    for (const entry of this.dialogueHistory) {
      const speaker = entry.speaker === 'user' ? 'ユーザー' : 'アシスタント';
      const time = new Date(entry.timestamp).toLocaleString('ja-JP');
      markdown += `## ${speaker} (${time})\n\n`;
      markdown += `${entry.message}\n\n`;
    }

    return markdown;
  }

  /**
   * ファイルを保存
   * @param isUpdate - 既存プロジェクトの更新かどうか（デフォルト: false = 新規作成）
   */
  async saveFiles(isUpdate: boolean = false): Promise<void> {
    try {
      const projectId = this.projectData.project_id || 'unknown';
      const projectName = this.projectData.project_name || 'unknown';
      const prefix = isUpdate ? 'update_' : 'new_';
      const fileName = `${prefix}${projectId}_${projectName}`;

      // JSONファイル保存（outputs/ディレクトリに保存）
      const outputsDir = path.join(process.cwd(), 'outputs');
      if (!fs.existsSync(outputsDir)) {
        fs.mkdirSync(outputsDir, { recursive: true });
      }
      const jsonData = this.generateJSON();
      const jsonPath = path.join(outputsDir, `${fileName}.json`);
      fs.writeFileSync(jsonPath, JSON.stringify(jsonData, null, 2), 'utf-8');
      console.log(`✓ JSONファイル保存: ${jsonPath}`);

      // 対話履歴保存
      const markdown = this.generateDialogueMarkdown();
      const docsDir = path.join(process.cwd(), 'docs');
      if (!fs.existsSync(docsDir)) {
        fs.mkdirSync(docsDir, { recursive: true });
      }
      const mdPath = path.join(docsDir, `${fileName}.md`);
      fs.writeFileSync(mdPath, markdown, 'utf-8');
      console.log(`✓ 対話履歴保存: ${mdPath}`);

      // Google Driveアップロード
      await this.uploadToGoogleDrive(jsonPath);

      // スプレッドシート情報の表示
      const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
      if (spreadsheetId) {
        const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
        console.log(`\n📊 スプレッドシートURL: ${spreadsheetUrl}`);
        console.log(`   ※ GASトリガーでGanttチャートが自動生成されます`);
      }

      // Discord通知（成功）
      await this.sendDiscordNotification(
        `プロジェクト「${projectName}」のJSONファイルが正常に生成されました。\n` +
        `JSONファイル: \`outputs/${fileName}.json\`\n` +
        `対話履歴: \`docs/${fileName}.md\`\n` +
        (spreadsheetId ? `\nスプレッドシート: https://docs.google.com/spreadsheets/d/${spreadsheetId}` : ''),
        false
      );
    } catch (error) {
      // Discord通知（エラー）
      await this.sendDiscordNotification(
        `プロジェクトファイルの保存中にエラーが発生しました。\n\n` +
        `エラー内容: ${error instanceof Error ? error.message : String(error)}`,
        true
      );
      throw error;
    }
  }

  /**
   * Google Driveにアップロード（重複チェック機能付き）
   */
  private async uploadToGoogleDrive(filePath: string): Promise<void> {
    try {
      const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
      if (!folderId) {
        console.warn('⚠ GOOGLE_DRIVE_FOLDER_ID が設定されていません。アップロードをスキップします。');
        return;
      }

      // アップロード履歴チェック
      const fileName = path.basename(filePath);
      const historyPath = path.join(process.cwd(), '.upload-history.json');
      let uploadHistory: Record<string, string> = {};

      if (fs.existsSync(historyPath)) {
        uploadHistory = JSON.parse(fs.readFileSync(historyPath, 'utf-8'));
      }

      if (uploadHistory[fileName]) {
        console.log(`⏭  スキップ: ${fileName} は既にアップロード済みです`);
        console.log(`   ファイルID: ${uploadHistory[fileName]}`);
        console.log(`\n⏰ GAS側で1分以内に自動的にGanttチャートが生成されます。`);

        const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
        if (spreadsheetId) {
          const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
          console.log(`📊 スプレッドシート: ${spreadsheetUrl}`);
        }
        return;
      }

      // OAuth2認証設定（環境変数から読み込み想定）
      const auth = new google.auth.OAuth2(
        process.env.GOOGLE_CLIENT_ID,
        process.env.GOOGLE_CLIENT_SECRET,
        process.env.GOOGLE_REDIRECT_URI
      );

      // トークン設定（事前に取得したトークンを使用）
      const tokenPath = path.join(process.cwd(), '.google-token.json');
      if (fs.existsSync(tokenPath)) {
        const token = JSON.parse(fs.readFileSync(tokenPath, 'utf-8'));
        auth.setCredentials(token);
      } else {
        console.warn('⚠ Google認証トークンが見つかりません。初回認証が必要です。');
        // TODO: 初回認証フローの実装
        return;
      }

      const drive = google.drive({ version: 'v3', auth });

      const fileMetadata = {
        name: fileName,
        parents: [folderId]
      };

      const media = {
        mimeType: 'application/json',
        body: fs.createReadStream(filePath)
      };

      const response = await drive.files.create({
        requestBody: fileMetadata,
        media: media,
        fields: 'id, name, webViewLink'
      });

      const fileId = response.data.id!;
      console.log(`✓ Google Driveにアップロード完了`);
      console.log(`  ファイルID: ${fileId}`);
      console.log(`  URL: ${response.data.webViewLink}`);
      console.log(`\n⏰ GAS側で1分以内に自動的にGanttチャートが生成されます。`);

      const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
      if (spreadsheetId) {
        const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
        console.log(`📊 スプレッドシート: ${spreadsheetUrl}`);
      }

      // アップロード履歴を保存
      uploadHistory[fileName] = fileId;
      fs.writeFileSync(historyPath, JSON.stringify(uploadHistory, null, 2), 'utf-8');

    } catch (error) {
      console.error('✗ Google Driveアップロードエラー:', error);

      // Discord通知（Google Driveエラー）
      await this.sendDiscordNotification(
        `Google Driveへのアップロード中にエラーが発生しました。\n\n` +
        `エラー内容: ${error instanceof Error ? error.message : String(error)}\n\n` +
        `ファイル: \`${path.basename(filePath)}\``,
        true
      );

      throw error;
    }
  }

  /**
   * 状態をJSON形式で取得（デバッグ用）
   */
  getState(): object {
    return {
      currentPhase: this.currentPhase,
      projectData: this.projectData,
      tasksCount: this.tasks.length,
      dialogueHistoryCount: this.dialogueHistory.length
    };
  }

  /**
   * Discord Webhookに通知を送信
   */
  private async sendDiscordNotification(message: string, isError: boolean = false): Promise<void> {
    const webhookUrl = process.env.DISCORD_WEBHOOK_URL;
    if (!webhookUrl) {
      console.warn('⚠ DISCORD_WEBHOOK_URL が設定されていません。通知をスキップします。');
      return;
    }

    try {
      const url = new URL(webhookUrl);
      const payload = JSON.stringify({
        embeds: [{
          title: isError ? '❌ Ganttチャート作成エラー' : '✅ Ganttチャート作成完了',
          description: message,
          color: isError ? 0xff0000 : 0x00ff00,
          timestamp: new Date().toISOString(),
          footer: {
            text: 'Gantt Chart Generator'
          },
          fields: [
            {
              name: 'プロジェクトID',
              value: this.projectData.project_id || '未設定',
              inline: true
            },
            {
              name: 'プロジェクト名',
              value: this.projectData.project_name || '未設定',
              inline: true
            },
            {
              name: 'タスク数',
              value: this.tasks.length.toString(),
              inline: true
            }
          ]
        }]
      });

      const options = {
        hostname: url.hostname,
        path: url.pathname + url.search,
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Content-Length': Buffer.byteLength(payload)
        }
      };

      await new Promise<void>((resolve, reject) => {
        const req = https.request(options, (res) => {
          if (res.statusCode === 204) {
            console.log('✓ Discord通知送信完了');
            resolve();
          } else {
            reject(new Error(`Discord通知失敗: ${res.statusCode}`));
          }
        });

        req.on('error', (error) => {
          reject(error);
        });

        req.write(payload);
        req.end();
      });
    } catch (error) {
      console.error('✗ Discord通知送信エラー:', error);
      // Discord通知の失敗は致命的エラーではないため、処理を続行
    }
  }
}

/**
 * Google Drive OAuth2 初回認証
 */
async function authenticateGoogleDrive(): Promise<void> {
  const clientId = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const redirectUri = process.env.GOOGLE_REDIRECT_URI || 'http://localhost:19204/oauth2callback';

  if (!clientId || !clientSecret) {
    throw new Error('GOOGLE_CLIENT_ID と GOOGLE_CLIENT_SECRET を .env ファイルに設定してください。');
  }

  const oauth2Client = new google.auth.OAuth2(clientId, clientSecret, redirectUri);

  // 認証URLを生成
  const authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    prompt: 'consent',
    scope: [
      'https://www.googleapis.com/auth/drive.file'
    ]
  });

  console.log('\n🔐 Google Drive OAuth2 認証を開始します\n');
  console.log('以下のURLをブラウザで開いてください:');
  console.log(`\n${authUrl}\n`);

  // ローカルサーバーを起動してコールバックを受け取る
  return new Promise((resolve, reject) => {
    const server = http.createServer(async (req, res) => {
      try {
        const reqUrl = url.parse(req.url || '', true);

        if (reqUrl.pathname === '/oauth2callback') {
          const code = reqUrl.query.code as string;

          if (!code) {
            res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
            res.end('<h1>認証失敗</h1><p>認証コードが取得できませんでした。</p>');
            reject(new Error('認証コードが取得できませんでした'));
            server.close();
            return;
          }

          // トークン取得
          const { tokens } = await oauth2Client.getToken(code);
          oauth2Client.setCredentials(tokens);

          // トークンをファイルに保存
          const tokenPath = path.join(process.cwd(), '.google-token.json');
          fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2), 'utf-8');

          res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
          res.end(`
            <html>
              <head><title>認証成功</title></head>
              <body style="font-family: sans-serif; text-align: center; padding: 50px;">
                <h1 style="color: #4CAF50;">✅ 認証成功！</h1>
                <p>Google Drive との連携が完了しました。</p>
                <p>このウィンドウを閉じてターミナルに戻ってください。</p>
              </body>
            </html>
          `);

          console.log('\n✓ Google Drive 認証完了');
          console.log(`✓ トークン保存: ${tokenPath}`);
          console.log('\nこれで Google Drive へのアップロードが可能になりました。');

          server.close();
          resolve();
        }
      } catch (error) {
        res.writeHead(500, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end('<h1>エラー</h1><p>認証処理中にエラーが発生しました。</p>');
        reject(error);
        server.close();
      }
    });

    server.listen(19204, () => {
      console.log('ローカル認証サーバーを起動しました (http://localhost:19204)');
      console.log('ブラウザでログイン後、自動的にトークンが保存されます...\n');
    });

    server.on('error', (error) => {
      reject(error);
    });
  });
}

/**
 * コマンドライン引数パーサー
 */
function parseArgs(): { command: string; isUpdate: boolean } {
  const args = process.argv.slice(2);
  const command = args[0] || 'help';
  const isUpdate = args.includes('--update');

  return { command, isUpdate };
}

/**
 * メイン処理
 */
async function main() {
  const { command, isUpdate } = parseArgs();

  const manager = new GanttDialogueManager();

  switch (command) {
    case 'init':
      console.log('Ganttチャート作成を開始します...');
      console.log('対話を進めてください。完了後に `save` コマンドを実行してください。');
      break;

    case 'save':
      if (isUpdate) {
        console.log('既存プロジェクトを更新しています...');
      } else {
        console.log('新規プロジェクトを作成しています...');
      }
      await manager.saveFiles(isUpdate);
      console.log('✓ すべてのファイルが保存されました');
      break;

    case 'auth':
      console.log('Google Drive OAuth2 認証を開始します...');
      await authenticateGoogleDrive();
      break;

    case 'status':
      console.log('現在の状態:');
      console.log(JSON.stringify(manager.getState(), null, 2));
      break;

    case 'help':
    default:
      console.log(`
Ganttチャート自動生成ヘルパー

使い方:
  npm run gantt:<command> [options]

コマンド:
  init    - 新規プロジェクト作成開始
  save    - JSONと対話履歴を保存（自動的にGoogle Driveにアップロード＆Ganttチャート生成）
            オプション:
              --update  既存プロジェクトの更新として保存（update_ プレフィックス付き）
              省略時    新規プロジェクトとして保存（new_ プレフィックス付き）
  auth    - Google Drive OAuth2 初回認証
  status  - 現在の状態を表示
  help    - このヘルプを表示

使用例:
  npm run gantt:save           # 新規プロジェクト作成（new_192_プロジェクト.json）
  npm run gantt:save -- --update  # 既存プロジェクト更新（update_192_プロジェクト.json）

初回セットアップ手順:
  1. .env ファイルに以下を設定:
     GOOGLE_CLIENT_ID=your_client_id
     GOOGLE_CLIENT_SECRET=your_client_secret
     GOOGLE_REDIRECT_URI=http://localhost:19204/oauth2callback
     GOOGLE_DRIVE_FOLDER_ID=your_folder_id
     GOOGLE_SCRIPT_ID=your_script_id (GASのスクリプトID)
     GOOGLE_SPREADSHEET_ID=your_spreadsheet_id

  2. npm run gantt:auth を実行してGoogle認証

  3. /gantt コマンドでプロジェクト作成

  4. npm run gantt:save で保存＆アップロード
     → 自動的にGanttチャートが生成されます
      `);
      break;
  }
}

// スクリプト実行
if (require.main === module) {
  main().catch(console.error);
}

export { GanttDialogueManager, ProjectData, Task, DialogueEntry };
