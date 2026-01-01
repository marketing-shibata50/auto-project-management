#!/usr/bin/env ts-node

/**
 * 既存のJSONファイルをGoogle Driveにアップロードするスクリプト
 */

import * as fs from 'fs';
import * as path from 'path';
import { google } from 'googleapis';
import * as dotenv from 'dotenv';

// 環境変数読み込み
dotenv.config();

/**
 * Google Driveにアップロード
 */
async function uploadToGoogleDrive(filePath: string): Promise<void> {
  try {
    const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
    if (!folderId) {
      throw new Error('GOOGLE_DRIVE_FOLDER_ID が設定されていません。');
    }

    // OAuth2認証設定
    const auth = new google.auth.OAuth2(
      process.env.GOOGLE_CLIENT_ID,
      process.env.GOOGLE_CLIENT_SECRET,
      process.env.GOOGLE_REDIRECT_URI
    );

    // トークン設定
    const tokenPath = path.join(process.cwd(), '.google-token.json');
    if (!fs.existsSync(tokenPath)) {
      throw new Error('Google認証トークンが見つかりません。先に npm run gantt:auth を実行してください。');
    }

    const token = JSON.parse(fs.readFileSync(tokenPath, 'utf-8'));
    auth.setCredentials(token);

    const drive = google.drive({ version: 'v3', auth });

    // ファイル名取得
    const fileName = path.basename(filePath);

    // アップロード履歴チェック
    const historyPath = path.join(process.cwd(), '.upload-history.json');
    let uploadHistory: Record<string, string> = {};

    if (fs.existsSync(historyPath)) {
      uploadHistory = JSON.parse(fs.readFileSync(historyPath, 'utf-8'));
    }

    if (uploadHistory[fileName]) {
      console.log(`⏭  スキップ: ${fileName} は既にアップロード済みです`);
      console.log(`   ファイルID: ${uploadHistory[fileName]}`);

      const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
      if (spreadsheetId) {
        const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
        console.log(`\n📊 スプレッドシート: ${spreadsheetUrl}`);
      }
      return;
    }

    // ファイルメタデータ
    const fileMetadata = {
      name: fileName,
      parents: [folderId]
    };

    const media = {
      mimeType: 'application/json',
      body: fs.createReadStream(filePath)
    };

    console.log(`📤 アップロード中: ${fileName}`);

    // アップロード実行
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
    throw error;
  }
}

/**
 * メイン処理
 */
async function main() {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.error('使い方: ts-node src/upload-to-drive.ts <ファイルパス>');
    console.error('例: ts-node src/upload-to-drive.ts 0027_お助けマンサービスHPの開発.json');
    process.exit(1);
  }

  const filePath = args[0];

  // 絶対パスでない場合、カレントディレクトリからの相対パスとして扱う
  const absolutePath = path.isAbsolute(filePath)
    ? filePath
    : path.join(process.cwd(), filePath);

  if (!fs.existsSync(absolutePath)) {
    console.error(`✗ ファイルが見つかりません: ${absolutePath}`);
    process.exit(1);
  }

  await uploadToGoogleDrive(absolutePath);
}

// スクリプト実行
if (require.main === module) {
  main().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}

export { uploadToGoogleDrive };
