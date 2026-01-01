#!/usr/bin/env ts-node

/**
 * JSONファイルのタグを最新の定義に更新するスクリプト
 */

import * as fs from 'fs';
import * as path from 'path';

// タグのマッピング定義
const TAG_MAPPING: Record<string, string> = {
  '企画': '企画・計画',
  '設計': '設計',
  '開発': '開発・実装',
  'デザイン': '設計', // UIデザインも設計カテゴリに含める
  'テスト': 'テスト・検証',
  'レビュー': 'テスト・検証', // レビューもテスト・検証に含める
  '資料作成': '事務・管理',
  '調整': '事務・管理',
  '運用': '運用・保守',
  'その他': 'その他'
};

/**
 * タグを更新する
 */
function updateTags(filePath: string): void {
  try {
    // JSONファイルを読み込み
    const jsonData = JSON.parse(fs.readFileSync(filePath, 'utf-8'));

    let updateCount = 0;

    // 各タスクのタグを更新
    for (const task of jsonData.tasks) {
      if (task.tags && Array.isArray(task.tags)) {
        const updatedTags = task.tags.map((tag: string) => {
          const newTag = TAG_MAPPING[tag] || tag;
          if (newTag !== tag) {
            updateCount++;
          }
          return newTag;
        });

        // 重複を削除
        task.tags = [...new Set(updatedTags)];
      }
    }

    // 更新したJSONを保存
    fs.writeFileSync(filePath, JSON.stringify(jsonData, null, 2), 'utf-8');

    console.log(`✓ タグ更新完了: ${filePath}`);
    console.log(`  更新されたタグ数: ${updateCount}`);
  } catch (error) {
    console.error('✗ エラー:', error);
    throw error;
  }
}

/**
 * メイン処理
 */
function main() {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.error('使い方: ts-node src/update-tags.ts <ファイルパス>');
    console.error('例: ts-node src/update-tags.ts 0027_お助けマンサービスHPの開発.json');
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

  updateTags(absolutePath);
}

// スクリプト実行
if (require.main === module) {
  main();
}

export { updateTags, TAG_MAPPING };
