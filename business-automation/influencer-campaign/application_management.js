/**
 * インフルエンサー施策 - 応募管理スプレッドシート設定
 *
 * 【セットアップ手順】
 * 1. 新しいスプレッドシートを作成（または案件管理と同じシートに別シートとして追加）
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付けて保存
 * 4. setupApplicationSheet() を実行（初回のみ）
 */

// ============================================
// 設定
// ============================================
const APPLICATION_CONFIG = {
  // シート名
  SHEET_NAME: '応募管理',

  // 案件管理シート名（同じスプレッドシート内にある場合）
  PROJECT_SHEET_NAME: '案件管理',

  // ヘッダー行
  HEADER_ROW: 1,

  // データ開始行
  DATA_START_ROW: 2,

  // 応募ステータス選択肢
  STATUS_OPTIONS: ['新規', '確認中', '承認', '見送り', '完了'],

  // 通知先メールアドレス（クライアント企業担当者）
  NOTIFICATION_EMAIL: 'YOUR_EMAIL@example.com',

  // ヘッダー色
  HEADER_COLOR: '#11998e',
  HEADER_FONT_COLOR: '#ffffff',

  // ステータス別色
  STATUS_COLORS: {
    '新規': '#fff2cc',
    '確認中': '#cfe2f3',
    '承認': '#d9ead3',
    '見送り': '#f4cccc',
    '完了': '#d9d9d9'
  }
};

// ============================================
// カラム定義（15列）
// ============================================
const APPLICATION_COLUMNS = [
  // 基本情報
  { name: '応募ID', width: 100 },
  { name: 'タイムスタンプ', width: 150 },
  { name: '応募ステータス', width: 100 },

  // 案件情報
  { name: '案件ID', width: 120 },
  { name: '店舗名', width: 150 },

  // インフルエンサー情報
  { name: 'インフルエンサー名', width: 150 },
  { name: 'メールアドレス', width: 180 },
  { name: '電話番号', width: 120 },

  // SNS情報
  { name: 'Instagram', width: 180 },
  { name: 'X（Twitter）', width: 180 },
  { name: 'TikTok', width: 180 },
  { name: 'フォロワー数', width: 100 },

  // 希望・備考
  { name: '希望日程', width: 150 },
  { name: '一言メッセージ', width: 250 },
  { name: '備考（管理用）', width: 200 }
];

// ============================================
// メイン関数
// ============================================

/**
 * 応募管理シートを初期設定する（初回実行用）
 */
function setupApplicationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(APPLICATION_CONFIG.SHEET_NAME);

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(APPLICATION_CONFIG.SHEET_NAME);
  }

  // ヘッダー設定
  setupApplicationHeaders(sheet);

  // 列幅設定
  setupApplicationColumnWidths(sheet);

  // データ入力規則（ドロップダウン）設定
  setupApplicationDataValidation(sheet);

  // 条件付き書式設定
  setupApplicationConditionalFormatting(sheet);

  // 表示形式設定
  setupApplicationNumberFormats(sheet);

  SpreadsheetApp.getUi().alert('応募管理シートのセットアップが完了しました！');
}

/**
 * ヘッダーを設定
 */
function setupApplicationHeaders(sheet) {
  const headers = APPLICATION_COLUMNS.map(col => col.name);
  const headerRange = sheet.getRange(APPLICATION_CONFIG.HEADER_ROW, 1, 1, headers.length);

  headerRange.setValues([headers]);
  headerRange.setBackground(APPLICATION_CONFIG.HEADER_COLOR);
  headerRange.setFontColor(APPLICATION_CONFIG.HEADER_FONT_COLOR);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // ヘッダー行を固定
  sheet.setFrozenRows(APPLICATION_CONFIG.HEADER_ROW);
}

/**
 * 列幅を設定
 */
function setupApplicationColumnWidths(sheet) {
  APPLICATION_COLUMNS.forEach((col, index) => {
    sheet.setColumnWidth(index + 1, col.width);
  });
}

/**
 * データ入力規則（ドロップダウン）を設定
 */
function setupApplicationDataValidation(sheet) {
  const lastRow = 1000;

  // 応募ステータス（C列 = 3）
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(APPLICATION_CONFIG.STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(APPLICATION_CONFIG.DATA_START_ROW, 3, lastRow, 1).setDataValidation(statusRule);
}

/**
 * 条件付き書式（ステータス別色分け）を設定
 */
function setupApplicationConditionalFormatting(sheet) {
  const lastRow = 1000;
  const lastCol = APPLICATION_COLUMNS.length;
  const range = sheet.getRange(APPLICATION_CONFIG.DATA_START_ROW, 1, lastRow, lastCol);

  // 既存の条件付き書式をクリア
  sheet.setConditionalFormatRules([]);

  const newRules = [];

  // ステータス別の色設定
  Object.entries(APPLICATION_CONFIG.STATUS_COLORS).forEach(([status, color]) => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$C2="${status}"`)
      .setBackground(color)
      .setRanges([range])
      .build();
    newRules.push(rule);
  });

  sheet.setConditionalFormatRules(newRules);
}

/**
 * 表示形式を設定
 */
function setupApplicationNumberFormats(sheet) {
  const lastRow = 1000;

  // タイムスタンプ（B列 = 2）
  sheet.getRange(APPLICATION_CONFIG.DATA_START_ROW, 2, lastRow, 1)
    .setNumberFormat('yyyy/mm/dd hh:mm:ss');
}

// ============================================
// 応募一覧→応募管理 自動転記
// ============================================

/**
 * 応募フォーム送信時に応募一覧から応募管理シートへ転記
 * ※トリガー設定: 応募フォームのスプレッドシートで「フォーム送信時」トリガーを設定
 */
function onApplicationFormSubmit(e) {
  const sourceSheet = e.range.getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(APPLICATION_CONFIG.SHEET_NAME);

  // 応募管理シートが存在しない場合は終了
  if (!targetSheet) return;

  // フォームの回答データを取得（12列）
  const row = e.range.getRow();
  const sourceData = sourceSheet.getRange(row, 1, 1, 12).getValues()[0];

  // 応募一覧の列順:
  // [0] タイムスタンプ
  // [1] 案件ID
  // [2] 店舗名（フォームから）
  // [3] お名前（活動名可）
  // [4] メールアドレス
  // [5] 電話番号（任意）
  // [6] InstagramアカウントURL
  // [7] X（Twitter）アカウントURL（任意）
  // [8] TikTokアカウントURL（任意）
  // [9] 合計フォロワー数
  // [10] 希望日程
  // [11] 一言メッセージ（任意）

  const timestamp = sourceData[0];
  const projectId = sourceData[1];
  // 店舗名は案件管理から取得、なければフォームの値を使用
  const shopName = getShopNameFromProjectId(ss, projectId) || sourceData[2];

  // 応募管理シートに追加するデータを作成
  const newId = generateApplicationId(targetSheet);
  const newRow = [
    newId,           // 応募ID（自動採番）
    timestamp,       // タイムスタンプ
    '新規',          // 応募ステータス（初期値）
    projectId,       // 案件ID
    shopName,        // 店舗名
    sourceData[3],   // インフルエンサー名（お名前）
    sourceData[4],   // メールアドレス
    sourceData[5],   // 電話番号
    sourceData[6],   // Instagram
    sourceData[7],   // X（Twitter）
    sourceData[8],   // TikTok
    sourceData[9],   // フォロワー数
    sourceData[10],  // 希望日程
    sourceData[11],  // 一言メッセージ
    ''               // 備考（管理用）
  ];

  // 応募管理シートの最終行に追加
  targetSheet.appendRow(newRow);

  // 追加した行番号を取得
  const addedRow = targetSheet.getLastRow();

  // メール通知を送信
  sendNotificationEmail(targetSheet, addedRow);

  // 案件管理シートのステータスを「相談中」に更新
  updateProjectStatusById(ss, projectId);
}

/**
 * 案件IDから店舗名を取得
 */
function getShopNameFromProjectId(ss, projectId) {
  if (!projectId) return '';

  const projectSheet = ss.getSheetByName(APPLICATION_CONFIG.PROJECT_SHEET_NAME);
  if (!projectSheet) return '';

  const data = projectSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) { // A列 = 案件ID
      return data[i][3]; // D列 = 店舗名
    }
  }

  return '';
}

/**
 * 案件IDで案件管理シートのステータスを更新
 */
function updateProjectStatusById(ss, projectId) {
  if (!projectId) return;

  const projectSheet = ss.getSheetByName(APPLICATION_CONFIG.PROJECT_SHEET_NAME);
  if (!projectSheet) return;

  const data = projectSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) { // A列 = 案件ID
      const currentStatus = data[i][2]; // C列 = ステータス

      // 「公開中」の場合のみ「相談中」に変更
      if (currentStatus === '公開中') {
        projectSheet.getRange(i + 1, 3).setValue('相談中');
      }
      break;
    }
  }
}

/**
 * 応募IDを生成（APP-YYYYMM-001形式）
 */
function generateApplicationId(sheet) {
  const now = new Date();
  const yearMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');
  const prefix = `APP-${yearMonth}-`;

  const data = sheet.getDataRange().getValues();
  let maxNum = 0;

  data.forEach(row => {
    const id = row[0];
    if (id && id.toString().startsWith(prefix)) {
      const num = parseInt(id.toString().split('-')[2], 10);
      if (num > maxNum) maxNum = num;
    }
  });

  const newNum = String(maxNum + 1).padStart(3, '0');
  return `${prefix}${newNum}`;
}

/**
 * 新規応募の通知メールを送信
 */
function sendNotificationEmail(sheet, row) {
  const data = sheet.getRange(row, 1, 1, APPLICATION_COLUMNS.length).getValues()[0];

  const applicationId = data[0];
  const projectId = data[3];
  const shopName = data[4];
  const influencerName = data[5];
  const email = data[6];
  const instagram = data[8];
  const message = data[13];

  const subject = `【新規応募】${shopName} への応募がありました（${applicationId}）`;

  const body = `
新規応募がありました。

━━━━━━━━━━━━━━━━━━━━━━━━
■ 応募情報
━━━━━━━━━━━━━━━━━━━━━━━━
応募ID: ${applicationId}
案件ID: ${projectId}
店舗名: ${shopName}

■ インフルエンサー情報
━━━━━━━━━━━━━━━━━━━━━━━━
名前: ${influencerName}
メール: ${email}
Instagram: ${instagram}

■ 一言メッセージ
━━━━━━━━━━━━━━━━━━━━━━━━
${message || 'なし'}

━━━━━━━━━━━━━━━━━━━━━━━━
スプレッドシートで詳細を確認してください。
${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
  `;

  try {
    MailApp.sendEmail({
      to: APPLICATION_CONFIG.NOTIFICATION_EMAIL,
      subject: subject,
      body: body
    });
  } catch (error) {
    console.error('メール送信エラー:', error);
  }
}


// ============================================
// カスタムメニュー
// ============================================

/**
 * スプレッドシート起動時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('応募管理')
    .addItem('シートを初期設定', 'setupApplicationSheet')
    .addSeparator()
    .addItem('新規応募数を確認', 'countNewApplications')
    .addItem('ステータス別集計', 'showApplicationStatusSummary')
    .addToUi();
}

/**
 * 新規応募数をカウント
 */
function countNewApplications() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APPLICATION_CONFIG.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('応募管理シートが見つかりません。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let count = 0;

  data.forEach((row, index) => {
    if (index === 0) return;
    if (row[2] === '新規') count++;
  });

  SpreadsheetApp.getUi().alert(`新規応募: ${count}件`);
}

/**
 * ステータス別集計を表示
 */
function showApplicationStatusSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APPLICATION_CONFIG.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('応募管理シートが見つかりません。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const summary = {};

  APPLICATION_CONFIG.STATUS_OPTIONS.forEach(status => {
    summary[status] = 0;
  });

  data.forEach((row, index) => {
    if (index === 0) return;
    const status = row[2];
    if (summary.hasOwnProperty(status)) {
      summary[status]++;
    }
  });

  let message = '【応募ステータス別集計】\n\n';
  Object.entries(summary).forEach(([status, count]) => {
    message += `${status}: ${count}件\n`;
  });

  SpreadsheetApp.getUi().alert(message);
}
