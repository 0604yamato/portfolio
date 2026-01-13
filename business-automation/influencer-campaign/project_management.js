/**
 * インフルエンサー施策 - 案件管理スプレッドシート設定
 *
 * 【セットアップ手順】
 * 1. 新しいスプレッドシートを作成
 * 2. 拡張機能 → Apps Script を開く
 * 3. このコードを貼り付けて保存
 * 4. setupProjectSheet() を実行（初回のみ）
 */

// ============================================
// 設定
// ============================================
const PROJECT_CONFIG = {
  // シート名
  SHEET_NAME: '案件管理',

  // ヘッダー行
  HEADER_ROW: 1,

  // データ開始行
  DATA_START_ROW: 2,

  // ステータス選択肢
  STATUS_OPTIONS: ['下書き', '審査中', '公開中', '相談中', '確定', '見送り', '終了'],

  // ジャンル選択肢
  GENRE_OPTIONS: [
    'ラーメン', '寿司', '焼肉', '居酒屋', 'カフェ',
    'イタリアン', 'フレンチ', '中華', '和食', '洋食',
    'カレー', '焼き鳥', 'うどん・そば', 'バー', 'その他'
  ],

  // エリア選択肢
  AREA_OPTIONS: [
    '北海道', '東北', '関東', '東京', '中部',
    '近畿', '大阪', '中国', '四国', '九州・沖縄'
  ],

  // 報酬タイプ選択肢
  REWARD_TYPE_OPTIONS: ['無料提供のみ', '無料提供＋謝礼', '謝礼のみ', '要相談'],

  // ヘッダー色
  HEADER_COLOR: '#4a86e8',
  HEADER_FONT_COLOR: '#ffffff',

  // ステータス別色
  STATUS_COLORS: {
    '下書き': '#f3f3f3',
    '審査中': '#fff2cc',
    '公開中': '#d9ead3',
    '相談中': '#cfe2f3',
    '確定': '#d0e0e3',
    '見送り': '#f4cccc',
    '終了': '#d9d9d9'
  }
};

// ============================================
// カラム定義（21列）
// ============================================
const PROJECT_COLUMNS = [
  // 基本情報
  { name: '案件ID', width: 80 },
  { name: 'タイムスタンプ', width: 150 },
  { name: 'ステータス', width: 100 },

  // 店舗情報
  { name: '店舗名', width: 150 },
  { name: '担当者 氏名', width: 100 },
  { name: '担当者 電話番号', width: 120 },
  { name: '担当者 メールアドレス', width: 180 },

  // 所在地
  { name: '店舗 郵便番号', width: 90 },
  { name: '店舗 都道府県', width: 80 },
  { name: '店舗 市区町村', width: 120 },
  { name: '店舗 番地・建物名', width: 150 },
  { name: '店舗 エリア', width: 100 },
  { name: '店舗 ジャンル', width: 100 },

  // 案件詳細
  { name: '案件タイトル', width: 200 },
  { name: 'PR内容', width: 300 },
  { name: '来店可能曜日', width: 120 },
  { name: '来店可能時間', width: 120 },

  // 報酬
  { name: '報酬金額', width: 150 },

  // その他
  { name: 'NG事項', width: 200 },
  { name: '画像URL', width: 200 },
  { name: '備考', width: 200 }
];

// ============================================
// メイン関数
// ============================================

/**
 * スプレッドシートを初期設定する（初回実行用）
 */
function setupProjectSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PROJECT_CONFIG.SHEET_NAME);

  // シートがなければ作成
  if (!sheet) {
    sheet = ss.insertSheet(PROJECT_CONFIG.SHEET_NAME);
  }

  // ヘッダー設定
  setupHeaders(sheet);

  // 列幅設定
  setupColumnWidths(sheet);

  // データ入力規則（ドロップダウン）設定
  setupDataValidation(sheet);

  // 条件付き書式設定
  setupConditionalFormatting(sheet);

  // 表示形式設定
  setupNumberFormats(sheet);

  SpreadsheetApp.getUi().alert('案件管理シートのセットアップが完了しました！');
}

/**
 * ヘッダーを設定
 */
function setupHeaders(sheet) {
  const headers = PROJECT_COLUMNS.map(col => col.name);
  const headerRange = sheet.getRange(PROJECT_CONFIG.HEADER_ROW, 1, 1, headers.length);

  headerRange.setValues([headers]);
  headerRange.setBackground(PROJECT_CONFIG.HEADER_COLOR);
  headerRange.setFontColor(PROJECT_CONFIG.HEADER_FONT_COLOR);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');

  // ヘッダー行を固定
  sheet.setFrozenRows(PROJECT_CONFIG.HEADER_ROW);
}

/**
 * 列幅を設定
 */
function setupColumnWidths(sheet) {
  PROJECT_COLUMNS.forEach((col, index) => {
    sheet.setColumnWidth(index + 1, col.width);
  });
}

/**
 * データ入力規則（ドロップダウン）を設定
 */
function setupDataValidation(sheet) {
  const lastRow = 1000; // 十分な行数

  // ステータス（C列 = 3）
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PROJECT_CONFIG.STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(PROJECT_CONFIG.DATA_START_ROW, 3, lastRow, 1).setDataValidation(statusRule);

  // 店舗 エリア（L列 = 12）
  const areaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PROJECT_CONFIG.AREA_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(PROJECT_CONFIG.DATA_START_ROW, 12, lastRow, 1).setDataValidation(areaRule);

  // 店舗 ジャンル（M列 = 13）
  const genreRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(PROJECT_CONFIG.GENRE_OPTIONS, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(PROJECT_CONFIG.DATA_START_ROW, 13, lastRow, 1).setDataValidation(genreRule);
}

/**
 * 条件付き書式（ステータス別色分け）を設定
 */
function setupConditionalFormatting(sheet) {
  const lastRow = 1000;
  const lastCol = PROJECT_COLUMNS.length;
  const range = sheet.getRange(PROJECT_CONFIG.DATA_START_ROW, 1, lastRow, lastCol);

  // 既存の条件付き書式をクリア
  const rules = sheet.getConditionalFormatRules();
  sheet.setConditionalFormatRules([]);

  const newRules = [];

  // ステータス別の色設定
  Object.entries(PROJECT_CONFIG.STATUS_COLORS).forEach(([status, color]) => {
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
function setupNumberFormats(sheet) {
  const lastRow = 1000;

  // タイムスタンプ（B列 = 2）
  sheet.getRange(PROJECT_CONFIG.DATA_START_ROW, 2, lastRow, 1)
    .setNumberFormat('yyyy/mm/dd hh:mm:ss');
}

// ============================================
// 案件ID自動採番 & 自動転記
// ============================================

/**
 * フォーム送信時に店舗登録シートから案件管理シートへ転記
 */
function onFormSubmit(e) {
  const sourceSheet = e.range.getSheet();
  const sourceSheetName = sourceSheet.getName();

  // 応募一覧からの送信は無視（応募管理用のトリガーで処理）
  if (sourceSheetName === '応募一覧') return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(PROJECT_CONFIG.SHEET_NAME);

  // 案件管理シートが存在しない場合は終了
  if (!targetSheet) return;

  // フォームの回答データを取得（タイムスタンプ + 20項目）
  const row = e.range.getRow();
  const sourceData = sourceSheet.getRange(row, 1, 1, 21).getValues()[0];

  // タイムスタンプ
  const timestamp = sourceData[0];

  // フォームデータ（質問1〜20）
  const formData = sourceData.slice(1);

  // 案件管理シートに追加するデータを作成
  // [案件ID, タイムスタンプ, ステータス, 店舗名, 担当者名, 電話, メール, 〒, 都道府県, 市区町村, 番地, エリア, ジャンル, タイトル, PR内容, 開始日, 終了日, 曜日時間, 報酬タイプ, 報酬詳細, NG, 画像URL, 備考]
  const newId = generateProjectId(targetSheet);
  const newRow = [
    newId,           // 案件ID（自動採番）
    timestamp,       // タイムスタンプ
    '下書き',        // ステータス（初期値）
    ...formData      // フォームデータ20項目
  ];

  // 案件管理シートの最終行に追加
  targetSheet.appendRow(newRow);
}

/**
 * 案件IDを生成（PRJ-YYYYMM-001形式）
 */
function generateProjectId(sheet) {
  const now = new Date();
  const yearMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');
  const prefix = `PRJ-${yearMonth}-`;

  // 既存の最大番号を取得
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

// ============================================
// カスタムメニュー
// ============================================

/**
 * スプレッドシート起動時にメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('案件管理')
    .addItem('シートを初期設定', 'setupProjectSheet')
    .addSeparator()
    .addItem('公開中の案件数を確認', 'countPublishedProjects')
    .addItem('ステータス別集計', 'showStatusSummary')
    .addToUi();
}

/**
 * 公開中の案件数をカウント
 */
function countPublishedProjects() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_CONFIG.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('案件管理シートが見つかりません。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  let count = 0;

  data.forEach((row, index) => {
    if (index === 0) return; // ヘッダーをスキップ
    if (row[2] === '公開中') count++;
  });

  SpreadsheetApp.getUi().alert(`公開中の案件: ${count}件`);
}

/**
 * ステータス別集計を表示
 */
function showStatusSummary() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROJECT_CONFIG.SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('案件管理シートが見つかりません。');
    return;
  }

  const data = sheet.getDataRange().getValues();
  const summary = {};

  PROJECT_CONFIG.STATUS_OPTIONS.forEach(status => {
    summary[status] = 0;
  });

  data.forEach((row, index) => {
    if (index === 0) return; // ヘッダーをスキップ
    const status = row[2];
    if (summary.hasOwnProperty(status)) {
      summary[status]++;
    }
  });

  let message = '【ステータス別集計】\n\n';
  Object.entries(summary).forEach(([status, count]) => {
    message += `${status}: ${count}件\n`;
  });

  SpreadsheetApp.getUi().alert(message);
}

// ============================================
// Web API用関数（案件一覧UI用）
// ============================================

/**
 * 公開中の案件を取得（JSON形式）
 */
function getPublishedProjects() {
  const SPREADSHEET_ID = 'YOUR_PROJECT_SPREADSHEET_ID';
  const SHEET_NAME = '案件管理';

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    const values = sh.getDataRange().getValues();

    // ヘッダー取得
    const header = [];
    for (var i = 0; i < values[0].length; i++) {
      header.push(String(values[0][i]).trim());
    }
    const statusIdx = header.indexOf('ステータス');

    // 公開中の案件を取得
    const projects = [];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var status = String(row[statusIdx] || '').trim();
      if (status === '公開中') {
        var project = {};
        for (var c = 0; c < header.length; c++) {
          var val = row[c];
          // DateオブジェクトをISO文字列に変換（Webアプリで必要）
          if (val instanceof Date) {
            project[header[c]] = val.toISOString();
          } else {
            project[header[c]] = val;
          }
        }
        projects.push(project);
      }
    }

    return {
      meta: { count: projects.length },
      projects: projects
    };
  } catch (e) {
    return {
      meta: { error: e.message },
      projects: []
    };
  }
}

/**
 * Webアプリとしてデプロイ時のGETハンドラー
 * ?action=getProjects → JSON API（案件一覧取得）
 * ?page=register → 案件登録フォーム（店舗用）
 * それ以外 → 案件一覧（インフルエンサー用）
 */
function doGet(e) {
  var action = '';
  var page = '';

  if (e && e.parameter) {
    action = e.parameter.action || '';
    page = e.parameter.page || '';
  }

  // API リクエスト
  if (action === 'getProjects') {
    var result = getPublishedProjects();
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // HTMLページ
  if (page === 'register') {
    return HtmlService.createHtmlOutputFromFile('ProjectRegister')
      .setTitle('案件登録フォーム - クライアント企業 インフルエンサー施策')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (page === 'test') {
    return HtmlService.createHtmlOutputFromFile('TestPage')
      .setTitle('テストページ')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    return HtmlService.createHtmlOutputFromFile('案件一覧')
      .setTitle('クライアント企業 インフルエンサー案件一覧')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * POSTリクエストハンドラー（API用）
 */
function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;
    var data = requestData.data;

    var result;
    if (action === 'submitApplication') {
      result = submitApplication(data);
    } else if (action === 'submitProject') {
      result = submitProject(data);
    } else {
      result = { success: false, error: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// 応募データ保存（独自フォーム用）
// ============================================

/**
 * 応募データをスプレッドシートに保存
 */
function submitApplication(data) {
  const SPREADSHEET_ID = 'YOUR_PROJECT_SPREADSHEET_ID';
  const SHEET_NAME = '応募管理';

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);

    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // ヘッダーを設定
      const headers = [
        '応募ID', 'タイムスタンプ', '応募ステータス', '案件ID', '店舗名',
        'インフルエンサー名', 'メールアドレス', '電話番号',
        'Instagram', 'X（Twitter）', 'TikTok', 'フォロワー数',
        '希望日程', '一言メッセージ', '備考（管理用）'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // 応募IDを生成
    const applicationId = generateApplicationIdForSubmit(sheet);

    // タイムスタンプ
    const timestamp = new Date();

    // 新しい行のデータ
    const newRow = [
      applicationId,          // 応募ID
      timestamp,              // タイムスタンプ
      '新規',                 // 応募ステータス
      data.projectId || '',   // 案件ID
      data.shopName || '',    // 店舗名
      data.name || '',        // インフルエンサー名
      data.email || '',       // メールアドレス
      data.phone || '',       // 電話番号
      data.instagram || '',   // Instagram
      data.twitter || '',     // X（Twitter）
      data.tiktok || '',      // TikTok
      data.followers || '',   // フォロワー数
      data.schedule || '',    // 希望日程
      data.message || '',     // 一言メッセージ
      ''                      // 備考（管理用）
    ];

    // 最終行に追加
    sheet.appendRow(newRow);

    // 案件のステータスを「相談中」に更新
    updateProjectStatusToConsulting(ss, data.projectId);

    // メール通知（オプション）
    sendApplicationNotification(data, applicationId);

    return { success: true, applicationId: applicationId };

  } catch (e) {
    console.error('応募保存エラー:', e);
    return { success: false, error: e.message };
  }
}

/**
 * 応募IDを生成（APP-YYYYMM-001形式）
 */
function generateApplicationIdForSubmit(sheet) {
  const now = new Date();
  const yearMonth = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMM');
  const prefix = 'APP-' + yearMonth + '-';

  const data = sheet.getDataRange().getValues();
  let maxNum = 0;

  for (var i = 1; i < data.length; i++) {
    var id = data[i][0];
    if (id && id.toString().indexOf(prefix) === 0) {
      var num = parseInt(id.toString().split('-')[2], 10);
      if (num > maxNum) maxNum = num;
    }
  }

  var newNum = String(maxNum + 1);
  while (newNum.length < 3) newNum = '0' + newNum;
  return prefix + newNum;
}

/**
 * 案件のステータスを「相談中」に更新
 */
function updateProjectStatusToConsulting(ss, projectId) {
  if (!projectId) return;

  const sheet = ss.getSheetByName('案件管理');
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === projectId) {
      var currentStatus = data[i][2];
      if (currentStatus === '公開中') {
        sheet.getRange(i + 1, 3).setValue('相談中');
      }
      break;
    }
  }
}

/**
 * 応募通知メールを送信
 */
function sendApplicationNotification(data, applicationId) {
  // 通知先メールアドレス（必要に応じて変更）
  const NOTIFICATION_EMAIL = 'YOUR_EMAIL@example.com';

  // メールアドレスが設定されていなければスキップ
  if (NOTIFICATION_EMAIL === 'YOUR_EMAIL@example.com') return;

  try {
    const subject = '【新規応募】' + data.shopName + ' への応募（' + applicationId + '）';
    const body = [
      '新規応募がありました。',
      '',
      '━━━━━━━━━━━━━━━━━━━━━━━━',
      '■ 応募情報',
      '━━━━━━━━━━━━━━━━━━━━━━━━',
      '応募ID: ' + applicationId,
      '案件ID: ' + data.projectId,
      '店舗名: ' + data.shopName,
      '',
      '■ インフルエンサー情報',
      '━━━━━━━━━━━━━━━━━━━━━━━━',
      '名前: ' + data.name,
      'メール: ' + data.email,
      'Instagram: ' + data.instagram,
      '希望日程: ' + data.schedule,
      '',
      '━━━━━━━━━━━━━━━━━━━━━━━━',
      'スプレッドシートで詳細を確認してください。'
    ].join('\n');

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      body: body
    });
  } catch (e) {
    console.error('メール送信エラー:', e);
  }
}

// ============================================
// 案件登録（店舗用フォーム）
// ============================================

/**
 * 案件登録データをスプレッドシートに保存
 */
function submitProject(data) {
  const SPREADSHEET_ID = 'YOUR_PROJECT_SPREADSHEET_ID';
  const SHEET_NAME = '案件管理';

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);

    // シートがなければ作成（通常は setupProjectSheet で作成済み）
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      setupHeaders(sheet);
    }

    // 案件IDを生成
    var projectId = generateProjectId(sheet);

    // タイムスタンプ
    var timestamp = new Date();

    // 新しい行のデータ（カラム順序に従う）
    // [案件ID, タイムスタンプ, ステータス, 店舗名, 担当者氏名, 担当者電話, 担当者メール, 〒, 都道府県, 市区町村, 番地, エリア, ジャンル, タイトル, PR内容, 来店可能曜日, 来店可能時間, 報酬金額, NG, 画像URL, 備考]
    var newRow = [
      projectId,                    // 案件ID
      timestamp,                    // タイムスタンプ
      '下書き',                     // ステータス（初期値）
      data.shopName || '',          // 店舗名
      data.contactName || '',       // 担当者 氏名
      data.contactPhone || '',      // 担当者 電話番号
      data.contactEmail || '',      // 担当者 メールアドレス
      data.postalCode || '',        // 店舗 郵便番号
      data.prefecture || '',        // 店舗 都道府県
      data.city || '',              // 店舗 市区町村
      data.address || '',           // 店舗 番地・建物名
      data.area || '',              // 店舗 エリア
      data.genre || '',             // 店舗 ジャンル
      data.projectTitle || '',      // 案件タイトル
      data.prContent || '',         // PR内容
      data.availableDays || '',     // 来店可能曜日
      data.availableTime || '',     // 来店可能時間
      data.rewardAmount || '',      // 報酬金額
      data.ngItems || '',           // NG事項
      data.imageUrl || '',          // 画像URL
      data.remarks || ''            // 備考
    ];

    // 最終行に追加
    sheet.appendRow(newRow);

    // 登録通知メールを送信
    sendProjectNotification(data, projectId);

    return { success: true, projectId: projectId };

  } catch (e) {
    console.error('案件登録エラー:', e);
    return { success: false, error: e.message };
  }
}

/**
 * 案件登録通知メールを送信
 */
function sendProjectNotification(data, projectId) {
  // 通知先メールアドレス（必要に応じて変更）
  const NOTIFICATION_EMAIL = 'YOUR_EMAIL@example.com';

  // メールアドレスが設定されていなければスキップ
  if (NOTIFICATION_EMAIL === 'YOUR_EMAIL@example.com') return;

  try {
    var subject = '【新規案件登録】' + data.shopName + '（' + projectId + '）';
    var body = [
      '新規案件が登録されました。',
      '',
      '━━━━━━━━━━━━━━━━━━━━━━━━',
      '■ 案件情報',
      '━━━━━━━━━━━━━━━━━━━━━━━━',
      '案件ID: ' + projectId,
      '店舗名: ' + data.shopName,
      '担当者: ' + data.contactName,
      'タイトル: ' + data.projectTitle,
      'ジャンル: ' + data.genre,
      'エリア: ' + data.area,
      '報酬金額: ' + data.rewardAmount,
      '',
      '━━━━━━━━━━━━━━━━━━━━━━━━',
      '審査後、ステータスを「公開中」に変更してください。'
    ].join('\n');

    MailApp.sendEmail({
      to: NOTIFICATION_EMAIL,
      subject: subject,
      body: body
    });
  } catch (e) {
    console.error('メール送信エラー:', e);
  }
}

