/**
 * 顧客マスターシートの設定スクリプト
 * ステータス選択式、ナンバー自動入力、項目の色付け
 */

// ===== 設定項目 =====
const MASTER_CONFIG = {
  SHEET_NAME: '顧客マスター', // シート名
  HEADER_ROW: 1, // ヘッダー行
  DATA_START_ROW: 2, // データ開始行
  TOTAL_COLUMNS: 31, // 総列数

  // 列の設定（31項目）
  COLUMN_MEMBER_NO: 1,        // 会員No（A列）
  COLUMN_DEPARTMENT: 2,       // 所属（B列）
  COLUMN_EMPLOYEE_NO: 3,      // 社員No（C列）
  COLUMN_EMPLOYEE_NAME: 4,    // 社員氏名（D列）
  COLUMN_CUSTOMER_NAME: 5,    // 顧客名（E列）
  COLUMN_LIST_TYPE: 6,        // リスト区分（F列）
  COLUMN_INDUSTRY: 7,         // 業種（G列）
  COLUMN_CONTACT: 8,          // 担当者接触（H列）
  COLUMN_REVIEW_PERIOD: 9,    // 検討時期（I列）
  COLUMN_TENANT: 10,          // テナント扱い有無（J列）
  COLUMN_SUCCESSION: 11,      // 事業承継ニーズ（K列）
  COLUMN_FC: 12,              // FCニーズ（L列）
  COLUMN_ESTIMATE_SENT: 13,   // 見積送付（M列）
  COLUMN_APPOINTMENT: 14,     // アポ見込み（N列）
  COLUMN_CLOSED_WON: 15,      // 成約獲得（O列）
  COLUMN_CALL_NG: 16,         // 架電NG（P列）
  COLUMN_CALL_EXCLUDED: 17,   // 架電対象外（Q列）
  COLUMN_NG_REASON: 18,       // NG理由（R列）
  COLUMN_GUARANTEE: 19,       // 保証会社（S列）
  COLUMN_LIST_COMPLETE: 20,   // リスト完了（T列）
  COLUMN_STATUS: 21,          // 最終ステータス（U列）
  COLUMN_TEL1: 22,            // Tel1（V列）
  COLUMN_TEL2: 23,            // Tel2（W列）
  COLUMN_TEL3: 24,            // Tel3 部署番号①（X列）
  COLUMN_ADDRESS: 25,         // 店舗住所（Y列）
  COLUMN_POSTAL_CODE: 26,     // 郵便番号（Z列）
  COLUMN_ADDRESS1: 27,        // 第１住所（AA列）
  COLUMN_ADDRESS2: 28,        // 第２住所（AB列）
  COLUMN_ADDRESS3: 29,        // 第３住所（AC列）
  COLUMN_EMAIL: 30,           // メールアドレス（AD列）
  COLUMN_LAST_CALL_DATE: 31,  // 最終架電日（AE列）

  // 〇✖選択肢（担当者接触〜架電対象外）
  CIRCLE_CROSS_OPTIONS: ['〇', '✖'],

  // リスト区分の選択肢
  LIST_TYPE_OPTIONS: ['事業承継', '家賃保証'],

  // リスト区分の色設定
  LIST_TYPE_COLORS: {
    '事業承継': '#fce5cd',  // 淡いオレンジ
    '家賃保証': '#d9ead3'   // 淡い緑
  },

  // 〇✖の色設定
  CIRCLE_CROSS_COLORS: {
    '〇': '#d9ead3',  // 淡い緑
    '✖': '#f4cccc'   // 淡い赤
  },

  // ステータスの選択肢
  STATUS_OPTIONS: [
    '接点無し',
    '検討中',
    '見積もり提出済',
    '回答待ち',
    '契約準備中',
    '成約',
    '失注',
    '保留'
  ],

  // ステータスと関係値の対応
  STATUS_RELATION_SCORES: {
    '接点無し': 1,
    '検討中': 3,
    '見積もり提出済': 5,
    '回答待ち': 6,
    '契約準備中': 8,
    '成約': 10,
    '失注': 0,
    '保留': 2
  },

  // 色設定
  HEADER_COLOR_RED: '#f4cccc', // ヘッダーの背景色（淡い赤）：会員No〜最終ステータス
  HEADER_COLOR_BLUE: '#cfe2f3', // ヘッダーの背景色（淡い青）：Tel1〜最終架電日
  HEADER_FONT_COLOR: '#333333', // ヘッダーの文字色（濃いグレー）

  // ステータスごとの色設定
  STATUS_COLORS: {
    '接点無し': '#f3f3f3',         // 薄いグレー
    '検討中': '#fff2cc',          // 黄色
    '見積もり提出済': '#fce5cd',  // オレンジ
    '回答待ち': '#cfe2f3',         // 薄い青
    '契約準備中': '#d9ead3',       // 薄い緑
    '成約': '#b6d7a8',             // 緑
    '失注': '#cccccc',             // グレー
    '保留': '#ead1dc'              // 薄いピンク
  }
};

/**
 * 顧客マスターシートの初期設定
 */
function setupCustomerMasterSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(MASTER_CONFIG.SHEET_NAME);

    // シートが存在しない場合は新規作成
    if (!sheet) {
      sheet = createCustomerMasterSheet(spreadsheet);
      Logger.log('顧客マスターシートを新規作成しました');
    }

    // 1. ヘッダー行に色を付ける
    formatHeaderRow(sheet);

    // 2. ステータス列にドロップダウンリストを設定
    setupStatusDropdown(sheet);

    // 3. リスト区分列にドロップダウンリストを設定
    setupListTypeDropdown(sheet);

    // 4. 〇✖選択列にドロップダウンリストを設定
    setupCircleCrossDropdown(sheet);

    // 5. 会員No列に連番を設定
    setupMemberNoColumn(sheet);

    // 6. 列幅を設定
    setColumnWidths(sheet);

    // 7. 条件付き書式を設定（色付け）
    setupConditionalFormatting(sheet);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      '顧客マスターシートの設定が完了しました',
      '設定完了',
      5
    );

    Logger.log('顧客マスターシートの設定が完了しました');

  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'エラー: ' + error.message,
      'エラー',
      10
    );
  }
}

/**
 * 顧客マスターシートを新規作成
 */
function createCustomerMasterSheet(spreadsheet) {
  // 新しいシートを作成
  const sheet = spreadsheet.insertSheet(MASTER_CONFIG.SHEET_NAME);

  // ヘッダー行を設定（31項目）
  const headers = [
    '会員No',
    '所属',
    '社員No',
    '社員氏名',
    '顧客名',
    'リスト区分',
    '業種',
    '担当者接触',
    '検討時期',
    'テナント扱い有無',
    '事業承継ニーズ',
    'FCニーズ',
    '見積送付',
    'アポ見込み',
    '成約獲得',
    '架電NG',
    '架電対象外',
    'NG理由',
    '保証会社',
    'リスト完了',
    '最終ステータス',
    'Tel1',
    'Tel2',
    'Tel3 部署番号①',
    '店舗住所',
    '郵便番号',
    '第１住所',
    '第２住所',
    '第３住所',
    'メールアドレス',
    '最終架電日'
  ];

  // ヘッダーを設定
  const headerRange = sheet.getRange(MASTER_CONFIG.HEADER_ROW, 1, 1, headers.length);
  headerRange.setValues([headers]);

  // シートを一番左に移動
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.moveActiveSheet(1);

  return sheet;
}

/**
 * ヘッダー行のフォーマット設定
 */
function formatHeaderRow(sheet) {
  // 会員No〜最終ステータス（1〜21列）：淡い赤
  const redRange = sheet.getRange(MASTER_CONFIG.HEADER_ROW, 1, 1, MASTER_CONFIG.COLUMN_STATUS);
  redRange.setBackground(MASTER_CONFIG.HEADER_COLOR_RED);

  // Tel1〜最終架電日（22〜32列）：淡い青
  const blueRange = sheet.getRange(MASTER_CONFIG.HEADER_ROW, MASTER_CONFIG.COLUMN_TEL1, 1, MASTER_CONFIG.TOTAL_COLUMNS - MASTER_CONFIG.COLUMN_TEL1 + 1);
  blueRange.setBackground(MASTER_CONFIG.HEADER_COLOR_BLUE);

  // 全体の文字設定
  const headerRange = sheet.getRange(MASTER_CONFIG.HEADER_ROW, 1, 1, MASTER_CONFIG.TOTAL_COLUMNS);
  headerRange.setFontColor(MASTER_CONFIG.HEADER_FONT_COLOR);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');

  // 高さを調整
  sheet.setRowHeight(MASTER_CONFIG.HEADER_ROW, 30);

  Logger.log('ヘッダー行のフォーマット完了');
}

/**
 * ステータス列にドロップダウンリストを設定
 */
function setupStatusDropdown(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 100); // 最低100行分設定

  // ステータス列全体に入力規則を設定
  const statusRange = sheet.getRange(
    MASTER_CONFIG.DATA_START_ROW,
    MASTER_CONFIG.COLUMN_STATUS,
    lastRow - MASTER_CONFIG.DATA_START_ROW + 1,
    1
  );

  // ドロップダウンリストの作成
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MASTER_CONFIG.STATUS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();

  statusRange.setDataValidation(rule);

  Logger.log('ステータスのドロップダウンリスト設定完了');
}

/**
 * リスト区分列にドロップダウンリストを設定
 */
function setupListTypeDropdown(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 100); // 最低100行分設定

  // リスト区分列に入力規則を設定
  const listTypeRange = sheet.getRange(
    MASTER_CONFIG.DATA_START_ROW,
    MASTER_CONFIG.COLUMN_LIST_TYPE,
    lastRow - MASTER_CONFIG.DATA_START_ROW + 1,
    1
  );

  // ドロップダウンリストの作成
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MASTER_CONFIG.LIST_TYPE_OPTIONS, true)
    .setAllowInvalid(false)
    .build();

  listTypeRange.setDataValidation(rule);

  Logger.log('リスト区分のドロップダウンリスト設定完了');
}

/**
 * 〇✖選択列にドロップダウンリストを設定（担当者接触〜架電対象外）
 */
function setupCircleCrossDropdown(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 100); // 最低100行分設定

  // 〇✖ドロップダウンリストの作成
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(MASTER_CONFIG.CIRCLE_CROSS_OPTIONS, true)
    .setAllowInvalid(false)
    .build();

  // 担当者接触から架電対象外までの列（8〜17列目）に設定
  const columns = [
    MASTER_CONFIG.COLUMN_CONTACT,        // 担当者接触
    MASTER_CONFIG.COLUMN_REVIEW_PERIOD,  // 検討時期
    MASTER_CONFIG.COLUMN_TENANT,         // テナント扱い有無
    MASTER_CONFIG.COLUMN_SUCCESSION,     // 事業承継ニーズ
    MASTER_CONFIG.COLUMN_FC,             // FCニーズ
    MASTER_CONFIG.COLUMN_ESTIMATE_SENT,  // 見積送付
    MASTER_CONFIG.COLUMN_APPOINTMENT,    // アポ見込み
    MASTER_CONFIG.COLUMN_CLOSED_WON,     // 成約獲得
    MASTER_CONFIG.COLUMN_CALL_NG,        // 架電NG
    MASTER_CONFIG.COLUMN_CALL_EXCLUDED   // 架電対象外
  ];

  columns.forEach(col => {
    const range = sheet.getRange(
      MASTER_CONFIG.DATA_START_ROW,
      col,
      lastRow - MASTER_CONFIG.DATA_START_ROW + 1,
      1
    );
    range.setDataValidation(rule);
  });

  Logger.log('〇✖選択のドロップダウンリスト設定完了');
}

/**
 * 列幅を設定
 */
function setColumnWidths(sheet) {
  // 基本情報
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_MEMBER_NO, 80);      // 会員No
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_DEPARTMENT, 100);    // 所属
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_EMPLOYEE_NO, 80);    // 社員No
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_EMPLOYEE_NAME, 100); // 社員氏名
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_CUSTOMER_NAME, 150); // 顧客名
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_LIST_TYPE, 100);     // リスト区分
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_INDUSTRY, 100);      // 業種

  // 分類・ニーズ情報
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_CONTACT, 100);       // 担当者接触
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_REVIEW_PERIOD, 100); // 検討時期
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_TENANT, 100);        // テナント扱い有無
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_SUCCESSION, 100);    // 事業承継ニーズ
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_FC, 80);             // FCニーズ

  // ステータス管理
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_ESTIMATE_SENT, 80);  // 見積送付
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_APPOINTMENT, 80);    // アポ見込み
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_CLOSED_WON, 80);     // 成約獲得
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_CALL_NG, 80);        // 架電NG
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_CALL_EXCLUDED, 80);  // 架電対象外
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_NG_REASON, 150);     // NG理由
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_GUARANTEE, 100);     // 保証会社
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_LIST_COMPLETE, 80);  // リスト完了
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_STATUS, 120);        // 最終ステータス

  // 連絡先情報
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_TEL1, 120);          // Tel1
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_TEL2, 120);          // Tel2
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_TEL3, 120);          // Tel3 部署番号①

  // 住所情報
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_ADDRESS, 200);       // 店舗住所
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_POSTAL_CODE, 100);   // 郵便番号
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_ADDRESS1, 120);      // 第１住所
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_ADDRESS2, 120);      // 第２住所
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_ADDRESS3, 150);      // 第３住所
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_EMAIL, 200);         // メールアドレス
  sheet.setColumnWidth(MASTER_CONFIG.COLUMN_LAST_CALL_DATE, 100);// 最終架電日

  Logger.log('列幅の設定完了');
}

/**
 * 条件付き書式を設定（色付け）
 */
function setupConditionalFormatting(sheet) {
  const rules = [];
  const maxRow = 1000; // 十分な行数を確保

  // 〇✖選択列のリスト
  const circleCrossColumns = [
    MASTER_CONFIG.COLUMN_CONTACT,
    MASTER_CONFIG.COLUMN_REVIEW_PERIOD,
    MASTER_CONFIG.COLUMN_TENANT,
    MASTER_CONFIG.COLUMN_SUCCESSION,
    MASTER_CONFIG.COLUMN_FC,
    MASTER_CONFIG.COLUMN_ESTIMATE_SENT,
    MASTER_CONFIG.COLUMN_APPOINTMENT,
    MASTER_CONFIG.COLUMN_CLOSED_WON,
    MASTER_CONFIG.COLUMN_CALL_NG,
    MASTER_CONFIG.COLUMN_CALL_EXCLUDED
  ];

  // === リスト区分の条件付き書式 ===
  const listTypeCol = MASTER_CONFIG.COLUMN_LIST_TYPE;
  const listTypeRange = sheet.getRange(MASTER_CONFIG.DATA_START_ROW, listTypeCol, maxRow, 1);

  // 事業承継
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('事業承継')
    .setBackground(MASTER_CONFIG.LIST_TYPE_COLORS['事業承継'])
    .setRanges([listTypeRange])
    .build());

  // 家賃保証
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('家賃保証')
    .setBackground(MASTER_CONFIG.LIST_TYPE_COLORS['家賃保証'])
    .setRanges([listTypeRange])
    .build());

  // === 〇✖選択列の条件付き書式 ===
  circleCrossColumns.forEach(col => {
    const range = sheet.getRange(MASTER_CONFIG.DATA_START_ROW, col, maxRow, 1);

    // 〇
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('〇')
      .setBackground(MASTER_CONFIG.CIRCLE_CROSS_COLORS['〇'])
      .setRanges([range])
      .build());

    // ✖
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('✖')
      .setBackground(MASTER_CONFIG.CIRCLE_CROSS_COLORS['✖'])
      .setRanges([range])
      .build());
  });

  // === ステータス列の条件付き書式（行全体に色） ===
  const statusCol = MASTER_CONFIG.COLUMN_STATUS;
  const fullRowRange = sheet.getRange(MASTER_CONFIG.DATA_START_ROW, 1, maxRow, MASTER_CONFIG.TOTAL_COLUMNS);

  Object.keys(MASTER_CONFIG.STATUS_COLORS).forEach(status => {
    const color = MASTER_CONFIG.STATUS_COLORS[status];
    const formula = `=$U2="${status}"`;

    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(formula)
      .setBackground(color)
      .setRanges([fullRowRange])
      .build());
  });

  // 条件付き書式を適用
  sheet.setConditionalFormatRules(rules);

  Logger.log('条件付き書式の設定完了');
}

/**
 * 会員No列に連番を設定
 */
function setupMemberNoColumn(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < MASTER_CONFIG.DATA_START_ROW) {
    return; // データがない場合は何もしない
  }

  // データ行の範囲を取得
  const dataRows = lastRow - MASTER_CONFIG.DATA_START_ROW + 1;
  const memberNoRange = sheet.getRange(
    MASTER_CONFIG.DATA_START_ROW,
    MASTER_CONFIG.COLUMN_MEMBER_NO,
    dataRows,
    1
  );

  // 連番を生成
  const numbers = [];
  for (let i = 1; i <= dataRows; i++) {
    numbers.push([i]);
  }

  // 連番を設定
  memberNoRange.setValues(numbers);

  // 中央揃え
  memberNoRange.setHorizontalAlignment('center');

  Logger.log(`会員No列に1〜${dataRows}の連番を設定しました`);
}

/**
 * 新しい行を追加した時に自動で会員Noを振る
 * ※色付けは条件付き書式で自動処理されるため不要
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();

  // 顧客マスターシート以外は無視
  if (sheet.getName() !== MASTER_CONFIG.SHEET_NAME) {
    return;
  }

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // ヘッダー行は無視
  if (row <= MASTER_CONFIG.HEADER_ROW) {
    return;
  }

  // 会員No列以外が編集された場合、会員Noを自動入力
  if (col !== MASTER_CONFIG.COLUMN_MEMBER_NO && col >= MASTER_CONFIG.COLUMN_DEPARTMENT && col <= MASTER_CONFIG.COLUMN_LAST_CALL_DATE) {
    const memberNoCell = sheet.getRange(row, MASTER_CONFIG.COLUMN_MEMBER_NO);

    // 会員Noが空の場合のみ設定
    if (!memberNoCell.getValue()) {
      // 会員No列のデータを一括取得して最大値を求める
      const lastRow = sheet.getLastRow();
      if (lastRow >= MASTER_CONFIG.DATA_START_ROW) {
        const memberNoRange = sheet.getRange(MASTER_CONFIG.DATA_START_ROW, MASTER_CONFIG.COLUMN_MEMBER_NO, lastRow - MASTER_CONFIG.DATA_START_ROW + 1, 1);
        const values = memberNoRange.getValues().flat().filter(v => typeof v === 'number');
        const maxNumber = values.length > 0 ? Math.max(...values) : 0;

        // 次の会員Noを設定
        memberNoCell.setValue(maxNumber + 1);
        memberNoCell.setHorizontalAlignment('center');
      } else {
        memberNoCell.setValue(1);
        memberNoCell.setHorizontalAlignment('center');
      }
    }
  }
}

/**
 * カスタムメニューを追加（顧客マスター設定 + アラート通知）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 顧客マスター設定メニュー
  ui.createMenu('顧客マスター設定')
    .addItem('初期設定を実行', 'setupCustomerMasterSheet')
    .addSeparator()
    .addItem('会員Noを再採番', 'renumberMemberNo')
    .addItem('条件付き書式を再設定', 'reapplyConditionalFormatting')
    .addToUi();

  // 顧客管理メニュー（アラート通知）
  ui.createMenu('顧客管理')
    .addItem('要対応リストを更新', 'updateOverdueNotifications')
    .addSeparator()
    .addItem('自動更新を設定（毎日9時）', 'setupDailyTrigger')
    .addToUi();
}

/**
 * 会員Noを再採番する（メニューから実行用）
 */
function renumberMemberNo() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(MASTER_CONFIG.SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `シート「${MASTER_CONFIG.SHEET_NAME}」が見つかりません`,
      'エラー',
      5
    );
    return;
  }

  setupMemberNoColumn(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '会員Noの再採番が完了しました',
    '完了',
    3
  );
}

/**
 * 条件付き書式を再設定する（メニューから実行用）
 */
function reapplyConditionalFormatting() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(MASTER_CONFIG.SHEET_NAME);

  if (!sheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `シート「${MASTER_CONFIG.SHEET_NAME}」が見つかりません`,
      'エラー',
      5
    );
    return;
  }

  setupConditionalFormatting(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    '条件付き書式の設定が完了しました',
    '完了',
    3
  );
}
