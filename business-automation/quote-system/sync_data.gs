/**
 * データ同期スクリプト
 * 管理者のデータシートを担当者のデータシートに反映します
 */

// 担当者管理シート名
const STAFF_SHEET_NAME = '担当者管理';

/**
 * メニューを追加（スプレッドシート起動時に実行）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 見積書システム')
    .addItem('🔄 全担当者にデータを同期', 'syncDataToAllStaff')
    .addSeparator()
    .addItem('➕ 担当者を追加', 'addStaffMember')
    .addItem('➖ 担当者を削除', 'removeStaffMember')
    .addItem('👥 担当者一覧を表示', 'showStaffList')
    .addSeparator()
    .addItem('📝 見積書を履歴に保存', 'saveQuoteToHistory')
    .addItem('📂 履歴から見積書を復元', 'showHistorySheet')
    .addToUi();
}

/**
 * 担当者管理シートを取得（なければ作成）
 */
function getOrCreateStaffSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let staffSheet = ss.getSheetByName(STAFF_SHEET_NAME);

  if (!staffSheet) {
    staffSheet = ss.insertSheet(STAFF_SHEET_NAME);

    // ヘッダーを設定
    staffSheet.getRange('A1:B1').setValues([['担当者名', 'スプレッドシートID']]);
    staffSheet.getRange('A1:B1').setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');

    // 列幅を調整
    staffSheet.setColumnWidth(1, 150);
    staffSheet.setColumnWidth(2, 400);

    // 行を固定
    staffSheet.setFrozenRows(1);
  }

  return staffSheet;
}

/**
 * 担当者を追加
 */
function addStaffMember() {
  const ui = SpreadsheetApp.getUi();

  // 担当者名を入力
  const nameResponse = ui.prompt(
    '➕ 担当者を追加',
    '担当者名を入力してください（例：田中太郎）',
    ui.ButtonSet.OK_CANCEL
  );

  if (nameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const staffName = nameResponse.getResponseText().trim();
  if (!staffName) {
    ui.alert('エラー', '担当者名を入力してください。', ui.ButtonSet.OK);
    return;
  }

  // スプレッドシートIDを入力
  const idResponse = ui.prompt(
    '➕ 担当者を追加',
    `${staffName}さんのスプレッドシートIDを入力してください。\n\n` +
    '※ URLの https://docs.google.com/spreadsheets/d/【ここ】/edit の部分',
    ui.ButtonSet.OK_CANCEL
  );

  if (idResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const sheetId = idResponse.getResponseText().trim();
  if (!sheetId) {
    ui.alert('エラー', 'スプレッドシートIDを入力してください。', ui.ButtonSet.OK);
    return;
  }

  // 担当者管理シートに追加
  const staffSheet = getOrCreateStaffSheet();
  const lastRow = staffSheet.getLastRow();
  staffSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[staffName, sheetId]]);

  ui.alert('完了', `${staffName}さんを担当者に追加しました。`, ui.ButtonSet.OK);
}

/**
 * 担当者を削除
 */
function removeStaffMember() {
  const ui = SpreadsheetApp.getUi();
  const staffSheet = getOrCreateStaffSheet();

  // 担当者リストを取得
  const staffList = getStaffList();

  if (staffList.length === 0) {
    ui.alert('情報', '登録されている担当者がいません。', ui.ButtonSet.OK);
    return;
  }

  // 担当者名一覧を表示
  const staffNames = staffList.map((s, i) => `${i + 1}. ${s.name}`).join('\n');

  const response = ui.prompt(
    '➖ 担当者を削除',
    `削除する担当者の番号を入力してください。\n\n${staffNames}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const index = parseInt(response.getResponseText().trim()) - 1;

  if (isNaN(index) || index < 0 || index >= staffList.length) {
    ui.alert('エラー', '正しい番号を入力してください。', ui.ButtonSet.OK);
    return;
  }

  const staffToRemove = staffList[index];

  // 確認
  const confirm = ui.alert(
    '確認',
    `${staffToRemove.name}さんを削除しますか？`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }

  // 行を削除（ヘッダー行があるので +2）
  staffSheet.deleteRow(index + 2);

  ui.alert('完了', `${staffToRemove.name}さんを削除しました。`, ui.ButtonSet.OK);
}

/**
 * 担当者一覧を表示
 */
function showStaffList() {
  const ui = SpreadsheetApp.getUi();
  const staffList = getStaffList();

  if (staffList.length === 0) {
    ui.alert('担当者一覧', '登録されている担当者がいません。\n\nメニューから「➕ 担当者を追加」で追加してください。', ui.ButtonSet.OK);
    return;
  }

  const listText = staffList.map((s, i) => `${i + 1}. ${s.name}`).join('\n');

  ui.alert('担当者一覧', `登録済みの担当者：\n\n${listText}\n\n合計: ${staffList.length}名`, ui.ButtonSet.OK);
}

/**
 * 担当者リストを取得
 */
function getStaffList() {
  const staffSheet = getOrCreateStaffSheet();
  const lastRow = staffSheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  const data = staffSheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return data
    .filter(row => row[0] && row[1])
    .map(row => ({
      name: row[0],
      sheetId: row[1]
    }));
}

/**
 * 全担当者にデータを同期
 */
function syncDataToAllStaff() {
  const staffList = getStaffList();
  const ui = SpreadsheetApp.getUi();

  if (staffList.length === 0) {
    ui.alert('情報', '登録されている担当者がいません。\n\nメニューから「➕ 担当者を追加」で追加してください。', ui.ButtonSet.OK);
    return;
  }

  const results = [];

  for (const staff of staffList) {
    try {
      syncDataToStaff(staff.sheetId, staff.name);
      results.push(`✅ ${staff.name}: 同期成功`);
    } catch (e) {
      results.push(`❌ ${staff.name}: ${e.message}`);
    }
  }

  ui.alert('同期結果', results.join('\n'), ui.ButtonSet.OK);
}

/**
 * 指定した担当者のスプレッドシートにデータを同期
 * @param {string} targetSheetId - 同期先のスプレッドシートID
 * @param {string} staffName - 担当者名（ログ用）
 */
function syncDataToStaff(targetSheetId, staffName) {
  // 管理者のデータシートを取得
  const adminSS = SpreadsheetApp.getActiveSpreadsheet();
  const adminDataSheet = adminSS.getSheetByName(CONFIG.DATA_SHEET_NAME);

  if (!adminDataSheet) {
    throw new Error(`管理者のスプレッドシートに「${CONFIG.DATA_SHEET_NAME}」シートが見つかりません`);
  }

  // 担当者のスプレッドシートを取得
  const targetSS = SpreadsheetApp.openById(targetSheetId);
  const targetDataSheet = targetSS.getSheetByName(CONFIG.DATA_SHEET_NAME);

  if (!targetDataSheet) {
    throw new Error(`担当者のスプレッドシートに「${CONFIG.DATA_SHEET_NAME}」シートが見つかりません`);
  }

  // 管理者のデータを取得
  const adminData = adminDataSheet.getDataRange().getValues();
  const adminFormats = adminDataSheet.getDataRange().getNumberFormats();

  if (adminData.length === 0) {
    throw new Error('管理者のデータシートにデータがありません');
  }

  // 担当者のデータシートをクリアして上書き
  targetDataSheet.clearContents();

  // データを書き込み
  const range = targetDataSheet.getRange(1, 1, adminData.length, adminData[0].length);
  range.setValues(adminData);
  range.setNumberFormats(adminFormats);

  // ログに記録
  console.log(`${staffName}へのデータ同期完了: ${adminData.length}行`);
}

/**
 * 管理者のデータシート編集時に自動同期（トリガー用）
 */
function onEditTrigger(e) {
  const sheet = e.source.getActiveSheet();

  if (sheet.getName() !== CONFIG.DATA_SHEET_NAME) {
    return;
  }

  syncDataToAllStaff();
}

/**
 * 編集時自動同期トリガーを設定
 */
function setupAutoSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditTrigger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onEditTrigger')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('完了', '自動同期トリガーを設定しました。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 自動同期トリガーを解除
 */
function removeAutoSyncTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = false;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditTrigger') {
      ScriptApp.deleteTrigger(trigger);
      removed = true;
    }
  });

  if (removed) {
    SpreadsheetApp.getUi().alert('完了', '自動同期トリガーを解除しました。', SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('情報', '自動同期トリガーは設定されていません。', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
