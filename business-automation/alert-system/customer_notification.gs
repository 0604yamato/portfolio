/**
 * 顧客状況をChatworkに通知するスクリプト
 * 「次回やること」の日付を超えた顧客を通知します
 */

// ===== 設定項目 =====
const CONFIG = {
  // ChatworkのAPIトークン（https://www.chatwork.com/service/packages/chatwork/subpackages/api/token.php で取得）
  CHATWORK_API_TOKEN: 'YOUR_CHATWORK_API_TOKEN',

  // 通知先のChatworkルームID
  CHATWORK_ROOM_ID: 'YOUR_ROOM_ID',

  // スプレッドシートの設定
  SHEET_NAME: 'シート1', // シート名
  HEADER_ROW: 1, // ヘッダー行（通常は1）
  DATA_START_ROW: 2, // データ開始行

  // 列の設定（A列=1, B列=2...）
  COLUMN_CUSTOMER_NAME: 1, // 顧客名の列番号
  COLUMN_NEXT_ACTION_DATE: 2, // 次回やることの日付列番号
  COLUMN_NEXT_ACTION: 3, // 次回やることの内容列番号
  COLUMN_STATUS: 4, // ステータス列番号（任意）
  COLUMN_PERSON_IN_CHARGE: 5 // 担当者列番号（任意）
};

/**
 * メイン関数：日付超過の顧客をチェックして通知
 */
function checkAndNotifyOverdueCustomers() {
  try {
    const overdueCustomers = getOverdueCustomers();

    if (overdueCustomers.length === 0) {
      Logger.log('日付超過の顧客はいません');
      return;
    }

    const message = createChatworkMessage(overdueCustomers);
    sendToChatwork(message);

    Logger.log(`${overdueCustomers.length}件の顧客情報を通知しました`);

  } catch (error) {
    Logger.log('エラーが発生しました: ' + error.message);
    // エラー通知を送る場合
    sendToChatwork('[ERROR] 顧客通知スクリプトでエラーが発生しました: ' + error.message);
  }
}

/**
 * 日付超過の顧客データを取得
 */
function getOverdueCustomers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.DATA_START_ROW) {
    return [];
  }

  const dataRange = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, sheet.getLastColumn());
  const data = dataRange.getValues();

  const today = new Date();
  today.setHours(0, 0, 0, 0); // 時刻をリセット

  const overdueCustomers = [];

  data.forEach((row, index) => {
    const customerName = row[CONFIG.COLUMN_CUSTOMER_NAME - 1];
    const nextActionDate = row[CONFIG.COLUMN_NEXT_ACTION_DATE - 1];

    // 顧客名が空の行はスキップ
    if (!customerName) return;

    // 日付が入力されていない場合はスキップ
    if (!nextActionDate) return;

    // 日付型に変換
    const actionDate = new Date(nextActionDate);
    actionDate.setHours(0, 0, 0, 0);

    // 日付を超過している場合
    if (actionDate < today) {
      overdueCustomers.push({
        rowNumber: CONFIG.DATA_START_ROW + index,
        customerName: customerName,
        nextActionDate: Utilities.formatDate(actionDate, 'Asia/Tokyo', 'yyyy/MM/dd'),
        nextAction: row[CONFIG.COLUMN_NEXT_ACTION - 1] || '',
        status: row[CONFIG.COLUMN_STATUS - 1] || '',
        personInCharge: row[CONFIG.COLUMN_PERSON_IN_CHARGE - 1] || '',
        daysOverdue: Math.floor((today - actionDate) / (1000 * 60 * 60 * 24))
      });
    }
  });

  // 超過日数でソート（古い順）
  overdueCustomers.sort((a, b) => b.daysOverdue - a.daysOverdue);

  return overdueCustomers;
}

/**
 * Chatwork用のメッセージを作成
 */
function createChatworkMessage(customers) {
  let message = '[info][title]【要対応】次回やることの日付を超過している顧客[/title]';
  message += `対応が必要な顧客が ${customers.length} 件あります\n\n`;

  customers.forEach((customer, index) => {
    message += `${index + 1}. ${customer.customerName}\n`;
    message += `   予定日: ${customer.nextActionDate} (${customer.daysOverdue}日超過)\n`;

    if (customer.nextAction) {
      message += `   次回やること: ${customer.nextAction}\n`;
    }

    if (customer.status) {
      message += `   ステータス: ${customer.status}\n`;
    }

    if (customer.personInCharge) {
      message += `   担当者: ${customer.personInCharge}\n`;
    }

    message += `   行番号: ${customer.rowNumber}\n\n`;
  });

  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  message += `スプレッドシート: ${spreadsheetUrl}[/info]`;

  return message;
}

/**
 * Chatworkにメッセージを送信
 */
function sendToChatwork(message) {
  const url = `https://api.chatwork.com/v2/rooms/${CONFIG.CHATWORK_ROOM_ID}/messages`;

  const options = {
    method: 'post',
    headers: {
      'X-ChatWorkToken': CONFIG.CHATWORK_API_TOKEN
    },
    payload: {
      body: message
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error(`Chatwork API Error: ${responseCode} - ${response.getContentText()}`);
  }

  return JSON.parse(response.getContentText());
}

/**
 * 定期実行トリガーを設定（初回のみ実行）
 */
function setupDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkAndNotifyOverdueCustomers') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 毎日午前9時に実行するトリガーを設定
  ScriptApp.newTrigger('checkAndNotifyOverdueCustomers')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  Logger.log('毎日午前9時に実行するトリガーを設定しました');
}

/**
 * テスト実行用：現在の超過顧客を確認
 */
function testGetOverdueCustomers() {
  const customers = getOverdueCustomers();
  Logger.log(`日付超過の顧客: ${customers.length}件`);
  customers.forEach(customer => {
    Logger.log(customer);
  });
}
