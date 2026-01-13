/**
 * 見積書保存スクリプト
 * 見積書の内容を履歴に保存し、いつでも復元して編集できます
 * ※シートは増えません
 */

/**
 * 見積書を保存
 */
function saveQuoteToHistory() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const quoteSheet = ss.getSheetByName(CONFIG.QUOTE_SHEET_NAME);

  if (!quoteSheet) {
    ui.alert('エラー', `「${CONFIG.QUOTE_SHEET_NAME}」シートが見つかりません。`, ui.ButtonSet.OK);
    return;
  }

  // 保存名を入力
  const response = ui.prompt(
    '📝 見積書を保存',
    '保存名を入力してください（例：〇〇商事_厨房機器）',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const saveName = response.getResponseText().trim();
  if (!saveName) {
    ui.alert('エラー', '保存名を入力してください。', ui.ButtonSet.OK);
    return;
  }

  // 見積書シートの内容を取得
  const dataRange = quoteSheet.getDataRange();
  const values = dataRange.getValues();
  const formulas = dataRange.getFormulas();  // 数式を取得
  const formats = dataRange.getNumberFormats();

  // 結合セル情報を取得
  const mergedRanges = quoteSheet.getRange(1, 1, dataRange.getNumRows(), dataRange.getNumColumns())
    .getMergedRanges()
    .map(range => range.getA1Notation());

  // データ入力規則を取得
  const validations = [];
  const numRows = dataRange.getNumRows();
  const numCols = dataRange.getNumColumns();
  for (let row = 1; row <= numRows; row++) {
    for (let col = 1; col <= numCols; col++) {
      const cell = quoteSheet.getRange(row, col);
      const validation = cell.getDataValidation();
      if (validation) {
        validations.push({
          row: row,
          col: col,
          criteriaType: validation.getCriteriaType().toString(),
          criteriaValues: validation.getCriteriaValues(),
          helpText: validation.getHelpText()
        });
      }
    }
  }

  // データをJSON形式で保存
  const saveData = {
    values: values,
    formulas: formulas,  // 数式も保存
    formats: formats,
    mergedRanges: mergedRanges,  // 結合セル情報も保存
    validations: validations  // データ入力規則も保存
  };
  const jsonData = JSON.stringify(saveData);

  // 履歴シートに記録
  const historySheet = getOrCreateHistorySheet(ss);
  const now = new Date();
  const lastRow = historySheet.getLastRow();

  // 番号を計算
  const newNo = lastRow;  // ヘッダー行を除いた番号

  historySheet.getRange(lastRow + 1, 1, 1, 5).setValues([[
    newNo,
    now,
    saveName,
    jsonData,
    ''  // メモ欄
  ]]);

  // 日付フォーマット
  historySheet.getRange(lastRow + 1, 2).setNumberFormat('yyyy/mm/dd HH:mm');

  // データ列を非表示に
  historySheet.hideColumns(4);

  ui.alert('保存完了', `「${saveName}」を保存しました。`, ui.ButtonSet.OK);
}

/**
 * 履歴シートを取得または作成
 */
function getOrCreateHistorySheet(ss) {
  const historySheetName = '見積書履歴';
  let historySheet = ss.getSheetByName(historySheetName);

  if (!historySheet) {
    historySheet = ss.insertSheet(historySheetName);

    // ヘッダーを設定
    historySheet.getRange('A1:E1').setValues([['No', '保存日時', '見積書名', 'データ', 'メモ']]);
    historySheet.getRange('A1:E1').setBackground('#4a86e8').setFontColor('#ffffff').setFontWeight('bold');

    // 列幅を調整
    historySheet.setColumnWidth(1, 50);
    historySheet.setColumnWidth(2, 150);
    historySheet.setColumnWidth(3, 250);
    historySheet.setColumnWidth(4, 50);
    historySheet.setColumnWidth(5, 300);

    // データ列を非表示
    historySheet.hideColumns(4);

    // 行を固定
    historySheet.setFrozenRows(1);
  }

  return historySheet;
}

/**
 * 履歴から見積書を開く
 */
function showHistorySheet() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = getOrCreateHistorySheet(ss);

  const lastRow = historySheet.getLastRow();
  if (lastRow <= 1) {
    ui.alert('情報', '保存された見積書がありません。', ui.ButtonSet.OK);
    return;
  }

  // 一覧を取得（No, 保存日時, 見積書名）
  const data = historySheet.getRange(2, 1, lastRow - 1, 3).getValues();

  // 最新10件を取得（配列の後ろから10件）
  const latestData = data.slice(-10).reverse();

  const listText = latestData.map((row) => {
    const no = row[0];
    const date = Utilities.formatDate(new Date(row[1]), 'Asia/Tokyo', 'MM/dd HH:mm');
    const name = row[2];
    return `${no}. ${name}（${date}）`;
  }).join('\n');

  const totalCount = data.length;
  const showingText = totalCount > 10 ? `（最新10件を表示 / 全${totalCount}件）` : `（全${totalCount}件）`;

  const response = ui.prompt(
    '📂 履歴から見積書を復元',
    `${showingText}\n\n${listText}\n\n開く番号を入力：`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const inputNo = parseInt(response.getResponseText().trim());

  // 入力された番号に一致する行を検索
  const targetIndex = data.findIndex(row => row[0] === inputNo);

  if (isNaN(inputNo) || targetIndex === -1) {
    ui.alert('エラー', '正しい番号を入力してください。', ui.ButtonSet.OK);
    return;
  }

  // 選択した見積書を復元
  restoreQuote(ss, historySheet, targetIndex + 2, data[targetIndex][2]);
}

/**
 * 見積書を復元
 */
function restoreQuote(ss, historySheet, rowNum, saveName) {
  const ui = SpreadsheetApp.getUi();
  const quoteSheet = ss.getSheetByName(CONFIG.QUOTE_SHEET_NAME);

  if (!quoteSheet) {
    ui.alert('エラー', `「${CONFIG.QUOTE_SHEET_NAME}」シートが見つかりません。`, ui.ButtonSet.OK);
    return;
  }

  // JSONデータを取得（4列目）
  const jsonData = historySheet.getRange(rowNum, 4).getValue();

  try {
    const saveData = JSON.parse(jsonData);

    // 現在の見積書をクリア（データ入力規則も含む）
    const dataRange = quoteSheet.getDataRange();
    dataRange.clearContent();
    dataRange.clearDataValidations();  // データ入力規則をクリア

    // データを復元
    if (saveData.values && saveData.values.length > 0) {
      const numRows = saveData.values.length;
      const numCols = saveData.values[0].length;

      // 値と数式を結合した配列を作成
      const combinedData = [];
      for (let row = 0; row < numRows; row++) {
        const rowData = [];
        for (let col = 0; col < numCols; col++) {
          const formula = saveData.formulas ? saveData.formulas[row][col] : '';
          if (formula && formula !== '') {
            // 数式がある場合は数式を使用
            rowData.push(formula);
          } else {
            // 数式がない場合は値を使用
            rowData.push(saveData.values[row][col]);
          }
        }
        combinedData.push(rowData);
      }

      // 一括で設定（数式も値も同時に）
      const range = quoteSheet.getRange(1, 1, numRows, numCols);
      range.setValues(combinedData);

      // フォーマットを設定
      if (saveData.formats) {
        range.setNumberFormats(saveData.formats);
      }

      // 結合セルを復元
      if (saveData.mergedRanges && saveData.mergedRanges.length > 0) {
        for (const rangeNotation of saveData.mergedRanges) {
          quoteSheet.getRange(rangeNotation).merge();
        }
      }

      // データ入力規則を復元
      if (saveData.validations && saveData.validations.length > 0) {
        for (const v of saveData.validations) {
          try {
            const cell = quoteSheet.getRange(v.row, v.col);
            let rule = null;

            if (v.criteriaType === 'VALUE_IN_LIST') {
              rule = SpreadsheetApp.newDataValidation()
                .requireValueInList(v.criteriaValues[0], true)
                .setHelpText(v.helpText || '')
                .build();
            } else if (v.criteriaType === 'VALUE_IN_RANGE') {
              rule = SpreadsheetApp.newDataValidation()
                .requireValueInRange(quoteSheet.getParent().getRange(v.criteriaValues[0].getA1Notation()), true)
                .setHelpText(v.helpText || '')
                .build();
            }

            if (rule) {
              cell.setDataValidation(rule);
            }
          } catch (e) {
            // 個別のエラーは無視して続行
            console.log('Validation restore error: ' + e.message);
          }
        }
      }
    }

    // 見積書シートを表示
    ss.setActiveSheet(quoteSheet);

    ui.alert('復元完了', `「${saveName}」を開きました。\n\n編集後、保存してください。`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('エラー', `復元に失敗しました: ${e.message}`, ui.ButtonSet.OK);
  }
}
