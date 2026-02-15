/**
 * スプレッドシート操作（拡張版）
 * T〜W列（同意書同意、同意日時、企業URL、当日受付フラグ）を追加
 */

/**
 * スプレッドシートにデータを保存（拡張版）
 * @param {Object} data - フォームデータ
 * @param {boolean} isWalkIn - 当日受付かどうか
 * @returns {number} 保存した行番号
 */
function saveToSpreadsheet(data, isWalkIn) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  const initialStatus = isWalkIn ? STATUS.CONFIRMED : STATUS.PENDING;

  const newRow = [
    data.timestamp,           // A: タイムスタンプ
    data.id,                  // B: 申込ID
    data.name,                // C: お名前
    data.company,             // D: 貴社名
    data.email,               // E: メールアドレス
    data.phone,               // F: 電話番号
    data.position,            // G: 役職
    data.industry,            // H: 業種
    data.theme,               // I: 相談テーマ
    data.content,             // J: 相談内容
    data.date1 + (data.time ? ' ' + data.time : ''),  // K: 希望日時1（時間帯付き）
    data.date2,               // L: 希望日時2
    data.method,              // M: 相談方法
    initialStatus,            // N: ステータス
    '',                       // O: 担当者
    '',                       // P: 確定日時
    '',                       // Q: ZoomURL
    '',                       // R: ヒアリングシート
    data.remarks || data.notes || '',  // S: 備考
    '',                       // T: 同意書同意
    '',                       // U: 同意日時
    data.companyUrl || '',    // V: 企業URL
    isWalkIn ? 'TRUE' : 'FALSE',  // W: 当日受付フラグ
    ''                        // X: 場所
  ];

  sheet.appendRow(newRow);

  return sheet.getLastRow();
}

/**
 * 行データを取得（拡張版）
 */
function getRowData(rowIndex) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  const row = sheet.getRange(rowIndex, 1, 1, 24).getValues()[0];

  return {
    timestamp: row[COLUMNS.TIMESTAMP],
    id: row[COLUMNS.ID],
    name: row[COLUMNS.NAME],
    company: row[COLUMNS.COMPANY],
    email: row[COLUMNS.EMAIL],
    phone: row[COLUMNS.PHONE],
    position: row[COLUMNS.POSITION],
    industry: row[COLUMNS.INDUSTRY],
    theme: row[COLUMNS.THEME],
    content: row[COLUMNS.CONTENT],
    date1: row[COLUMNS.DATE1],
    date2: row[COLUMNS.DATE2],
    method: row[COLUMNS.METHOD],
    status: row[COLUMNS.STATUS],
    staff: row[COLUMNS.STAFF],
    confirmedDate: row[COLUMNS.CONFIRMED_DATE],
    zoomUrl: row[COLUMNS.ZOOM_URL],
    hearingSheet: row[COLUMNS.HEARING_SHEET],
    notes: row[COLUMNS.NOTES],
    ndaStatus: row[COLUMNS.NDA_STATUS],
    ndaDate: row[COLUMNS.NDA_DATE],
    companyUrl: row[COLUMNS.COMPANY_URL],
    walkInFlag: row[COLUMNS.WALK_IN_FLAG],
    location: row[COLUMNS.LOCATION]
  };
}

/**
 * ステータスを更新
 */
function updateStatus(rowIndex, newStatus) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(newStatus);
}

/**
 * スプレッドシートのヘッダーを設定（拡張版）
 */
function setupSpreadsheetHeaders() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  const headers = [
    'タイムスタンプ',
    '申込ID',
    'お名前',
    '貴社名',
    'メールアドレス',
    '電話番号',
    '役職',
    '業種',
    '相談テーマ',
    '相談内容',
    '希望日時1',
    '希望日時2',
    '相談方法',
    'ステータス',
    '担当者',
    '確定日時',
    'ZoomURL',
    'ヒアリングシート',
    '備考',
    '同意書同意',
    '同意日時',
    '企業URL',
    '当日受付フラグ',
    '場所'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅の調整
  sheet.setColumnWidth(1, 150);  // タイムスタンプ
  sheet.setColumnWidth(2, 120);  // 申込ID
  sheet.setColumnWidth(3, 100);  // お名前
  sheet.setColumnWidth(4, 150);  // 貴社名
  sheet.setColumnWidth(5, 200);  // メールアドレス
  sheet.setColumnWidth(6, 120);  // 電話番号
  sheet.setColumnWidth(7, 100);  // 役職
  sheet.setColumnWidth(8, 100);  // 業種
  sheet.setColumnWidth(9, 120);  // 相談テーマ
  sheet.setColumnWidth(10, 250); // 相談内容
  sheet.setColumnWidth(11, 150); // 希望日時1
  sheet.setColumnWidth(12, 150); // 希望日時2
  sheet.setColumnWidth(13, 100); // 相談方法
  sheet.setColumnWidth(14, 80);  // ステータス
  sheet.setColumnWidth(15, 100); // 担当者
  sheet.setColumnWidth(16, 150); // 確定日時
  sheet.setColumnWidth(17, 250); // ZoomURL
  sheet.setColumnWidth(18, 100); // ヒアリングシート
  sheet.setColumnWidth(19, 200); // 備考
  sheet.setColumnWidth(20, 80);  // 同意書同意
  sheet.setColumnWidth(21, 150); // 同意日時
  sheet.setColumnWidth(22, 250); // 企業URL
  sheet.setColumnWidth(23, 100); // 当日受付フラグ
  sheet.setColumnWidth(24, 120); // 場所

  // 1行目を固定
  sheet.setFrozenRows(1);

  // ステータス列にプルダウン設定
  const statusRange = sheet.getRange(2, COLUMNS.STATUS + 1, 1000, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      STATUS.PENDING,
      STATUS.NDA_AGREED,
      STATUS.RECEIVED,
      STATUS.CONFIRMED,
      STATUS.COMPLETED,
      STATUS.CANCELLED
    ])
    .build();
  statusRange.setDataValidation(statusRule);

  // 場所列にプルダウン設定
  const locationRange = sheet.getRange(2, COLUMNS.LOCATION + 1, 1000, 1);
  const locationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LOCATION_OPTIONS)
    .build();
  locationRange.setDataValidation(locationRule);

  // 同意書同意列にプルダウン設定
  const ndaRange = sheet.getRange(2, COLUMNS.NDA_STATUS + 1, 1000, 1);
  const ndaRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['済', '未'])
    .build();
  ndaRange.setDataValidation(ndaRule);

  console.log('スプレッドシート（拡張版）のセットアップが完了しました');
}

/**
 * ステータスに応じた条件付き書式を設定（拡張版）
 */
function setupConditionalFormatting() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  const range = sheet.getRange('N2:N1000');

  sheet.clearConditionalFormatRules();

  const rules = [];

  // 仮予約: 黄色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(STATUS.PENDING)
    .setBackground('#fff3cd')
    .setRanges([range])
    .build());

  // NDA同意済: オレンジ
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(STATUS.NDA_AGREED)
    .setBackground('#ffe0b2')
    .setRanges([range])
    .build());

  // 書類受領: 青
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(STATUS.RECEIVED)
    .setBackground('#cce5ff')
    .setRanges([range])
    .build());

  // 確定: 緑
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(STATUS.CONFIRMED)
    .setBackground('#d4edda')
    .setRanges([range])
    .build());

  // 完了: グレー
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(STATUS.COMPLETED)
    .setBackground('#e2e3e5')
    .setRanges([range])
    .build());

  // キャンセル: 赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(STATUS.CANCELLED)
    .setBackground('#f8d7da')
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(rules);

  console.log('条件付き書式の設定が完了しました');
}
