/**
 * スプレッドシート操作（拡張版）
 * 列順: A〜M(従来), N:場所, O:ステータス, P:担当者, Q:確定日時, R〜X
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
    '',                       // N: 場所
    initialStatus,            // O: ステータス
    '',                       // P: 担当者
    '',                       // Q: 確定日時
    '',                       // R: ZoomURL
    '',                       // S: ヒアリングシート
    data.remarks || data.notes || '',  // T: 備考
    '',                       // U: 同意書同意
    '',                       // V: 同意日時
    data.companyUrl || '',    // W: 企業URL
    isWalkIn ? 'TRUE' : 'FALSE',  // X: 当日受付フラグ
    '',                       // Y: リーダー
    ''                        // Z: レポート状態
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

  const row = sheet.getRange(rowIndex, 1, 1, 26).getValues()[0];

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
    location: row[COLUMNS.LOCATION],
    status: row[COLUMNS.STATUS],
    staff: row[COLUMNS.STAFF],
    confirmedDate: row[COLUMNS.CONFIRMED_DATE] instanceof Date
      ? Utilities.formatDate(row[COLUMNS.CONFIRMED_DATE], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
      : row[COLUMNS.CONFIRMED_DATE],
    zoomUrl: row[COLUMNS.ZOOM_URL],
    hearingSheet: row[COLUMNS.HEARING_SHEET],
    notes: row[COLUMNS.NOTES],
    ndaStatus: row[COLUMNS.NDA_STATUS],
    ndaDate: row[COLUMNS.NDA_DATE],
    companyUrl: row[COLUMNS.COMPANY_URL],
    walkInFlag: row[COLUMNS.WALK_IN_FLAG],
    leader: row[COLUMNS.LEADER],
    reportStatus: row[COLUMNS.REPORT_STATUS]
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
    'タイムスタンプ',    // A
    '申込ID',           // B
    'お名前',           // C
    '貴社名',           // D
    'メールアドレス',    // E
    '電話番号',         // F
    '役職',             // G
    '業種',             // H
    '相談テーマ',       // I
    '相談内容',         // J
    '希望日時1',        // K
    '希望日時2',        // L
    '相談方法',         // M
    '場所',             // N
    'ステータス',       // O
    '担当者',           // P
    '確定日時',         // Q
    'ZoomURL',          // R
    'ヒアリングシート',  // S
    '備考',             // T
    '同意書同意',       // U
    '同意日時',         // V
    '企業URL',          // W
    '当日受付フラグ',    // X
    'リーダー',         // Y
    'レポート状態'       // Z
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅の調整
  sheet.setColumnWidth(1, 150);  // A: タイムスタンプ
  sheet.setColumnWidth(2, 120);  // B: 申込ID
  sheet.setColumnWidth(3, 100);  // C: お名前
  sheet.setColumnWidth(4, 150);  // D: 貴社名
  sheet.setColumnWidth(5, 200);  // E: メールアドレス
  sheet.setColumnWidth(6, 120);  // F: 電話番号
  sheet.setColumnWidth(7, 100);  // G: 役職
  sheet.setColumnWidth(8, 100);  // H: 業種
  sheet.setColumnWidth(9, 120);  // I: 相談テーマ
  sheet.setColumnWidth(10, 250); // J: 相談内容
  sheet.setColumnWidth(11, 180); // K: 希望日時1
  sheet.setColumnWidth(12, 150); // L: 希望日時2
  sheet.setColumnWidth(13, 100); // M: 相談方法
  sheet.setColumnWidth(14, 120); // N: 場所
  sheet.setColumnWidth(15, 100); // O: ステータス
  sheet.setColumnWidth(16, 100); // P: 担当者
  sheet.setColumnWidth(17, 180); // Q: 確定日時
  sheet.setColumnWidth(18, 250); // R: ZoomURL
  sheet.setColumnWidth(19, 100); // S: ヒアリングシート
  sheet.setColumnWidth(20, 200); // T: 備考
  sheet.setColumnWidth(21, 80);  // U: 同意書同意
  sheet.setColumnWidth(22, 150); // V: 同意日時
  sheet.setColumnWidth(23, 250); // W: 企業URL
  sheet.setColumnWidth(24, 100); // X: 当日受付フラグ
  sheet.setColumnWidth(25, 100); // Y: リーダー
  sheet.setColumnWidth(26, 100); // Z: レポート状態

  // 1行目を固定
  sheet.setFrozenRows(1);

  // 場所列にプルダウン設定
  const locationRange = sheet.getRange(2, COLUMNS.LOCATION + 1, 1000, 1);
  const locationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LOCATION_OPTIONS)
    .build();
  locationRange.setDataValidation(locationRule);

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

  // ステータス列（動的に列番号を取得）
  const range = sheet.getRange(2, COLUMNS.STATUS + 1, 999, 1);

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

/**
 * 既存スプレッドシートに場所列を挿入するマイグレーション
 * M列（相談方法）の直後にN列（場所）を挿入し、既存データを自動シフト
 */
function migrateAddLocationColumn() {
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  // 既にN列が「場所」ならスキップ
  var headerN = sheet.getRange(1, 14).getValue();
  if (headerN === '場所') {
    // プルダウンだけ再設定
    var locRange = sheet.getRange(2, 14, 1000, 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(LOCATION_OPTIONS)
      .build();
    locRange.setDataValidation(rule);
    console.log('場所列は既に存在します。プルダウンを再設定しました。');
    return { success: true, message: '場所列は既に存在します。プルダウンを再設定しました。' };
  }

  // M列(13, 1-based)の後に列を挿入 → 既存N〜W列が自動的に右にシフト
  sheet.insertColumnAfter(13);

  // ヘッダー設定
  sheet.getRange(1, 14).setValue('場所');
  sheet.getRange(1, 14)
    .setBackground('#0f2350')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setColumnWidth(14, 120);

  // プルダウン設定
  var locationRange = sheet.getRange(2, 14, 1000, 1);
  var locationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LOCATION_OPTIONS)
    .build();
  locationRange.setDataValidation(locationRule);

  // ステータス列（旧N→O）のプルダウンも再設定
  var statusRange = sheet.getRange(2, COLUMNS.STATUS + 1, 1000, 1);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      STATUS.PENDING, STATUS.NDA_AGREED, STATUS.RECEIVED,
      STATUS.CONFIRMED, STATUS.COMPLETED, STATUS.CANCELLED
    ])
    .build();
  statusRange.setDataValidation(statusRule);

  // 条件付き書式を再設定
  setupConditionalFormatting();

  console.log('場所列をN列に追加しました。旧N〜W列はO〜X列にシフトしました。');
  return { success: true, message: '場所列をN列（相談方法の次）に追加しました。' };
}
