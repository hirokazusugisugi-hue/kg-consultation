/**
 * 日程管理機能（拡張版）
 * スプレッドシートの「日程設定」シートと連携
 * 配置点数による予約可否判定を追加
 */

/**
 * 日程設定シートのヘッダーを設定（拡張版）
 */
function setupScheduleSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SCHEDULE_SHEET_NAME);
  }

  // ヘッダー設定（拡張）
  const headers = [
    '日付', '時間帯', '対応可否', '対応方法', '担当者', '予約状況', '備考',
    '参加メンバー', '配置点数', '特別対応フラグ', '予約可能判定'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 120);  // 日付
  sheet.setColumnWidth(2, 100);  // 時間帯
  sheet.setColumnWidth(3, 80);   // 対応可否
  sheet.setColumnWidth(4, 100);  // 対応方法
  sheet.setColumnWidth(5, 100);  // 担当者
  sheet.setColumnWidth(6, 80);   // 予約状況
  sheet.setColumnWidth(7, 200);  // 備考
  sheet.setColumnWidth(8, 250);  // 参加メンバー
  sheet.setColumnWidth(9, 80);   // 配置点数
  sheet.setColumnWidth(10, 100); // 特別対応フラグ
  sheet.setColumnWidth(11, 100); // 予約可能判定

  // 対応可否のプルダウン
  const availabilityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['○', '×'])
    .build();
  sheet.getRange(2, 3, 500, 1).setDataValidation(availabilityRule);

  // 対応方法のプルダウン
  const methodRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['対面', 'オンライン', '両方'])
    .build();
  sheet.getRange(2, 4, 500, 1).setDataValidation(methodRule);

  // 予約状況のプルダウン
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['空き', '予約済み'])
    .build();
  sheet.getRange(2, 6, 500, 1).setDataValidation(statusRule);

  // 特別対応フラグのプルダウン
  const specialRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'])
    .build();
  sheet.getRange(2, 10, 500, 1).setDataValidation(specialRule);

  // 条件付き書式（対応可否）
  const rules = [];
  const range = sheet.getRange('C2:C500');

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('○')
    .setBackground('#d4edda')
    .setRanges([range])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('×')
    .setBackground('#f8d7da')
    .setRanges([range])
    .build());

  // 予約可能判定の条件付き書式
  const bookableRange = sheet.getRange('K2:K500');

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('○')
    .setBackground('#d4edda')
    .setRanges([bookableRange])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('×')
    .setBackground('#f8d7da')
    .setRanges([bookableRange])
    .build());

  sheet.setConditionalFormatRules(rules);

  // 1行目を固定
  sheet.setFrozenRows(1);

  console.log('日程設定シート（拡張版）のセットアップが完了しました');
}

/**
 * 日程設定シートの配置点数と予約可能判定を再計算
 */
function recalculateScheduleScores() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!sheet) return;

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const memberNames = data[i][SCHEDULE_COLUMNS.MEMBERS];
    const specialFlag = data[i][SCHEDULE_COLUMNS.SPECIAL_FLAG] === true ||
                        data[i][SCHEDULE_COLUMNS.SPECIAL_FLAG] === 'TRUE';

    if (memberNames) {
      const score = calculateStaffScore(memberNames.toString());
      const bookable = getBookableStatus(score, specialFlag);

      // I列: 配置点数を更新
      sheet.getRange(i + 1, SCHEDULE_COLUMNS.SCORE + 1).setValue(score);
      // K列: 予約可能判定を更新
      sheet.getRange(i + 1, SCHEDULE_COLUMNS.BOOKABLE + 1).setValue(bookable);
    }
  }

  console.log('配置点数と予約可能判定を再計算しました');
}

/**
 * 指定月のN回目の曜日かを返す
 * @param {number} year - 年
 * @param {number} month - 月（0-indexed）
 * @param {number} day - 日
 * @param {number} dow - 曜日（0=日, 6=土）
 * @returns {number} その月で何回目のその曜日か（1-indexed）
 */
function getNthDayOfWeek(year, month, day, dow) {
  let count = 0;
  for (let d = 1; d <= day; d++) {
    if (new Date(year, month, d).getDay() === dow) count++;
  }
  return count;
}

/**
 * 指定年月の日程データを生成
 * ルール:
 *   第1・第3 土日 → 14:00, 16:00, 18:00（両方）
 *   第2・第4 金曜 → 19:00, 20:30（Zoomのみ）
 * @param {number} year - 年
 * @param {number} month - 月（1-12）
 * @returns {Array} 日程データの2次元配列
 */
function generateScheduleData(year, month) {
  const data = [];
  const daysInMonth = new Date(year, month, 0).getDate();

  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month - 1, day);
    const dow = date.getDay();
    const dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');

    // 土日: 第1・第3のみ
    if (dow === 0 || dow === 6) {
      const nth = getNthDayOfWeek(year, month - 1, day, dow);
      if (nth === 1 || nth === 3) {
        ['14:00', '16:00', '18:00'].forEach(function(time) {
          data.push([dateStr, time, '○', '両方', '', '空き', '', '', 0, 'FALSE', '×']);
        });
      }
    }

    // 金曜: 第2・第4のみ（Zoomのみ）
    if (dow === 5) {
      const nth = getNthDayOfWeek(year, month - 1, day, dow);
      if (nth === 2 || nth === 4) {
        ['19:00', '20:30'].forEach(function(time) {
          data.push([dateStr, time, '○', 'オンライン', '', '空き', '', '', 0, 'FALSE', '×']);
        });
      }
    }
  }

  return data;
}

/**
 * 3月の日程サンプルを生成（新ルール版）
 */
function generateMarchSchedule() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!sheet) {
    console.log('日程設定シートが見つかりません。setupScheduleSheet()を実行してください。');
    return;
  }

  const data = generateScheduleData(2026, 3);

  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  }

  console.log(`3月の日程を${data.length}件生成しました`);
}

/**
 * 利用可能な日程を取得（API用・拡張版）
 * 予約可能判定が「○」の枠のみ返す
 * @param {string} method - 'visit' または 'zoom'
 * @returns {Object} 利用可能な日程データ
 */
function getAvailableSchedule(method) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!sheet) {
    return { error: '日程設定シートが見つかりません' };
  }

  const data = sheet.getDataRange().getValues();

  // 2ヶ月先までの制限
  const now = new Date();
  const maxDate = new Date(now.getFullYear(), now.getMonth() + 2, now.getDate());

  const available = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateStr = row[SCHEDULE_COLUMNS.DATE];
    const time = row[SCHEDULE_COLUMNS.TIME];
    const isAvailable = row[SCHEDULE_COLUMNS.AVAILABLE];
    const supportedMethod = row[SCHEDULE_COLUMNS.METHOD];
    const staff = row[SCHEDULE_COLUMNS.STAFF];
    const bookingStatus = row[SCHEDULE_COLUMNS.BOOKING_STATUS];
    const bookable = row[SCHEDULE_COLUMNS.BOOKABLE];

    // 2ヶ月先より後の日程は除外
    const rowDate = dateStr instanceof Date ? dateStr : new Date(dateStr);
    if (rowDate > maxDate) {
      continue;
    }

    // 対応可否が「○」かつ予約状況が「空き」かつ予約可能判定が「○」の場合のみ
    if (isAvailable !== '○' || bookingStatus === '予約済み') {
      continue;
    }

    // 予約可能判定チェック（配置点数ベース）
    if (bookable !== '○') {
      continue;
    }

    // 対応方法のチェック
    if (method === 'visit' && supportedMethod === 'オンライン') {
      continue;
    }
    if (method === 'zoom' && supportedMethod === '対面') {
      continue;
    }

    // 日付をキーにしてグループ化
    let dateKey;
    if (dateStr instanceof Date) {
      dateKey = Utilities.formatDate(dateStr, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else {
      const d = new Date(dateStr);
      dateKey = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    }

    if (!available[dateKey]) {
      available[dateKey] = [];
    }

    // 時間をフォーマット
    let timeStr;
    if (time instanceof Date) {
      timeStr = Utilities.formatDate(time, 'Asia/Tokyo', 'HH:mm');
    } else {
      timeStr = String(time);
    }

    available[dateKey].push({
      time: timeStr,
      staff: staff || ''
    });
  }

  return available;
}

/**
 * 予約状況を更新
 * @param {string} dateStr - 日付（yyyy/MM/dd形式）
 * @param {string} time - 時間帯
 */
function markAsBooked(dateStr, time) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!sheet) {
    return false;
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    let rowDate = data[i][SCHEDULE_COLUMNS.DATE];
    let rowTime = data[i][SCHEDULE_COLUMNS.TIME];

    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, 'Asia/Tokyo', 'yyyy/MM/dd');
    }

    if (rowTime instanceof Date) {
      rowTime = Utilities.formatDate(rowTime, 'Asia/Tokyo', 'HH:mm');
    } else {
      rowTime = String(rowTime);
    }

    if (rowDate === dateStr && rowTime === time) {
      sheet.getRange(i + 1, SCHEDULE_COLUMNS.BOOKING_STATUS + 1).setValue('予約済み');
      return true;
    }
  }

  return false;
}

/**
 * 翌々月の日程を自動生成（時間トリガー用・新ルール版）
 */
function generateNextMonthSchedule() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!sheet) {
    return;
  }

  const now = new Date();
  const targetDate = new Date(now.getFullYear(), now.getMonth() + 2, 1);
  const year = targetDate.getFullYear();
  const month = targetDate.getMonth() + 1;

  const data = generateScheduleData(year, month);

  const lastRow = sheet.getLastRow();
  if (data.length > 0) {
    sheet.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
  }

  console.log(`${year}年${month}月の日程を${data.length}件追加しました`);
}
