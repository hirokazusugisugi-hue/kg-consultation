/**
 * 回答集計シート ↔ 日程設定シート 双方向同期
 * 回答集計シートの○/×編集を日程設定に即時反映
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// onEdit トリガーハンドラ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 回答集計シート編集時のハンドラ（installable trigger）
 * ○/×変更を日程設定シートに即時反映し、参加数・配置点数も更新
 */
function onSummarySheetEdit(e) {
  try {
    if (!e || !e.source) return;

    var activeSheet = e.source.getActiveSheet();
    if (activeSheet.getName() !== CONFIG.SUMMARY_SHEET_NAME) return;

    var range = e.range;
    var row = range.getRow();
    var col = range.getColumn();
    var value = String(range.getValue()).trim();

    // ○ or × のみ処理
    if (value !== '○' && value !== '×') return;

    // メンバー列は6列目（F列）以降
    if (col < 6) return;

    // ヘッダー行を探す（上方向に'判定'を検索）
    var headerRow = findSummaryHeaderRow(activeSheet, row);
    if (!headerRow) return;

    // メンバー列かチェック（参加数・配置点数列は除外）
    var headerValue = String(activeSheet.getRange(headerRow, col).getValue()).trim();
    if (!headerValue || headerValue === '参加数' || headerValue === '配置点数') return;

    // データ行かチェック（日付と時間がある行）
    var dateLabel = String(activeSheet.getRange(row, 1).getValue()).trim();
    var timeVal = activeSheet.getRange(row, 2).getValue();
    var timeStr = timeVal instanceof Date
      ? Utilities.formatDate(timeVal, 'Asia/Tokyo', 'HH:mm')
      : String(timeVal).trim();
    var timeParsed = timeStr.match(/(\d{1,2}:\d{2})/);
    if (timeParsed) timeStr = timeParsed[1];
    if (!dateLabel || !timeStr) return;

    // 月コンテキスト取得（年・月）
    var monthInfo = findSummaryMonthContext(activeSheet, row);
    if (!monthInfo) return;

    // 日付パース: "2/27(金)" → day=27
    var dateMatch = dateLabel.match(/(\d{1,2})\/(\d{1,2})/);
    if (!dateMatch) return;
    var day = parseInt(dateMatch[2]);

    var targetDate = monthInfo.year + '/' +
      String(monthInfo.month).padStart(2, '0') + '/' +
      String(day).padStart(2, '0');

    // 苗字→フルネームマッピング
    var fullName = findMemberByShortName(headerValue);
    if (!fullName) {
      console.log('メンバー不明: ' + headerValue);
      return;
    }

    // 日程設定シートを更新
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    if (!scheduleSheet) return;

    var updated = updateScheduleMemberByDatetime(scheduleSheet, targetDate, timeStr, fullName, value === '○');

    if (updated) {
      // 回答集計の参加数・配置点数を更新
      updateSummaryRowCounts(activeSheet, row, headerRow);
      console.log('同期完了: ' + targetDate + ' ' + timeStr + ' ' + fullName + ' → ' + value);
    }

  } catch (err) {
    console.log('onSummarySheetEdit error: ' + err.message);
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 一括同期（回答集計→日程設定）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 回答集計シートの全データを日程設定シートに一括同期
 * 既存の編集内容を日程設定に反映する
 * @returns {Object} 処理結果
 */
function syncAllSummaryToSchedule() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var summarySheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!summarySheet || !scheduleSheet) {
    return { success: false, message: 'シートが見つかりません' };
  }

  var summaryData = summarySheet.getDataRange().getValues();
  var scheduleData = scheduleSheet.getDataRange().getValues();
  var allMembers = getScheduleMembers();

  var currentYear = null;
  var currentMonth = null;
  var memberFullNames = []; // ヘッダー順のフルネーム配列
  var MEMBER_START_COL = 5; // 0-indexed（F列＝6列目）
  var updatedCount = 0;
  var details = [];

  for (var i = 0; i < summaryData.length; i++) {
    var row = summaryData[i];

    // 月タイトル行: "2026年2月 回答集計"
    var titleMatch = String(row[0]).match(/(\d{4})年(\d{1,2})月/);
    if (titleMatch) {
      currentYear = parseInt(titleMatch[1]);
      currentMonth = parseInt(titleMatch[2]);
      continue;
    }

    // ヘッダー行: 5列目（E）が'判定'
    if (row[4] === '判定') {
      memberFullNames = [];
      for (var c = MEMBER_START_COL; c < row.length; c++) {
        var hdr = String(row[c]).trim();
        if (!hdr || hdr === '参加数' || hdr === '配置点数') break;
        var fn = findMemberByShortName(hdr);
        memberFullNames.push(fn); // null if not found
      }
      continue;
    }

    // データ行チェック
    var dateLabel = String(row[0]).trim();
    var timeVal = row[1];
    var timeStr = timeVal instanceof Date
      ? Utilities.formatDate(timeVal, 'Asia/Tokyo', 'HH:mm')
      : String(timeVal).trim();
    // HH:mm形式の抽出（Date文字列やゴミが混入した場合の対応）
    var timeMatch = timeStr.match(/(\d{1,2}:\d{2})/);
    if (timeMatch) timeStr = timeMatch[1];
    if (!dateLabel || !timeStr || !currentYear || !currentMonth) continue;

    var dateMatch = dateLabel.match(/(\d{1,2})\/(\d{1,2})/);
    if (!dateMatch) continue;

    var day = parseInt(dateMatch[2]);
    var targetDate = currentYear + '/' +
      String(currentMonth).padStart(2, '0') + '/' +
      String(day).padStart(2, '0');

    // メンバー列から参加者リストを構築
    var newMembers = [];
    for (var m = 0; m < memberFullNames.length; m++) {
      var cellVal = row[MEMBER_START_COL + m];
      if (cellVal === '○' && memberFullNames[m]) {
        newMembers.push(memberFullNames[m]);
      }
    }
    var newMembersStr = newMembers.join(', ');

    // 日程設定シートで該当行を検索し更新
    for (var s = 1; s < scheduleData.length; s++) {
      var schedDate = scheduleData[s][SCHEDULE_COLUMNS.DATE];
      var schedDateStr = schedDate instanceof Date
        ? Utilities.formatDate(schedDate, 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(schedDate);

      var schedTime = scheduleData[s][SCHEDULE_COLUMNS.TIME];
      var schedTimeStr = schedTime instanceof Date
        ? Utilities.formatDate(schedTime, 'Asia/Tokyo', 'HH:mm')
        : String(schedTime);

      if (schedDateStr === targetDate && schedTimeStr === timeStr) {
        var oldMembers = scheduleData[s][SCHEDULE_COLUMNS.MEMBERS]
          ? scheduleData[s][SCHEDULE_COLUMNS.MEMBERS].toString() : '';

        if (oldMembers !== newMembersStr) {
          scheduleSheet.getRange(s + 1, SCHEDULE_COLUMNS.MEMBERS + 1).setValue(newMembersStr);
          scheduleData[s][SCHEDULE_COLUMNS.MEMBERS] = newMembersStr;
          recalculateScheduleScoreForRow(s + 1, scheduleSheet);
          updatedCount++;
          details.push(targetDate + ' ' + timeStr + ': ' + newMembersStr);
        }
        break;
      }
    }

    // 回答集計の参加数・配置点数も更新
    var countCol = MEMBER_START_COL + memberFullNames.length;
    var scoreCol = countCol + 1;
    summarySheet.getRange(i + 1, countCol + 1).setValue(newMembers.length);

    var score = calculateStaffScore(newMembersStr);
    summarySheet.getRange(i + 1, scoreCol + 1).setValue(score);
  }

  var message = updatedCount + '件の日程を更新しました';
  if (details.length > 0) {
    message += '\n' + details.join('\n');
  }

  console.log('一括同期完了: ' + message);
  return { success: true, updated: updatedCount, message: message };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ヘルパー関数
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 指定行から上方向にヘッダー行を検索
 * @param {Sheet} sheet - 回答集計シート
 * @param {number} row - 現在の行番号（1-based）
 * @returns {number|null} ヘッダー行番号
 */
function findSummaryHeaderRow(sheet, row) {
  for (var r = row - 1; r >= 1; r--) {
    var val = sheet.getRange(r, 5).getValue();
    if (val === '判定') return r;
    // 月タイトル行を超えたら探索停止
    var firstCell = String(sheet.getRange(r, 1).getValue());
    if (firstCell.match(/\d{4}年\d{1,2}月/)) return null;
  }
  return null;
}

/**
 * 指定行から上方向に月コンテキスト（年・月）を検索
 * @param {Sheet} sheet - 回答集計シート
 * @param {number} row - 現在の行番号（1-based）
 * @returns {Object|null} { year, month }
 */
function findSummaryMonthContext(sheet, row) {
  for (var r = row - 1; r >= 1; r--) {
    var val = String(sheet.getRange(r, 1).getValue());
    var match = val.match(/(\d{4})年(\d{1,2})月/);
    if (match) {
      return { year: parseInt(match[1]), month: parseInt(match[2]) };
    }
  }
  return null;
}

/**
 * 苗字からフルネームを検索
 * @param {string} shortName - 苗字（例: "杉山"）
 * @returns {string|null} フルネーム（例: "杉山 宏和"）
 */
function findMemberByShortName(shortName) {
  if (!shortName) return null;
  var members = getScheduleMembers();

  // 苗字（姓）で完全一致
  for (var i = 0; i < members.length; i++) {
    var parts = members[i].name.split(/[\s　]+/);
    if (parts[0] === shortName) return members[i].name;
  }

  // フルネーム完全一致
  for (var i = 0; i < members.length; i++) {
    if (members[i].name === shortName) return members[i].name;
  }

  // 前方一致
  for (var i = 0; i < members.length; i++) {
    if (members[i].name.indexOf(shortName) === 0) return members[i].name;
  }

  return null;
}

/**
 * 日程設定シートの特定日時・メンバーを更新
 * @param {Sheet} scheduleSheet - 日程設定シート
 * @param {string} targetDate - 日付（yyyy/MM/dd）
 * @param {string} timeStr - 時間（HH:mm）
 * @param {string} memberName - メンバーフルネーム
 * @param {boolean} isParticipating - true=○（参加）, false=×（不参加）
 * @returns {boolean} 更新成功
 */
function updateScheduleMemberByDatetime(scheduleSheet, targetDate, timeStr, memberName, isParticipating) {
  var data = scheduleSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var schedDate = data[i][SCHEDULE_COLUMNS.DATE];
    var schedDateStr = schedDate instanceof Date
      ? Utilities.formatDate(schedDate, 'Asia/Tokyo', 'yyyy/MM/dd')
      : String(schedDate);

    var schedTime = data[i][SCHEDULE_COLUMNS.TIME];
    var schedTimeStr = schedTime instanceof Date
      ? Utilities.formatDate(schedTime, 'Asia/Tokyo', 'HH:mm')
      : String(schedTime);

    if (schedDateStr !== targetDate || schedTimeStr !== timeStr) continue;

    var currentMembers = data[i][SCHEDULE_COLUMNS.MEMBERS]
      ? data[i][SCHEDULE_COLUMNS.MEMBERS].toString() : '';
    var memberList = currentMembers
      ? currentMembers.split(',').map(function(n) { return n.trim(); }).filter(function(n) { return n; })
      : [];

    var idx = memberList.indexOf(memberName);

    if (isParticipating && idx === -1) {
      // 追加
      memberList.push(memberName);
    } else if (!isParticipating && idx !== -1) {
      // 削除
      memberList.splice(idx, 1);
    } else {
      // 変更なし
      return true;
    }

    var newMembersStr = memberList.join(', ');
    scheduleSheet.getRange(i + 1, SCHEDULE_COLUMNS.MEMBERS + 1).setValue(newMembersStr);
    recalculateScheduleScoreForRow(i + 1, scheduleSheet);
    return true;
  }

  console.log('該当行なし: ' + targetDate + ' ' + timeStr);
  return false;
}

/**
 * 回答集計の参加数・配置点数を更新
 * @param {Sheet} sheet - 回答集計シート
 * @param {number} dataRow - データ行番号（1-based）
 * @param {number} headerRow - ヘッダー行番号（1-based）
 */
function updateSummaryRowCounts(sheet, dataRow, headerRow) {
  var headerValues = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var dataValues = sheet.getRange(dataRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // メンバー列の範囲を特定
  var memberStartCol = 5; // 0-indexed
  var participantCount = 0;
  var participantNames = [];

  for (var c = memberStartCol; c < headerValues.length; c++) {
    var hdr = String(headerValues[c]).trim();
    if (!hdr || hdr === '参加数' || hdr === '配置点数') {
      // 参加数列 = c, 配置点数列 = c+1
      sheet.getRange(dataRow, c + 1).setValue(participantCount);

      // 配置点数を計算
      var fullNames = participantNames.map(function(sn) {
        return findMemberByShortName(sn);
      }).filter(function(n) { return n; });
      var score = calculateStaffScore(fullNames.join(', '));
      sheet.getRange(dataRow, c + 2).setValue(score);

      // 判定列（E列）も更新
      var specialFlag = false; // 回答集計からは特別対応フラグは読まない
      var bookable = getBookableStatus(score, specialFlag);
      // 予約済みでなければ判定を更新
      var bookingStatus = String(dataValues[3]).trim();
      if (bookingStatus !== '予約済み') {
        sheet.getRange(dataRow, 5).setValue(bookable);
        // セル色
        var cell = sheet.getRange(dataRow, 5);
        if (bookable === '○') {
          cell.setBackground('#d4edda');
        } else {
          cell.setBackground('#f8d7da');
        }
      }
      break;
    }

    if (dataValues[c] === '○') {
      participantCount++;
      participantNames.push(hdr);
    }
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// トリガー設定
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 回答集計シートのonEdit installableトリガーを設定
 */
function setupSummaryEditTrigger() {
  // 既存トリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'onSummarySheetEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // installable onEdit trigger
  ScriptApp.newTrigger('onSummarySheetEdit')
    .forSpreadsheet(CONFIG.SPREADSHEET_ID)
    .onEdit()
    .create();

  console.log('回答集計シートのonEditトリガーを設定しました');
  return { success: true, message: 'onEditトリガー設定完了' };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// テスト
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 回答集計↔日程設定の双方向同期テスト
 * 1. 回答集計シートで特定メンバーを×に変更
 * 2. 一括同期で日程設定に反映されるか確認
 * 3. 元に戻す
 * @returns {Object} テスト結果
 */
function testSummarySyncRoundtrip() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var summarySheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!summarySheet || !scheduleSheet) {
    return { success: false, message: 'シートが見つかりません' };
  }

  var results = [];

  // Step 1: 回答集計シートから2月のデータ行を探す
  var summaryData = summarySheet.getDataRange().getValues();
  var targetRow = -1;
  var targetCol = -1;
  var headerRow = -1;
  var memberFullName = null;

  for (var i = 0; i < summaryData.length; i++) {
    if (summaryData[i][4] === '判定') {
      headerRow = i;
    }
    // 2/27の行を探す
    var dateLabel = String(summaryData[i][0]);
    if (dateLabel.match(/2\/27/)) {
      targetRow = i;
      // 最初のメンバー列（F列=index5）で○のセルを探す
      for (var c = 5; c < summaryData[i].length; c++) {
        if (summaryData[i][c] === '○') {
          targetCol = c;
          break;
        }
      }
      break;
    }
  }

  if (targetRow === -1 || targetCol === -1 || headerRow === -1) {
    return { success: false, message: '2/27のデータ行が見つかりません (row=' + targetRow + ', col=' + targetCol + ')' };
  }

  var memberShortName = String(summaryData[headerRow][targetCol]);
  memberFullName = findMemberByShortName(memberShortName);
  var timeVal = summaryData[targetRow][1];
  var timeStr = timeVal instanceof Date
    ? Utilities.formatDate(timeVal, 'Asia/Tokyo', 'HH:mm')
    : String(timeVal).trim();
  var tMatch = timeStr.match(/(\d{1,2}:\d{2})/);
  if (tMatch) timeStr = tMatch[1];

  results.push('対象: 2/27 ' + timeStr + ' ' + memberFullName + ' (col=' + (targetCol + 1) + ')');
  results.push('変更前: ○');

  // Step 2: 回答集計で○→×に変更
  summarySheet.getRange(targetRow + 1, targetCol + 1).setValue('×');
  results.push('回答集計を×に変更');

  // Step 3: 一括同期実行
  var syncResult = syncAllSummaryToSchedule();
  results.push('同期結果: ' + syncResult.message);

  // Step 4: 日程設定シートを確認
  var schedData = scheduleSheet.getDataRange().getValues();
  var verified = false;
  for (var s = 1; s < schedData.length; s++) {
    var sDate = schedData[s][SCHEDULE_COLUMNS.DATE];
    var sDateStr = sDate instanceof Date
      ? Utilities.formatDate(sDate, 'Asia/Tokyo', 'yyyy/MM/dd') : String(sDate);
    var sTime = schedData[s][SCHEDULE_COLUMNS.TIME];
    var sTimeStr = sTime instanceof Date
      ? Utilities.formatDate(sTime, 'Asia/Tokyo', 'HH:mm') : String(sTime);

    // sTimeStrからHH:mmのみ抽出
    var sTimeParsed = sTimeStr.match(/(\d{1,2}:\d{2})/);
    if (sTimeParsed) sTimeStr = sTimeParsed[1];
    if (sDateStr === '2026/02/27' && sTimeStr === timeStr) {
      var members = schedData[s][SCHEDULE_COLUMNS.MEMBERS]
        ? schedData[s][SCHEDULE_COLUMNS.MEMBERS].toString() : '';
      if (members.indexOf(memberFullName) === -1) {
        results.push('検証OK: 日程設定から ' + memberFullName + ' が削除されている');
        verified = true;
      } else {
        results.push('検証NG: 日程設定に ' + memberFullName + ' がまだ残っている: ' + members);
      }
      break;
    }
  }

  if (!verified) {
    results.push('検証NG: 対象行が見つからない');
  }

  // Step 5: 元に戻す（○に復元）
  summarySheet.getRange(targetRow + 1, targetCol + 1).setValue('○');
  syncAllSummaryToSchedule();
  results.push('復元完了（○に戻し、再同期済み）');

  return { success: verified, results: results };
}
