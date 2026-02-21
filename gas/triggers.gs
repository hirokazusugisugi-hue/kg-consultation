/**
 * トリガー処理（拡張版）
 * リマインド前日/3日前、担当者個別通知対応
 */

/**
 * onEditトリガー設定用関数
 */
function setupOnEditTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSheetEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onSheetEdit')
    .forSpreadsheet(CONFIG.SPREADSHEET_ID)
    .onEdit()
    .create();

  console.log('onEditトリガーを設定しました');
}

/**
 * 編集時の処理
 */
function onSheetEdit(e) {
  const sheet = e.source.getActiveSheet();

  // 予約管理シート以外は無視
  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    // 日程設定シートの参加メンバー列が変更された場合、配置点数を再計算
    if (sheet.getName() === CONFIG.SCHEDULE_SHEET_NAME) {
      const col = e.range.getColumn();
      // H列（参加メンバー）またはJ列（特別対応フラグ）が変更された場合
      if (col === SCHEDULE_COLUMNS.MEMBERS + 1 || col === SCHEDULE_COLUMNS.SPECIAL_FLAG + 1) {
        recalculateScheduleScoreForRow(e.range.getRow(), sheet);
      }
    }
    return;
  }

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  // ヘッダー行は無視
  if (row === 1) {
    return;
  }

  // ステータス列の変更を検知
  if (col === COLUMNS.STATUS + 1) {
    handleStatusChange(row, e.oldValue, e.value);
  }
}

/**
 * 日程設定シートの特定行の配置点数・予約可能判定を再計算
 */
function recalculateScheduleScoreForRow(rowIndex, sheet) {
  if (rowIndex <= 1) return;

  const row = sheet.getRange(rowIndex, 1, 1, 11).getValues()[0];
  const memberNames = row[SCHEDULE_COLUMNS.MEMBERS];
  const specialFlag = row[SCHEDULE_COLUMNS.SPECIAL_FLAG] === true ||
                      row[SCHEDULE_COLUMNS.SPECIAL_FLAG] === 'TRUE';

  if (memberNames) {
    const score = calculateStaffScore(memberNames.toString());
    const diagCount = countDiagnosticians(memberNames.toString());
    const bookable = getBookableStatus(score, specialFlag, diagCount);

    sheet.getRange(rowIndex, SCHEDULE_COLUMNS.SCORE + 1).setValue(score);
    sheet.getRange(rowIndex, SCHEDULE_COLUMNS.BOOKABLE + 1).setValue(bookable);
  } else {
    sheet.getRange(rowIndex, SCHEDULE_COLUMNS.SCORE + 1).setValue(0);
    sheet.getRange(rowIndex, SCHEDULE_COLUMNS.BOOKABLE + 1).setValue('×');
  }
}

/**
 * ステータス変更時の処理
 */
function handleStatusChange(rowIndex, oldStatus, newStatus) {
  console.log(`ステータス変更: ${oldStatus} → ${newStatus} (行: ${rowIndex})`);

  const data = getRowData(rowIndex);

  // 確定に変更された場合
  if (newStatus === STATUS.CONFIRMED) {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
      .getSheetByName(CONFIG.SHEET_NAME);

    // P列が空の場合、K列（希望日時1）+ 日程設定シートから時間帯を補完して自動設定
    if (!data.confirmedDate) {
      const fullDateTime = resolveConfirmedDateTime(rowIndex, sheet);
      if (fullDateTime) {
        sheet.getRange(rowIndex, COLUMNS.CONFIRMED_DATE + 1).setValue(fullDateTime);
        data.confirmedDate = fullDateTime;
      } else {
        SpreadsheetApp.getUi().alert(
          '確定日時が入力されていません。\n希望日時（K列）または確定日時（P列）を入力してからステータスを変更してください。'
        );
        sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(oldStatus || STATUS.PENDING);
        return;
      }
    }

    // オフライン相談の場合、場所（X列）が必須
    const method = data.method || '';
    const isOnline = method.indexOf('オンライン') >= 0 || method.indexOf('Zoom') >= 0 || method.indexOf('zoom') >= 0;
    if (!isOnline) {
      const location = sheet.getRange(rowIndex, COLUMNS.LOCATION + 1).getValue();
      if (!location || location === '') {
        SpreadsheetApp.getUi().alert(
          '対面相談の場合、場所（N列）を設定してからステータスを「確定」に変更してください。\n\n選択肢: アプローズタワー / スミセスペース / ナレッジサロン / その他'
        );
        sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(oldStatus || STATUS.PENDING);
        return;
      }
      data.location = location;
    }

    // オンライン相談の場合、Zoomミーティングを自動作成
    if (isOnline && !data.zoomUrl) {
      var zoomUrl = createAndSaveZoomMeeting(data, rowIndex);
      if (zoomUrl) {
        SpreadsheetApp.getUi().alert(
          '【Zoom自動作成完了】\n\nZoomミーティングを自動作成しました。\nR列にURLが保存されました。\n\nURL: ' + zoomUrl
        );
      } else {
        SpreadsheetApp.getUi().alert(
          '【Zoom自動作成失敗】\n\nZoomミーティングの自動作成に失敗しました。\nZoom API設定を確認するか、R列にURLを手動入力してください。\n\n※確定メールには「前日までにお送りします」と記載されます。'
        );
      }
    }

    // 日程設定シートの予約状況を「予約済み」に更新
    var parsed = parseConfirmedDateTime(data.confirmedDate);
    if (parsed.date) {
      var booked = markAsBooked(parsed.date, parsed.time);
      console.log(`日程設定シート同期: ${parsed.date} ${parsed.time || ''} → ${booked ? '予約済み' : '該当なし'}`);
    }

    // 確定メールを送信
    sendConfirmedEmail(data);
    console.log(`確定メール送信完了: ${data.email}`);

    // 担当者・管理者への通知（企業URL情報を含む）
    const staffLineMsg = `✅ 予約確定

申込ID: ${data.id}
お名前: ${data.name}様
貴社名: ${data.company}
日時: ${data.confirmedDate}
方法: ${data.method}
テーマ: ${data.theme}
${data.companyUrl ? '企業URL: ' + data.companyUrl + '\n※事前リサーチをお願いします' : ''}`;

    const staffIsOnline = method.indexOf('オンライン') >= 0 || method.indexOf('Zoom') >= 0 || method.indexOf('zoom') >= 0;
    const staffZoomLine = staffIsOnline && data.zoomUrl ? '\nZoom URL：' + data.zoomUrl : '';

    const confirmEmailBody = `予約が確定しました。

申込ID：${data.id}
お名前：${data.name}様
貴社名：${data.company}
日時：${data.confirmedDate}
相談方法：${data.method}
テーマ：${data.theme}${staffZoomLine}
${data.companyUrl ? '\n企業URL：' + data.companyUrl + '\n※事前リサーチにAIツールの活用を推奨します' : ''}

事前準備をお願いいたします。`;

    const confirmEmailSubject = `【予約確定】${data.name}様（${data.company}） - ${data.confirmedDate}`;

    var confirmedSentEmails = {};

    // P列の担当者に通知
    if (data.staff) {
      sendStaffNotifications(
        data.staff,
        staffLineMsg,
        confirmEmailSubject,
        confirmEmailBody
      );
      var cStaffNames = data.staff.split(',').map(function(n) { return n.trim(); }).filter(function(n) { return n; });
      cStaffNames.forEach(function(name) {
        var m = getMemberByName(name);
        if (m && m.email) confirmedSentEmails[m.email] = true;
      });
    }

    // 日程設定シートの参加メンバー全員にも確定通知（P列と重複しない分）
    var confirmMembers = getScheduleMembersForDate_(data.confirmedDate);
    if (confirmMembers && confirmMembers.length > 0) {
      confirmMembers.forEach(function(cm) {
        var m = getMemberByName(cm.name);
        if (m && m.email && !confirmedSentEmails[m.email]) {
          GmailApp.sendEmail(m.email, confirmEmailSubject, confirmEmailBody, {
            name: CONFIG.SENDER_NAME
          });
          if (m.lineId) sendLineMessage(m.lineId, staffLineMsg);
          confirmedSentEmails[m.email] = true;
          console.log('確定通知（参加メンバー）: ' + cm.name + ' (' + m.email + ')');
        }
      });
    }

    // 誰にも送れなかった場合は管理者にフォールバック
    if (Object.keys(confirmedSentEmails).length === 0) {
      CONFIG.ADMIN_EMAILS.forEach(function(adminEmail) {
        GmailApp.sendEmail(adminEmail, confirmEmailSubject, confirmEmailBody, {
          name: CONFIG.SENDER_NAME
        });
      });
      console.log('確定通知: 担当者・メンバー未設定のため管理者にフォールバック');
    }

    // リーダー履歴に「予定」として記録
    if (data.leader) {
      var schedMembers = getParticipatingMembers(data.confirmedDate) || '';
      recordLeaderAssignment(data, data.leader, schedMembers, 0, '手動設定', '予定');
    }

    // LINE通知
    sendLineStatusNotification(data, newStatus);
  }

  // 書類受領に変更された場合
  if (newStatus === STATUS.RECEIVED) {
    const subject = `【書類受領】${data.name}様 - 書類受領`;
    const body = `書類を受領しました。

申込ID：${data.id}
お名前：${data.name}様
貴社名：${data.company}

日程を調整し、ステータスを「確定」に変更してください。`;

    CONFIG.ADMIN_EMAILS.forEach(email => {
      GmailApp.sendEmail(email, subject, body, {
        name: CONFIG.SENDER_NAME
      });
    });

    sendLineStatusNotification(data, newStatus);
  }

  // 完了に変更された場合
  if (newStatus === STATUS.COMPLETED) {
    // リーダー自動選定
    try {
      autoSelectLeaderOnComplete(rowIndex);
    } catch (leaderError) {
      console.error('リーダー選定エラー:', leaderError);
    }

    sendLineStatusNotification(data, newStatus);
  }

  // キャンセルに変更された場合
  if (newStatus === STATUS.CANCELLED) {
    processCancellation(rowIndex, data, '手動変更');
  }
}

/**
 * 毎日のリマインド送信
 * 相談者向け: 3日前 + 前日
 * 担当者向け: 1週間前 + 3日前
 */
function sendDailyReminders() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  const now = new Date();

  // 翌日の日付（相談者・前日リマインド用）
  const oneDayLater = new Date(now);
  oneDayLater.setDate(oneDayLater.getDate() + 1);
  const oneDayLaterStr = Utilities.formatDate(oneDayLater, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 3日後の日付（相談者・3日前 + 担当者・3日前）
  const threeDaysLater = new Date(now);
  threeDaysLater.setDate(threeDaysLater.getDate() + 3);
  const threeDaysLaterStr = Utilities.formatDate(threeDaysLater, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 7日後の日付（担当者・1週間前）
  const sevenDaysLater = new Date(now);
  sevenDaysLater.setDate(sevenDaysLater.getDate() + 7);
  const sevenDaysLaterStr = Utilities.formatDate(sevenDaysLater, 'Asia/Tokyo', 'yyyy/MM/dd');

  for (let i = 1; i < data.length; i++) {
    const status = data[i][COLUMNS.STATUS];
    const confirmedDate = data[i][COLUMNS.CONFIRMED_DATE];

    if ((status !== STATUS.CONFIRMED && status !== STATUS.NDA_AGREED) || !confirmedDate) {
      continue;
    }

    const dateStr = (confirmedDate instanceof Date || (typeof confirmedDate === 'object' && typeof confirmedDate.getTime === 'function'))
      ? Utilities.formatDate(confirmedDate, 'Asia/Tokyo', 'yyyy/MM/dd')
      : confirmedDate.toString().substring(0, 10).replace(/-/g, '/');
    const rowData = getRowData(i + 1);

    // ── 担当者向け: 1週間前リマインド ──
    if (dateStr === sevenDaysLaterStr) {
      sendStaffReminderWithMembers_(rowData, '1週間前');
      console.log(`担当者1週間前リマインド送信: ${rowData.id}`);
    }

    // ── 相談者 + 担当者向け: 3日前リマインド ──
    if (dateStr === threeDaysLaterStr) {
      // 相談者向け（メール）
      sendReminderEmail3DaysBefore(rowData);

      // 担当者向け
      sendStaffReminderWithMembers_(rowData, '3日前');

      console.log(`3日前リマインド送信: ${rowData.email}`);
    }

    // ── 相談者向け: 前日リマインド ──
    if (dateStr === oneDayLaterStr) {
      sendReminderEmailDayBefore(rowData);
      console.log(`前日リマインド送信: ${rowData.email}`);
    }
  }
}

/**
 * 相談担当者・オブザーバー向けリマインド送信
 * P列（担当者）＋ 日程設定シートの参加メンバー全員に送信
 */
function sendStaffReminderWithMembers_(rowData, daysBeforeLabel) {
  // 日程設定シートから参加メンバーを取得
  const memberList = getScheduleMembersForDate_(rowData.confirmedDate);

  const lineMsg = getStaffReminderLine(rowData, daysBeforeLabel, memberList);
  const emailSubject = `【${daysBeforeLabel}】${rowData.name}様（${rowData.company}） - ${rowData.confirmedDate}`;
  const emailBody = getStaffReminderEmail(rowData, daysBeforeLabel, memberList);

  var sentEmails = {};

  // P列の担当者に通知
  if (rowData.staff) {
    sendStaffNotifications(rowData.staff, lineMsg, emailSubject, emailBody);
    // 送信済みメールアドレスを記録（重複防止）
    var staffNames = rowData.staff.split(',').map(function(n) { return n.trim(); }).filter(function(n) { return n; });
    staffNames.forEach(function(name) {
      var member = getMemberByName(name);
      if (member && member.email) sentEmails[member.email] = true;
    });
  }

  // 日程設定シートの参加メンバー全員にリマインド送信（P列と重複しない分）
  if (memberList && memberList.length > 0) {
    memberList.forEach(function(m) {
      var member = getMemberByName(m.name);
      if (member && member.email && !sentEmails[member.email]) {
        GmailApp.sendEmail(member.email, emailSubject, emailBody, {
          name: CONFIG.SENDER_NAME
        });
        // LINE IDがあればLINEも送信
        if (member.lineId) {
          sendLineMessage(member.lineId, lineMsg);
        }
        sentEmails[member.email] = true;
        console.log('リマインド送信（参加メンバー）: ' + m.name + ' (' + member.email + ')');
      }
    });
  }

  // 誰にも送れなかった場合は管理者にフォールバック
  if (Object.keys(sentEmails).length === 0) {
    CONFIG.ADMIN_EMAILS.forEach(function(adminEmail) {
      GmailApp.sendEmail(adminEmail, emailSubject, emailBody, {
        name: CONFIG.SENDER_NAME
      });
    });
    console.log('リマインド: 担当者・メンバー未設定のため管理者にフォールバック');
  }
}

/**
 * 日程設定シートから指定日の参加メンバー情報を取得
 * @param {string} confirmedDate - 確定日時
 * @returns {Array<Object>} [{name, term, type}]
 */
function getScheduleMembersForDate_(confirmedDate) {
  if (!confirmedDate) return [];

  var dateStr;
  if (confirmedDate instanceof Date || (typeof confirmedDate === 'object' && typeof confirmedDate.getTime === 'function')) {
    dateStr = Utilities.formatDate(confirmedDate, 'Asia/Tokyo', 'yyyy/MM/dd');
  } else {
    var s = confirmedDate.toString();
    // "2026/02/27 HH:mm" or "2026-02-27" format
    var m1 = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m1) {
      dateStr = m1[1] + '/' + ('0' + m1[2]).slice(-2) + '/' + ('0' + m1[3]).slice(-2);
    } else {
      // Date.toString() format: "Fri Feb 27 2026 ..."
      var monthMap = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
      var m2 = s.match(/\w+\s+(\w+)\s+(\d{1,2})\s+(\d{4})/);
      if (m2 && monthMap[m2[1]]) {
        dateStr = m2[3] + '/' + monthMap[m2[1]] + '/' + ('0' + m2[2]).slice(-2);
      } else {
        dateStr = s.substring(0, 10).replace(/-/g, '/');
      }
    }
  }
  const schedSheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!schedSheet) return [];

  const schedData = schedSheet.getDataRange().getValues();
  for (let i = 1; i < schedData.length; i++) {
    const schedDate = schedData[i][SCHEDULE_COLUMNS.DATE];
    if (!schedDate) continue;
    const schedDateStr = Utilities.formatDate(new Date(schedDate), 'Asia/Tokyo', 'yyyy/MM/dd');
    if (schedDateStr === dateStr && schedData[i][SCHEDULE_COLUMNS.MEMBERS]) {
      return getMembersByNames(schedData[i][SCHEDULE_COLUMNS.MEMBERS].toString())
        .map(function(m) { return { name: m.name, term: m.term, type: m.type }; });
    }
  }
  return [];
}

/**
 * 3日前リマインドメール送信（予約者向け）
 */
function sendReminderEmail3DaysBefore(data) {
  const subject = '【3日後のご相談について】準備のご案内';
  const body = getReminderEmail3DaysBefore(data);

  GmailApp.sendEmail(data.email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });
}

/**
 * 前日リマインドメール送信（予約者向け）
 */
function sendReminderEmailDayBefore(data) {
  const subject = '【明日のご相談について】最終確認';
  const body = getReminderEmailDayBefore(data);

  GmailApp.sendEmail(data.email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });
}

/**
 * リマインドトリガーの設定（毎日午前9時に実行）
 */
function setupDailyReminderTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendDailyReminders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sendDailyReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  console.log('リマインドトリガーを設定しました（毎日9時）');
}

/**
 * すべてのトリガーを設定
 */
function setupAllTriggers() {
  setupOnEditTrigger();
  setupDailyReminderTrigger();
  setupFirstPollingTrigger();
  setupReminderPollingTrigger();
  setupFinalizeScheduleTrigger();
  setupReportDeadlineTrigger();
  setupCancellationEmailTrigger();
  console.log('すべてのトリガーを設定しました');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// キャンセルメール自動検知 & 通知
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * キャンセル処理の一元化
 * 日程解放、Zoom削除、担当者・オブザーバーへの通知を行う
 * @param {number} rowIndex - 行番号（1-based）
 * @param {Object} rowData - getRowData()の戻り値
 * @param {string} source - キャンセル元（'メール検知' or '手動変更'）
 */
function processCancellation(rowIndex, rowData, source) {
  // 日程設定シートの予約状況を「空き」に戻す
  if (rowData.confirmedDate) {
    var cancelParsed = parseConfirmedDateTime(rowData.confirmedDate);
    if (cancelParsed.date) {
      var freed = markAsAvailable(cancelParsed.date, cancelParsed.time);
      console.log('日程設定シート同期（キャンセル）: ' + cancelParsed.date + ' ' + (cancelParsed.time || '') + ' → ' + (freed ? '空き' : '該当なし'));
    }
  }

  // Zoomミーティングを削除
  if (rowData.zoomUrl) {
    var deleted = deleteZoomMeeting(rowData.zoomUrl);
    console.log('Zoomミーティング削除: ' + (deleted ? '成功' : '失敗またはスキップ'));
  }

  // LINE通知（グループ）
  sendLineStatusNotification(rowData, STATUS.CANCELLED);

  // 担当者・オブザーバーへの通知
  notifyCancellationToMembers(rowData);

  // 管理者への通知
  var adminSubject = '【キャンセル】' + (rowData.name || '') + '様（' + (rowData.company || '') + '） - ' + source;
  var adminBody = getCancellationNotificationEmail(rowData, '管理者');
  CONFIG.ADMIN_EMAILS.forEach(function(email) {
    GmailApp.sendEmail(email, adminSubject, adminBody, { name: CONFIG.SENDER_NAME });
  });

  // 相談者にキャンセル受領確認メール送信
  if (rowData.email) {
    var confirmBody = getCancellationConfirmEmail(rowData);
    GmailApp.sendEmail(rowData.email,
      '【キャンセル受付完了】無料経営相談のキャンセルを承りました',
      confirmBody,
      { name: CONFIG.SENDER_NAME, replyTo: CONFIG.REPLY_TO }
    );
    console.log('キャンセル確認メール送信: ' + rowData.email);
  }

  console.log('キャンセル処理完了: ' + rowData.id + ' (' + source + ')');
}

/**
 * 担当者・オブザーバーへキャンセル通知を送信
 * 日程設定シートの参加メンバー全員に通知
 * @param {Object} rowData - 予約データ
 */
function notifyCancellationToMembers(rowData) {
  // 日程設定シートから参加メンバーを取得
  var memberList = getScheduleMembersForDate_(rowData.confirmedDate);
  var notifiedEmails = {};

  // リーダーがいればメンバーリストに含まれていなくても通知
  if (rowData.leader) {
    var leaderMember = getMemberByName(rowData.leader);
    if (leaderMember && leaderMember.email) {
      var leaderBody = getCancellationNotificationEmail(rowData, rowData.leader);
      GmailApp.sendEmail(leaderMember.email,
        '【キャンセル】' + (rowData.company || '') + '様 - 経営相談キャンセルのお知らせ',
        leaderBody,
        { name: CONFIG.SENDER_NAME }
      );
      notifiedEmails[leaderMember.email] = true;
      console.log('キャンセル通知（リーダー）: ' + rowData.leader + ' (' + leaderMember.email + ')');
    }
  }

  // 日程設定シートの参加メンバー全員に通知
  if (memberList && memberList.length > 0) {
    memberList.forEach(function(m) {
      var member = getMemberByName(m.name);
      if (member && member.email && !notifiedEmails[member.email]) {
        var body = getCancellationNotificationEmail(rowData, m.name);
        GmailApp.sendEmail(member.email,
          '【キャンセル】' + (rowData.company || '') + '様 - 経営相談キャンセルのお知らせ',
          body,
          { name: CONFIG.SENDER_NAME }
        );
        notifiedEmails[member.email] = true;
        console.log('キャンセル通知（メンバー）: ' + m.name + ' (' + member.email + ')');
      }
    });
  }

  // 担当者（P列）にも通知（リーダーやメンバーと重複しない場合のみ）
  if (rowData.staff) {
    var staffNames = rowData.staff.split(',').map(function(n) { return n.trim(); }).filter(function(n) { return n; });
    staffNames.forEach(function(name) {
      var member = getMemberByName(name);
      if (member && member.email && !notifiedEmails[member.email]) {
        var body = getCancellationNotificationEmail(rowData, name);
        GmailApp.sendEmail(member.email,
          '【キャンセル】' + (rowData.company || '') + '様 - 経営相談キャンセルのお知らせ',
          body,
          { name: CONFIG.SENDER_NAME }
        );
        notifiedEmails[member.email] = true;
        console.log('キャンセル通知（担当者）: ' + name + ' (' + member.email + ')');
      }
    });
  }

  // LINE通知も各メンバーへ
  var lineMsg = '❌ 【キャンセル】\n\n' +
    '相談者: ' + (rowData.name || '') + '様\n' +
    '企業名: ' + (rowData.company || '') + '\n' +
    '日時: ' + (rowData.confirmedDate || '') + '\n' +
    'テーマ: ' + (rowData.theme || '') + '\n\n' +
    '当該日程の準備は不要です。';

  if (rowData.staff) {
    sendStaffNotifications(rowData.staff, lineMsg,
      '【キャンセル】' + (rowData.company || '') + '様',
      getCancellationNotificationEmail(rowData, rowData.staff)
    );
  }
}

/**
 * Gmailの受信メールからキャンセルを自動検知
 * 10分おきにトリガーで実行
 */
function checkCancellationEmails() {
  var labelName = 'キャンセル処理済';
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }

  // 過去24時間以内のキャンセル関連メール（未処理のみ）
  var query = 'is:unread newer_than:1d (キャンセル OR 取消 OR 取り消し OR 中止) -label:' + labelName;
  var threads = GmailApp.search(query, 0, 20);

  if (threads.length === 0) return;

  // 予約管理シートからアクティブな予約を取得
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  // メールアドレスをキーにした予約マップ（アクティブなもののみ）
  var activeBookings = {};
  var activeStatuses = [STATUS.PENDING, STATUS.NDA_AGREED, STATUS.RECEIVED, STATUS.CONFIRMED];
  for (var i = 1; i < data.length; i++) {
    var email = (data[i][COLUMNS.EMAIL] || '').toString().toLowerCase().trim();
    var status = data[i][COLUMNS.STATUS];
    if (email && activeStatuses.indexOf(status) >= 0) {
      if (!activeBookings[email]) activeBookings[email] = [];
      activeBookings[email].push({
        rowIndex: i + 1,
        id: data[i][COLUMNS.ID],
        name: data[i][COLUMNS.NAME],
        company: data[i][COLUMNS.COMPANY],
        status: status,
        confirmedDate: data[i][COLUMNS.CONFIRMED_DATE]
      });
    }
  }

  // 各スレッドを処理
  for (var t = 0; t < threads.length; t++) {
    var thread = threads[t];
    var messages = thread.getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      if (message.isStarred()) continue; // スター付きはスキップ（管理者が手動対応中）

      var senderEmail = extractEmailAddress(message.getFrom()).toLowerCase().trim();
      var bookings = activeBookings[senderEmail];

      if (!bookings || bookings.length === 0) continue;

      // 最新の予約を対象にキャンセル処理
      var booking = bookings[bookings.length - 1];
      var rowData = getRowData(booking.rowIndex);

      // ステータスをキャンセルに更新
      sheet.getRange(booking.rowIndex, COLUMNS.STATUS + 1).setValue(STATUS.CANCELLED);

      // キャンセル処理実行（相談者への確認メールもprocessCancellation内で送信）
      processCancellation(booking.rowIndex, rowData, 'メール検知');

      console.log('キャンセル自動検知: ' + booking.id + ' (' + senderEmail + ')');
    }

    // ラベルを付けて処理済みにする
    thread.addLabel(label);
    thread.markRead();
  }
}

/**
 * メールアドレスを抽出（"名前 <email>" → "email"）
 * @param {string} from - Fromヘッダー
 * @returns {string} メールアドレス
 */
function extractEmailAddress(from) {
  if (!from) return '';
  var match = from.match(/<([^>]+)>/);
  if (match) return match[1];
  return from.trim();
}

/**
 * キャンセルメール検知トリガーのセットアップ（10分おき）
 */
function setupCancellationEmailTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkCancellationEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkCancellationEmails')
    .timeBased()
    .everyMinutes(10)
    .create();

  console.log('キャンセルメール検知トリガーをセットアップしました（10分おき）');
}

/**
 * 指定行の予約について、日程設定シートの参加メンバー全員に確定通知を送信
 * ScriptPropertiesの PENDING_NOTIFY_ROW に行番号をセットして呼び出す
 */
function sendScheduledStaffNotification() {
  var props = PropertiesService.getScriptProperties();
  var rowStr = props.getProperty('PENDING_NOTIFY_ROW');
  if (!rowStr) {
    console.log('PENDING_NOTIFY_ROW が未設定です');
    return;
  }

  var rowIndex = parseInt(rowStr);
  props.deleteProperty('PENDING_NOTIFY_ROW');

  // 一回限りのトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendScheduledStaffNotification') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  var data = getRowData(rowIndex);
  if (!data.id) {
    console.log('行 ' + rowIndex + ' にデータがありません');
    return;
  }

  var emailResult = buildStaffNotificationEmail_(data);
  var lineMsg = buildStaffNotificationLine_(data);
  var senderName = emailResult.senderName || CONFIG.SENDER_NAME;

  var sentEmails = {};

  // P列の担当者
  if (data.staff) {
    sendStaffNotifications(data.staff, lineMsg, emailResult.subject, emailResult.body);
    var staffNames = data.staff.split(',').map(function(n) { return n.trim(); }).filter(function(n) { return n; });
    staffNames.forEach(function(name) {
      var m = getMemberByName(name);
      if (m && m.email) sentEmails[m.email] = true;
    });
  }

  // 日程設定シートの参加メンバー
  var memberList = getScheduleMembersForDate_(data.confirmedDate);
  if (memberList && memberList.length > 0) {
    memberList.forEach(function(cm) {
      var m = getMemberByName(cm.name);
      if (m && m.email && !sentEmails[m.email]) {
        GmailApp.sendEmail(m.email, emailResult.subject, emailResult.body, { name: senderName });
        if (m.lineId) sendLineMessage(m.lineId, lineMsg);
        sentEmails[m.email] = true;
        console.log('確定通知送信: ' + cm.name + ' (' + m.email + ')');
      }
    });
  }

  console.log('スケジュール確定通知完了: 行' + rowIndex + ', 送信数: ' + Object.keys(sentEmails).length);
}

/**
 * 相談担当者向け通知メールの件名・本文を構築
 * @param {Object} data - getRowData形式のデータ
 * @returns {Object} { subject, body }
 */
function buildStaffNotificationEmail_(data) {
  var method = data.method || '';
  var isOnline = method.indexOf('オンライン') >= 0 || method.indexOf('Zoom') >= 0;

  var subject = '関学無料経営相談依頼案件（' + data.confirmedDate + '）';

  var body = '下記の経営相談が確定しましたのでお知らせいたします。\n' +
    '事前準備をお願いいたします。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 相談概要\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '申込ID　：' + data.id + '\n' +
    '相談日時：' + data.confirmedDate + '\n' +
    '相談方法：' + method + '\n' +
    (data.location ? '場所　　：' + data.location + '\n' : '') +
    (isOnline && data.zoomUrl ? 'Zoom URL：' + data.zoomUrl + '\n' : '') +
    '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 相談者情報\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'お名前　：' + data.name + ' 様\n' +
    '貴社名　：' + data.company + '\n' +
    '業種　　：' + (data.industry || '未入力') + '\n' +
    '役職　　：' + (data.position || '未入力') + '\n' +
    'メール　：' + data.email + '\n' +
    '電話番号：' + (data.phone || '未入力') + '\n' +
    (data.companyUrl ? '企業URL ：' + data.companyUrl + '\n' : '') +
    '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 相談内容\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'テーマ　：' + (data.theme || '未入力') + '\n' +
    '内容　　：\n' + (data.content || '未入力') + '\n' +
    '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 担当\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'リーダー：' + (data.leader || '未定') + '\n' +
    '参加メンバー：' + (getParticipatingMembers(data.confirmedDate) || '未定') + '\n' +
    '\n' +
    (data.companyUrl ? '※事前に企業URLを確認し、リサーチをお願いします。\n\n' : '') +
    'よろしくお願いいたします。\n' +
    '中小企業経営診断研究会 無料経営相談分科会';

  return { subject: subject, body: body, senderName: '中小企業経営診断研究会 無料経営相談分科会' };
}

/**
 * 相談担当者向けLINE通知メッセージを構築
 * @param {Object} data - getRowData形式のデータ
 * @returns {string} LINE通知テキスト
 */
function buildStaffNotificationLine_(data) {
  var method = data.method || '';
  var isOnline = method.indexOf('オンライン') >= 0 || method.indexOf('Zoom') >= 0;

  return '📋 経営相談依頼案件\n\n' +
    '日時: ' + data.confirmedDate + '\n' +
    '方法: ' + method + '\n' +
    (isOnline && data.zoomUrl ? 'Zoom: ' + data.zoomUrl + '\n' : '') +
    '\n' +
    data.name + ' 様（' + data.company + '）\n' +
    '業種: ' + (data.industry || '-') + '\n' +
    'テーマ: ' + (data.theme || '-') + '\n' +
    (data.companyUrl ? '企業URL: ' + data.companyUrl + '\n' : '') +
    'リーダー: ' + (data.leader || '未定');
}
