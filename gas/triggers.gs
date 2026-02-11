/**
 * トリガー処理（拡張版）
 * リマインド2日前/3日前、担当者個別通知対応
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
    const bookable = getBookableStatus(score, specialFlag);

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
    // 確定日時が入力されているか確認
    if (!data.confirmedDate) {
      SpreadsheetApp.getUi().alert(
        '確定日時が入力されていません。\n確定日時（P列）を入力してからステータスを変更してください。'
      );
      const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
        .getSheetByName(CONFIG.SHEET_NAME);
      sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(oldStatus || STATUS.PENDING);
      return;
    }

    // 確定メールを送信
    sendConfirmedEmail(data);
    console.log(`確定メール送信完了: ${data.email}`);

    // 担当者への通知（企業URL情報を含む）
    if (data.staff) {
      const staffLineMsg = `✅ 予約確定

申込ID: ${data.id}
お名前: ${data.name}様
貴社名: ${data.company}
日時: ${data.confirmedDate}
方法: ${data.method}
テーマ: ${data.theme}
${data.companyUrl ? '企業URL: ' + data.companyUrl + '\n※事前リサーチをお願いします' : ''}`;

      const staffEmailBody = `予約が確定しました。

申込ID：${data.id}
お名前：${data.name}様
貴社名：${data.company}
日時：${data.confirmedDate}
相談方法：${data.method}
テーマ：${data.theme}
${data.companyUrl ? '\n企業URL：' + data.companyUrl + '\n※事前リサーチにAIツールの活用を推奨します' : ''}

事前準備をお願いいたします。`;

      sendStaffNotifications(
        data.staff,
        staffLineMsg,
        `【予約確定】${data.name}様 - ${data.confirmedDate}`,
        staffEmailBody
      );
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
    sendLineStatusNotification(data, newStatus);
  }

  // キャンセルに変更された場合
  if (newStatus === STATUS.CANCELLED) {
    sendLineStatusNotification(data, newStatus);
  }
}

/**
 * 毎日のリマインド送信（2日前 + 3日前）
 */
function sendDailyReminders() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  const now = new Date();

  // 2日後の日付
  const twoDaysLater = new Date(now);
  twoDaysLater.setDate(twoDaysLater.getDate() + 2);
  const twoDaysLaterStr = Utilities.formatDate(twoDaysLater, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 3日後の日付
  const threeDaysLater = new Date(now);
  threeDaysLater.setDate(threeDaysLater.getDate() + 3);
  const threeDaysLaterStr = Utilities.formatDate(threeDaysLater, 'Asia/Tokyo', 'yyyy/MM/dd');

  for (let i = 1; i < data.length; i++) {
    const status = data[i][COLUMNS.STATUS];
    const confirmedDate = data[i][COLUMNS.CONFIRMED_DATE];

    if (status !== STATUS.CONFIRMED || !confirmedDate) {
      continue;
    }

    const dateStr = confirmedDate.toString().substring(0, 10).replace(/-/g, '/');
    const rowData = getRowData(i + 1);

    // 3日前リマインド
    if (dateStr === threeDaysLaterStr || threeDaysLaterStr === dateStr) {
      // 予約者向け（メール）
      sendReminderEmail3DaysBefore(rowData);

      // 担当者向け（LINE優先 → メールフォールバック）
      if (rowData.staff) {
        const lineMsg = getStaffReminderLine(rowData, '3日前');
        const emailSubject = `【3日前リマインド】${rowData.name}様 - ${rowData.confirmedDate}`;
        const emailBody = getStaffReminderEmail(rowData, '3日前');
        sendStaffNotifications(rowData.staff, lineMsg, emailSubject, emailBody);
      }

      console.log(`3日前リマインド送信: ${rowData.email}`);
    }

    // 2日前リマインド
    if (dateStr === twoDaysLaterStr || twoDaysLaterStr === dateStr) {
      // 予約者向け（メール）
      sendReminderEmail2DaysBefore(rowData);

      // 担当者向け（LINE優先 → メールフォールバック）
      if (rowData.staff) {
        const lineMsg = getStaffReminderLine(rowData, '2日前');
        const emailSubject = `【2日前・最終確認】${rowData.name}様 - ${rowData.confirmedDate}`;
        const emailBody = getStaffReminderEmail(rowData, '2日前');
        sendStaffNotifications(rowData.staff, lineMsg, emailSubject, emailBody);
      }

      console.log(`2日前リマインド送信: ${rowData.email}`);
    }
  }
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
 * 2日前リマインドメール送信（予約者向け）
 */
function sendReminderEmail2DaysBefore(data) {
  const subject = '【明後日のご相談について】最終確認';
  const body = getReminderEmail2DaysBefore(data);

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
  console.log('すべてのトリガーを設定しました');
}
