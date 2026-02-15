/**
 * Googleフォームによるメンバー参加可能日程アンケートシステム v2
 * - オプトアウト方式（デフォルト全参加、不可日をチェック）
 * - 2回送信（1回目: 25日にオプトアウト / 2回目: 翌月15日に確認+追加）
 * - 予約済み枠の保護（離脱時に管理者通知）
 * - 条件未達枠の自動クローズ（21日に確定）
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// フォーム作成
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 日程設定シートから対象月の日程候補を読み取る
 * @param {number} year - 対象年
 * @param {number} month - 対象月（1-12）
 * @returns {Object} { dateKeys: [...], scheduleByDate: {...} }
 */
function getScheduleCandidates(year, month) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!sheet) {
    throw new Error('日程設定シートが見つかりません');
  }

  const data = sheet.getDataRange().getValues();
  const scheduleByDate = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[SCHEDULE_COLUMNS.DATE];
    if (!dateVal) continue;

    const date = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (date.getFullYear() !== year || (date.getMonth() + 1) !== month) continue;

    const dateKey = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
    const dow = date.getDay();
    const dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    const dateLabel = month + '/' + date.getDate() + '(' + dayNames[dow] + ')';

    // 金曜日にはZoomのみ表記を追加
    const isFriday = (dow === 5);
    const nthWeek = getNthDayOfWeek(year, month - 1, date.getDate(), dow);
    let suffix = '';
    if (isFriday) {
      suffix = ' [第' + nthWeek + '金曜・Zoomのみ]';
    }

    let timeStr;
    const time = row[SCHEDULE_COLUMNS.TIME];
    if (time instanceof Date) {
      timeStr = Utilities.formatDate(time, 'Asia/Tokyo', 'HH:mm');
    } else {
      timeStr = String(time);
    }

    if (!scheduleByDate[dateKey]) {
      scheduleByDate[dateKey] = {
        label: dateLabel + suffix,
        times: []
      };
    }
    scheduleByDate[dateKey].times.push(timeStr);
  }

  const dateKeys = Object.keys(scheduleByDate).sort();
  if (dateKeys.length === 0) {
    throw new Error(year + '年' + month + '月の日程候補が見つかりません');
  }

  return { dateKeys: dateKeys, scheduleByDate: scheduleByDate };
}

/**
 * 1回目: オプトアウトフォームを作成
 * メンバーは参加できない日時にチェックを入れる
 * @param {number} year - 対象年
 * @param {number} month - 対象月（1-12）
 * @returns {GoogleAppsScript.Forms.Form} 作成されたフォーム
 */
function createOptOutForm(year, month) {
  const candidates = getScheduleCandidates(year, month);

  const title = year + '年' + month + '月 参加不可日程の申告';
  const form = FormApp.create(title);
  form.setDescription(
    'デフォルトでは全日程に参加可能として登録されます。\n' +
    '参加できない日時がある場合のみ、該当箇所にチェックを入れてください。\n' +
    '全日程OKの方は回答不要です。\n' +
    '再送信で修正可能です。\n\n' +
    '回答期限: ' + CONFIG.FORM_POLLING.FIRST_DEADLINE_DAYS + '日以内'
  );
  form.setCollectEmail(true);
  form.setAllowResponseEdits(true);
  form.setLimitOneResponsePerUser(true);

  // 日付ごとにチェックボックス質問を生成
  candidates.dateKeys.forEach(function(dateKey) {
    var entry = candidates.scheduleByDate[dateKey];
    var item = form.addCheckboxItem();
    item.setTitle(entry.label);
    item.setChoices(entry.times.map(function(t) {
      return item.createChoice(t);
    }));
    item.setRequired(false);
  });

  // フォーム送信トリガーを設定
  setupFormSubmitTrigger(form.getId(), 'processFormResponse');

  console.log('オプトアウトフォーム作成完了: ' + title);
  console.log('URL: ' + form.getPublishedUrl());

  return form;
}

/**
 * 2回目: 確認+追加フォームを作成
 * セクション分岐: OK → 送信 / 追加あり → チェックボックス
 * @param {number} year - 対象年
 * @param {number} month - 対象月（1-12）
 * @returns {GoogleAppsScript.Forms.Form} 作成されたフォーム
 */
function createConfirmationForm(year, month) {
  const candidates = getScheduleCandidates(year, month);

  const title = year + '年' + month + '月 日程確認・追加フォーム';
  const form = FormApp.create(title);
  form.setDescription(
    '1回目アンケートの結果を踏まえた最終確認です。\n' +
    '変更がなければ「現在の内容でOK」を選んで送信してください。\n' +
    '追加で参加可能になった日がある場合はチェックしてください。\n\n' +
    '最終期限: ' + CONFIG.FORM_POLLING.FINAL_DEADLINE_DAYS + '日以内'
  );
  form.setCollectEmail(true);
  form.setAllowResponseEdits(true);
  form.setLimitOneResponsePerUser(true);

  // セクション1: 確認質問
  var confirmItem = form.addMultipleChoiceItem();
  confirmItem.setTitle('現在の登録状況を確認してください');
  confirmItem.setRequired(true);

  // セクション2: 追加参加日の申告（ページブレーク付き）
  var addSection = form.addPageBreakItem();
  addSection.setTitle('追加参加日の申告');
  addSection.setHelpText('追加で参加可能になった日時にチェックしてください。');

  // 送信セクション（最終ページ）
  var submitSection = form.addPageBreakItem();
  submitSection.setTitle('送信');

  // 確認質問の選択肢にセクション遷移を設定
  confirmItem.setChoices([
    confirmItem.createChoice('現在の内容でOK（変更なし）', submitSection),
    confirmItem.createChoice('追加で参加可能な日がある', addSection)
  ]);

  // 追加セクションに日付ごとのチェックボックスを生成
  candidates.dateKeys.forEach(function(dateKey) {
    var entry = candidates.scheduleByDate[dateKey];
    var item = form.addCheckboxItem();
    item.setTitle(entry.label);
    item.setChoices(entry.times.map(function(t) {
      return item.createChoice(t);
    }));
    item.setRequired(false);
  });

  // フォーム送信トリガーを設定
  setupFormSubmitTrigger(form.getId(), 'processConfirmationResponse');

  console.log('確認+追加フォーム作成完了: ' + title);
  console.log('URL: ' + form.getPublishedUrl());

  return form;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// メンバー一括登録
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 全メンバーを対象月の全枠にデフォルト登録する
 * オプトアウト方式のため、初期状態は全員参加可能
 * @param {number} year - 対象年
 * @param {number} month - 対象月（1-12）
 */
function registerAllMembersForMonth(year, month) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sheet) return;

  const members = getScheduleMembers();
  const memberNames = members.map(function(m) { return m.name; }).filter(function(n) { return n; });
  const allMembersStr = memberNames.join(', ');

  const data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
    if (!dateVal) continue;

    var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (date.getFullYear() !== year || (date.getMonth() + 1) !== month) continue;

    // 全メンバーを登録
    sheet.getRange(i + 1, SCHEDULE_COLUMNS.MEMBERS + 1).setValue(allMembersStr);

    // スコア再計算
    recalculateScheduleScoreForRow(i + 1, sheet);
  }

  console.log(year + '年' + month + '月: 全メンバー(' + memberNames.length + '名)をデフォルト登録');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// フォーム回答処理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 1回目フォーム送信時のトリガーハンドラ（オプトアウト方式）
 * メールでメンバー特定 → 全枠に登録 → 不可日を除外 → スコア再計算
 * @param {Object} e - フォーム送信イベント
 */
function processFormResponse(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    console.log('ロック取得失敗: ' + err.message);
    return;
  }

  try {
    var response = e.response;
    var email = response.getRespondentEmail();

    // メールアドレスからメンバー特定
    var member = getMemberByEmail(email);
    if (!member) {
      console.log('メンバー不明: ' + email);
      return;
    }

    var memberName = member.name;
    console.log('オプトアウトフォーム回答処理: ' + memberName + ' (' + email + ')');

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();

    // フォームの質問から対象日付を取得
    var form = FormApp.openById(e.source.getId());
    var items = form.getItems(FormApp.ItemType.CHECKBOX);

    var formDateLabels = {};
    items.forEach(function(item) {
      formDateLabels[item.getTitle()] = true;
    });

    // 回答内容をパース: 不可日時 { "4/5(土)": ["14:00", "16:00"], ... }
    var unavailable = {};
    var itemResponses = response.getItemResponses();
    itemResponses.forEach(function(ir) {
      var title = ir.getItem().getTitle();
      var value = ir.getResponse();
      if (value && value.length > 0) {
        unavailable[title] = value;
      }
    });

    var changedRows = {};
    var warnings = [];

    // Phase 1: 対象月の全行からそのメンバーをクリア
    for (var i = 1; i < data.length; i++) {
      var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
      if (!dateVal) continue;

      var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
      var dow = date.getDay();
      var dayNames = ['日', '月', '火', '水', '木', '金', '土'];
      var dateLabel = (date.getMonth() + 1) + '/' + date.getDate() + '(' + dayNames[dow] + ')';

      // 金曜日にはZoomのみ表記を追加
      if (dow === 5) {
        var nthWeek = getNthDayOfWeek(date.getFullYear(), date.getMonth(), date.getDate(), dow);
        dateLabel = dateLabel + ' [第' + nthWeek + '金曜・Zoomのみ]';
      }

      if (!formDateLabels[dateLabel]) continue;

      var currentMembers = data[i][SCHEDULE_COLUMNS.MEMBERS]
        ? data[i][SCHEDULE_COLUMNS.MEMBERS].toString()
        : '';

      if (currentMembers.indexOf(memberName) !== -1) {
        var memberList = currentMembers.split(',').map(function(n) { return n.trim(); });
        var filtered = memberList.filter(function(n) { return n !== memberName && n !== ''; });
        var newValue = filtered.join(', ');
        sheet.getRange(i + 1, SCHEDULE_COLUMNS.MEMBERS + 1).setValue(newValue);
        data[i][SCHEDULE_COLUMNS.MEMBERS] = newValue;
        changedRows[i + 1] = true;
      }
    }

    // Phase 2: 対象月の全行にメンバーを追加（デフォルト = 全参加）
    for (var i = 1; i < data.length; i++) {
      var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
      if (!dateVal) continue;

      var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
      var dow = date.getDay();
      var dayNames = ['日', '月', '火', '水', '木', '金', '土'];
      var dateLabel = (date.getMonth() + 1) + '/' + date.getDate() + '(' + dayNames[dow] + ')';

      if (dow === 5) {
        var nthWeek = getNthDayOfWeek(date.getFullYear(), date.getMonth(), date.getDate(), dow);
        dateLabel = dateLabel + ' [第' + nthWeek + '金曜・Zoomのみ]';
      }

      if (!formDateLabels[dateLabel]) continue;

      var currentMembers = data[i][SCHEDULE_COLUMNS.MEMBERS]
        ? data[i][SCHEDULE_COLUMNS.MEMBERS].toString()
        : '';

      // 既にいなければ追加
      if (currentMembers.indexOf(memberName) === -1) {
        var newValue = currentMembers ? currentMembers + ', ' + memberName : memberName;
        sheet.getRange(i + 1, SCHEDULE_COLUMNS.MEMBERS + 1).setValue(newValue);
        data[i][SCHEDULE_COLUMNS.MEMBERS] = newValue;
        changedRows[i + 1] = true;
      }
    }

    // Phase 3: チェックされた（不可）日時の行からメンバーを削除
    for (var i = 1; i < data.length; i++) {
      var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
      if (!dateVal) continue;

      var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
      var dow = date.getDay();
      var dayNames = ['日', '月', '火', '水', '木', '金', '土'];
      var dateLabel = (date.getMonth() + 1) + '/' + date.getDate() + '(' + dayNames[dow] + ')';

      if (dow === 5) {
        var nthWeek = getNthDayOfWeek(date.getFullYear(), date.getMonth(), date.getDate(), dow);
        dateLabel = dateLabel + ' [第' + nthWeek + '金曜・Zoomのみ]';
      }

      if (!unavailable[dateLabel]) continue;

      var timeStr;
      var time = data[i][SCHEDULE_COLUMNS.TIME];
      if (time instanceof Date) {
        timeStr = Utilities.formatDate(time, 'Asia/Tokyo', 'HH:mm');
      } else {
        timeStr = String(time);
      }

      if (unavailable[dateLabel].indexOf(timeStr) === -1) continue;

      var currentMembers = data[i][SCHEDULE_COLUMNS.MEMBERS]
        ? data[i][SCHEDULE_COLUMNS.MEMBERS].toString()
        : '';

      if (currentMembers.indexOf(memberName) !== -1) {
        // 予約済み枠かチェック
        var bookingStatus = data[i][SCHEDULE_COLUMNS.BOOKING_STATUS];
        if (bookingStatus === '予約済み') {
          var dateStr = dateVal instanceof Date
            ? Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd')
            : String(dateVal);
          warnings.push({
            date: dateStr,
            time: timeStr,
            memberName: memberName
          });
        }

        var memberList = currentMembers.split(',').map(function(n) { return n.trim(); });
        var filtered = memberList.filter(function(n) { return n !== memberName && n !== ''; });
        var newValue = filtered.join(', ');
        sheet.getRange(i + 1, SCHEDULE_COLUMNS.MEMBERS + 1).setValue(newValue);
        data[i][SCHEDULE_COLUMNS.MEMBERS] = newValue;
        changedRows[i + 1] = true;
      }
    }

    // Phase 4: 予約済み枠から離脱した場合、管理者に警告
    if (warnings.length > 0) {
      sendBookingWarningEmail(warnings, sheet, data);
    }

    // Phase 5: 変更された行のスコア再計算
    Object.keys(changedRows).forEach(function(rowIndex) {
      recalculateScheduleScoreForRow(parseInt(rowIndex), sheet);
    });

    console.log('オプトアウト回答処理完了: ' + memberName + ' - ' + Object.keys(changedRows).length + '行更新');

  } finally {
    lock.releaseLock();
  }
}

/**
 * 2回目フォーム送信時のトリガーハンドラ（確認+追加）
 * @param {Object} e - フォーム送信イベント
 */
function processConfirmationResponse(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    console.log('ロック取得失敗: ' + err.message);
    return;
  }

  try {
    var response = e.response;
    var email = response.getRespondentEmail();

    var member = getMemberByEmail(email);
    if (!member) {
      console.log('メンバー不明: ' + email);
      return;
    }

    var memberName = member.name;
    console.log('確認フォーム回答処理: ' + memberName + ' (' + email + ')');

    var itemResponses = response.getItemResponses();
    if (itemResponses.length === 0) return;

    // 最初の回答が確認質問
    var confirmAnswer = itemResponses[0].getResponse();

    if (confirmAnswer === '現在の内容でOK（変更なし）') {
      console.log(memberName + ': 現状維持で確定');
      return;
    }

    // 追加参加日がある場合
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var changedRows = {};

    // チェックボックス回答をパース（2番目以降の回答）
    var additions = {};
    for (var r = 1; r < itemResponses.length; r++) {
      var title = itemResponses[r].getItem().getTitle();
      var value = itemResponses[r].getResponse();
      if (value && value.length > 0) {
        additions[title] = value;
      }
    }

    // チェックされた日時の行にメンバーを追加（追加のみ、削除なし）
    for (var i = 1; i < data.length; i++) {
      var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
      if (!dateVal) continue;

      var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
      var dow = date.getDay();
      var dayNames = ['日', '月', '火', '水', '木', '金', '土'];
      var dateLabel = (date.getMonth() + 1) + '/' + date.getDate() + '(' + dayNames[dow] + ')';

      if (dow === 5) {
        var nthWeek = getNthDayOfWeek(date.getFullYear(), date.getMonth(), date.getDate(), dow);
        dateLabel = dateLabel + ' [第' + nthWeek + '金曜・Zoomのみ]';
      }

      if (!additions[dateLabel]) continue;

      var timeStr;
      var time = data[i][SCHEDULE_COLUMNS.TIME];
      if (time instanceof Date) {
        timeStr = Utilities.formatDate(time, 'Asia/Tokyo', 'HH:mm');
      } else {
        timeStr = String(time);
      }

      if (additions[dateLabel].indexOf(timeStr) === -1) continue;

      var currentMembers = data[i][SCHEDULE_COLUMNS.MEMBERS]
        ? data[i][SCHEDULE_COLUMNS.MEMBERS].toString()
        : '';

      // 既にいる場合はスキップ
      if (currentMembers.indexOf(memberName) !== -1) continue;

      var newValue = currentMembers ? currentMembers + ', ' + memberName : memberName;
      sheet.getRange(i + 1, SCHEDULE_COLUMNS.MEMBERS + 1).setValue(newValue);
      data[i][SCHEDULE_COLUMNS.MEMBERS] = newValue;
      changedRows[i + 1] = true;
    }

    // 変更された行のスコア再計算
    Object.keys(changedRows).forEach(function(rowIndex) {
      recalculateScheduleScoreForRow(parseInt(rowIndex), sheet);
    });

    console.log('確認フォーム処理完了: ' + memberName + ' - ' + Object.keys(changedRows).length + '行追加');

  } finally {
    lock.releaseLock();
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 予約保護: 管理者警告
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 予約済み枠からメンバーが離脱した場合、管理者に警告メール
 * @param {Array} warnings - 警告情報の配列
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 日程設定シート
 * @param {Array} data - シートデータ
 */
function sendBookingWarningEmail(warnings, sheet, data) {
  var body = '以下の予約済み枠から、メンバーが参加不可の申告をしました。\n' +
    '代替メンバーの確保が必要な場合があります。\n\n';

  warnings.forEach(function(w) {
    body += '━━━━━━━━━━━━━━━━━━━━\n';
    body += '対象枠: ' + w.date + ' ' + w.time + '\n';
    body += '離脱メンバー: ' + w.memberName + '\n';

    // 現在の残りメンバーを取得
    for (var i = 1; i < data.length; i++) {
      var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
      if (!dateVal) continue;
      var dateStr = dateVal instanceof Date
        ? Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(dateVal);
      var timeStr;
      var time = data[i][SCHEDULE_COLUMNS.TIME];
      if (time instanceof Date) {
        timeStr = Utilities.formatDate(time, 'Asia/Tokyo', 'HH:mm');
      } else {
        timeStr = String(time);
      }
      if (dateStr === w.date && timeStr === w.time) {
        var remaining = data[i][SCHEDULE_COLUMNS.MEMBERS] || '';
        var score = data[i][SCHEDULE_COLUMNS.SCORE] || 0;
        var bookable = data[i][SCHEDULE_COLUMNS.BOOKABLE] || '×';
        body += '残りメンバー: ' + (remaining || 'なし') + '\n';
        body += '現在の配置点数: ' + score + '\n';
        body += '予約可能判定: ' + bookable + '\n';
        break;
      }
    }
    body += '\n';
  });

  body += '※ 配置点数が不足する場合、LPから予約枠が非表示になります。';

  var subject = '【要確認】予約済み枠からのメンバー離脱';

  CONFIG.ADMIN_EMAILS.forEach(function(email) {
    GmailApp.sendEmail(email, subject, body, {
      name: CONFIG.SENDER_NAME
    });
  });

  console.log('予約済み枠警告メール送信: ' + warnings.length + '件');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// メール送信
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 1回目: 全メンバーにオプトアウトフォームURLをメール送信
 * @param {string} formUrl - フォームURL
 * @param {number} year - 対象年
 * @param {number} month - 対象月
 */
function sendOptOutFormToMembers(formUrl, year, month) {
  var members = getScheduleMembers();
  var sentCount = 0;

  var deadline = new Date();
  deadline.setDate(deadline.getDate() + CONFIG.FORM_POLLING.FIRST_DEADLINE_DAYS);
  var deadlineStr = Utilities.formatDate(deadline, 'Asia/Tokyo', 'yyyy/MM/dd');

  var subject = '【日程アンケート】' + year + '年' + month + '月 参加不可日の申告をお願いします';

  members.forEach(function(member) {
    if (!member.email) {
      console.log('メールアドレス未設定: ' + member.name);
      return;
    }

    var body = member.name + ' 様\n\n' +
      CONFIG.ORG.NAME + 'の日程調整アンケートです。\n\n' +
      year + '年' + month + '月の相談対応日程について、参加できない日時を申告してください。\n' +
      '全日程に参加可能な方は回答不要です（自動的に全枠参加として登録されます）。\n\n' +
      '▼ 回答フォーム\n' + formUrl + '\n\n' +
      '回答期限: ' + deadlineStr + '\n' +
      '※ 回答後も期限内であれば再送信で修正可能です。\n' +
      '※ ' + (month - 1 > 0 ? month - 1 : 12) + '月15日に最終確認のご連絡をいたします。\n\n' +
      'よろしくお願いいたします。\n\n' +
      CONFIG.ORG.NAME;

    GmailApp.sendEmail(member.email, subject, body, {
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.REPLY_TO
    });

    sentCount++;
    console.log('1回目メール送信完了: ' + member.name + ' (' + member.email + ')');
  });

  console.log('1回目アンケートメール送信完了: ' + sentCount + '名に送信');
}

/**
 * 2回目: 確認+追加フォームURLをメール送信
 * 1回目未回答者と回答済み者で文面を分ける
 * @param {string} formUrl - フォームURL
 * @param {number} year - 対象年
 * @param {number} month - 対象月
 * @param {string} firstFormId - 1回目フォームID（回答済み判定用）
 */
function sendConfirmationFormToMembers(formUrl, year, month, firstFormId) {
  var members = getScheduleMembers();
  var sentCount = 0;

  var deadline = new Date();
  deadline.setDate(deadline.getDate() + CONFIG.FORM_POLLING.FINAL_DEADLINE_DAYS);
  var deadlineStr = Utilities.formatDate(deadline, 'Asia/Tokyo', 'yyyy/MM/dd');

  // 1回目フォームの回答者メール一覧を取得
  var respondedEmails = {};
  try {
    var firstForm = FormApp.openById(firstFormId);
    var responses = firstForm.getResponses();
    responses.forEach(function(resp) {
      var respEmail = resp.getRespondentEmail();
      if (respEmail) {
        respondedEmails[respEmail.toLowerCase()] = true;
      }
    });
  } catch (err) {
    console.log('1回目フォーム回答取得失敗: ' + err.message);
  }

  var subject = '【最終確認】' + year + '年' + month + '月 日程の確定確認';

  members.forEach(function(member) {
    if (!member.email) return;

    var hasResponded = respondedEmails[member.email.toLowerCase()] === true;
    var body;

    if (hasResponded) {
      // 回答済み者向け
      body = member.name + ' 様\n\n' +
        year + '年' + month + '月の日程について、ご回答ありがとうございました。\n\n' +
        '予定に変更がなければ「現在の内容でOK」を選択して送信してください。\n' +
        '追加で参加可能になった日がある場合はチェックして送信してください。\n\n' +
        '▼ 確認フォーム\n' + formUrl + '\n\n' +
        '最終期限: ' + deadlineStr + '\n' +
        '※ 期限を過ぎると現在の登録内容で確定となります。\n\n' +
        'よろしくお願いいたします。\n\n' +
        CONFIG.ORG.NAME;
    } else {
      // 未回答者向け
      body = member.name + ' 様\n\n' +
        year + '年' + month + '月の日程について、1回目のアンケートにご回答がなかったため、\n' +
        '現在「全日程参加可能」として登録されています。\n\n' +
        'このまま確定してよい場合は「現在の内容でOK」を選択して送信してください。\n' +
        '参加できない日がある場合はお早めにお知らせください。\n\n' +
        '▼ 確認フォーム\n' + formUrl + '\n\n' +
        '最終期限: ' + deadlineStr + '\n' +
        '※ 期限を過ぎると現在の登録内容で確定となります。\n\n' +
        'よろしくお願いいたします。\n\n' +
        CONFIG.ORG.NAME;
    }

    GmailApp.sendEmail(member.email, subject, body, {
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.REPLY_TO
    });

    sentCount++;
    console.log('2回目メール送信完了: ' + member.name + (hasResponded ? ' (回答済み)' : ' (未回答)'));
  });

  console.log('2回目確認メール送信完了: ' + sentCount + '名に送信');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 自動クローズ（条件未達枠の確定処理）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 条件未達枠を自動クローズし、管理者に確定結果レポートを送信
 * 毎月21日に実行
 */
function finalizeSchedule() {
  var now = new Date();
  // 対象月を計算（21日実行 → 翌月分を確定）
  var targetDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var year = targetDate.getFullYear();
  var month = targetDate.getMonth() + 1;

  console.log('日程確定処理開始: ' + year + '年' + month + '月');

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var confirmedSlots = [];
  var cancelledSlots = [];

  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
    if (!dateVal) continue;

    var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (date.getFullYear() !== year || (date.getMonth() + 1) !== month) continue;

    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
    var dow = date.getDay();
    var dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    var dateLabel = month + '/' + date.getDate() + '(' + dayNames[dow] + ')';

    var timeStr;
    var time = data[i][SCHEDULE_COLUMNS.TIME];
    if (time instanceof Date) {
      timeStr = Utilities.formatDate(time, 'Asia/Tokyo', 'HH:mm');
    } else {
      timeStr = String(time);
    }

    var members = data[i][SCHEDULE_COLUMNS.MEMBERS] ? data[i][SCHEDULE_COLUMNS.MEMBERS].toString() : '';
    var score = data[i][SCHEDULE_COLUMNS.SCORE] || 0;
    var bookable = data[i][SCHEDULE_COLUMNS.BOOKABLE] || '×';

    var slotInfo = {
      dateLabel: dateLabel,
      time: timeStr,
      score: score,
      members: members,
      bookable: bookable
    };

    if (bookable === '○') {
      confirmedSlots.push(slotInfo);
    } else {
      // 条件未達: C列（対応可否）を×に変更
      sheet.getRange(i + 1, SCHEDULE_COLUMNS.AVAILABLE + 1).setValue('×');
      cancelledSlots.push(slotInfo);
    }
  }

  // 管理者に確定レポートメール送信
  var reportBody = year + '年' + month + '月の日程アンケート結果を確定しました。\n\n';

  reportBody += '━━ 開催確定枠（' + confirmedSlots.length + '枠） ━━\n';
  if (confirmedSlots.length > 0) {
    confirmedSlots.forEach(function(slot) {
      reportBody += slot.dateLabel + ' ' + slot.time +
        '  配置点数:' + slot.score +
        '  メンバー: ' + (slot.members || 'なし') + '  ○\n';
    });
  } else {
    reportBody += '（なし）\n';
  }

  reportBody += '\n━━ 開催中止枠（' + cancelledSlots.length + '枠） ━━\n';
  if (cancelledSlots.length > 0) {
    cancelledSlots.forEach(function(slot) {
      reportBody += slot.dateLabel + ' ' + slot.time +
        '  配置点数:' + slot.score +
        '  メンバー: ' + (slot.members || 'なし') + '  × メンバー不足\n';
    });
  } else {
    reportBody += '（なし）\n';
  }

  reportBody += '\n※ 開催確定枠のみLPに表示されます。';

  var reportSubject = '【日程確定】' + year + '年' + month + '月 開催枠の確定結果';

  CONFIG.ADMIN_EMAILS.forEach(function(email) {
    GmailApp.sendEmail(email, reportSubject, reportBody, {
      name: CONFIG.SENDER_NAME
    });
  });

  console.log('日程確定完了: 確定' + confirmedSlots.length + '枠, 中止' + cancelledSlots.length + '枠');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 月次自動実行
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 1回目: 毎月25日に自動実行
 * 翌々月の日程生成 → 全メンバー登録 → オプトアウトフォーム作成 → メール送信
 */
function runFirstPolling() {
  var now = new Date();
  var targetDate = new Date(now.getFullYear(), now.getMonth() + 2, 1);
  var year = targetDate.getFullYear();
  var month = targetDate.getMonth() + 1;

  console.log('1回目ポーリング開始: ' + year + '年' + month + '月');

  // Step 1: 翌々月の日程を生成
  generateNextMonthSchedule();
  console.log('日程生成完了');

  // Step 2: 全メンバーをデフォルト登録
  registerAllMembersForMonth(year, month);
  console.log('全メンバーデフォルト登録完了');

  // Step 3: オプトアウトフォーム作成
  var form = createOptOutForm(year, month);
  var formUrl = form.getPublishedUrl();

  // Step 4: フォーム情報を保存
  PropertiesService.getScriptProperties().setProperty(
    'polling_form_' + year + '_' + month,
    JSON.stringify({
      formId: form.getId(),
      formUrl: formUrl,
      editUrl: form.getEditUrl(),
      targetYear: year,
      targetMonth: month,
      createdAt: new Date().toISOString(),
      round: 1
    })
  );

  // Step 5: メンバーにメール送信
  sendOptOutFormToMembers(formUrl, year, month);

  // Step 6: 管理者にも通知
  var adminSubject = '【自動通知】' + year + '年' + month + '月 日程アンケート（1回目）送信完了';
  var adminBody = year + '年' + month + '月の日程アンケート（オプトアウト形式）を作成し、メンバーに送信しました。\n\n' +
    'フォームURL: ' + formUrl + '\n' +
    '編集URL: ' + form.getEditUrl() + '\n\n' +
    '回答状況はフォームの「回答」タブから確認できます。\n' +
    '2回目の確認フォームは翌月15日に自動送信されます。';

  CONFIG.ADMIN_EMAILS.forEach(function(email) {
    GmailApp.sendEmail(email, adminSubject, adminBody, {
      name: CONFIG.SENDER_NAME
    });
  });

  console.log('1回目ポーリング完了');
}

/**
 * 2回目: 毎月15日に自動実行
 * 確認+追加フォーム作成 → メール送信（未回答者/回答済み者で文面分け）
 */
function runReminderPolling() {
  var now = new Date();
  // 翌月分（15日実行 = 前月25日に作った分の確認）
  var targetDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  var year = targetDate.getFullYear();
  var month = targetDate.getMonth() + 1;

  console.log('2回目ポーリング開始: ' + year + '年' + month + '月');

  // 1回目のフォーム情報を取得
  var propKey = 'polling_form_' + year + '_' + month;
  var propValue = PropertiesService.getScriptProperties().getProperty(propKey);
  var firstFormId = null;

  if (propValue) {
    try {
      var propData = JSON.parse(propValue);
      firstFormId = propData.formId;
    } catch (err) {
      console.log('1回目フォーム情報取得失敗: ' + err.message);
    }
  }

  // Step 1: 確認+追加フォーム作成
  var form = createConfirmationForm(year, month);
  var formUrl = form.getPublishedUrl();

  // Step 2: フォーム情報を保存
  PropertiesService.getScriptProperties().setProperty(
    'polling_confirm_form_' + year + '_' + month,
    JSON.stringify({
      formId: form.getId(),
      formUrl: formUrl,
      targetYear: year,
      targetMonth: month,
      createdAt: new Date().toISOString(),
      round: 2
    })
  );

  // Step 3: メンバーにメール送信（未回答/回答済みで文面分け）
  sendConfirmationFormToMembers(formUrl, year, month, firstFormId);

  // Step 4: 管理者にも通知
  var adminSubject = '【自動通知】' + year + '年' + month + '月 日程確認フォーム（2回目）送信完了';
  var adminBody = year + '年' + month + '月の日程確認フォームを作成し、メンバーに送信しました。\n\n' +
    'フォームURL: ' + formUrl + '\n' +
    '編集URL: ' + form.getEditUrl() + '\n\n' +
    '日程確定は21日に自動実行されます。';

  CONFIG.ADMIN_EMAILS.forEach(function(email) {
    GmailApp.sendEmail(email, adminSubject, adminBody, {
      name: CONFIG.SENDER_NAME
    });
  });

  console.log('2回目ポーリング完了');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// トリガー設定
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 毎月25日のトリガーを設定（1回目）
 */
function setupFirstPollingTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runFirstPolling') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runFirstPolling')
    .timeBased()
    .onMonthDay(CONFIG.FORM_POLLING.FIRST_SEND_DAY)
    .atHour(CONFIG.FORM_POLLING.FIRST_SEND_HOUR)
    .create();

  console.log('1回目ポーリングトリガー設定完了（毎月' + CONFIG.FORM_POLLING.FIRST_SEND_DAY + '日 ' + CONFIG.FORM_POLLING.FIRST_SEND_HOUR + '時）');
}

/**
 * 毎月15日のトリガーを設定（2回目）
 */
function setupReminderPollingTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'runReminderPolling') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runReminderPolling')
    .timeBased()
    .onMonthDay(CONFIG.FORM_POLLING.REMINDER_SEND_DAY)
    .atHour(CONFIG.FORM_POLLING.REMINDER_SEND_HOUR)
    .create();

  console.log('2回目ポーリングトリガー設定完了（毎月' + CONFIG.FORM_POLLING.REMINDER_SEND_DAY + '日 ' + CONFIG.FORM_POLLING.REMINDER_SEND_HOUR + '時）');
}

/**
 * 毎月21日のトリガーを設定（日程確定）
 */
function setupFinalizeScheduleTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'finalizeSchedule') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('finalizeSchedule')
    .timeBased()
    .onMonthDay(CONFIG.FORM_POLLING.FINALIZE_DAY)
    .atHour(CONFIG.FORM_POLLING.FINALIZE_HOUR)
    .create();

  console.log('日程確定トリガー設定完了（毎月' + CONFIG.FORM_POLLING.FINALIZE_DAY + '日 ' + CONFIG.FORM_POLLING.FINALIZE_HOUR + '時）');
}

/**
 * フォーム送信時のトリガーを設定
 * @param {string} formId - GoogleフォームのID
 * @param {string} handlerFunction - ハンドラ関数名
 */
function setupFormSubmitTrigger(formId, handlerFunction) {
  // このフォーム用の既存トリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handlerFunction &&
        trigger.getTriggerSourceId() === formId) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(handlerFunction)
    .forForm(formId)
    .onFormSubmit()
    .create();

  console.log('フォーム送信トリガー設定完了: ' + handlerFunction + ' (' + formId + ')');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// テスト・手動実行用関数
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * テスト用: オプトアウトフォーム作成テスト
 */
function testCreateOptOutForm() {
  var year = 2026;
  var month = 4;

  var form = createOptOutForm(year, month);
  console.log('テストフォーム作成完了');
  console.log('公開URL: ' + form.getPublishedUrl());
  console.log('編集URL: ' + form.getEditUrl());
}

/**
 * テスト用: 確認フォーム作成テスト
 */
function testCreateConfirmationForm() {
  var year = 2026;
  var month = 4;

  var form = createConfirmationForm(year, month);
  console.log('確認フォーム作成完了');
  console.log('公開URL: ' + form.getPublishedUrl());
  console.log('編集URL: ' + form.getEditUrl());
}

/**
 * テスト用: 全メンバーデフォルト登録テスト
 */
function testRegisterAllMembers() {
  registerAllMembersForMonth(2026, 4);
}

/**
 * テスト用: 日程確定テスト
 */
function testFinalizeSchedule() {
  finalizeSchedule();
}

/**
 * 手動実行: 指定月の1回目ポーリング
 */
function manualFirstPolling() {
  var year = 2026;
  var month = 4;

  generateNextMonthSchedule();
  registerAllMembersForMonth(year, month);
  var form = createOptOutForm(year, month);
  sendOptOutFormToMembers(form.getPublishedUrl(), year, month);

  console.log('手動1回目送信完了: ' + year + '年' + month + '月');
}

/**
 * 手動実行: 指定月の2回目ポーリング
 */
function manualReminderPolling() {
  var year = 2026;
  var month = 4;

  var propKey = 'polling_form_' + year + '_' + month;
  var propValue = PropertiesService.getScriptProperties().getProperty(propKey);
  var firstFormId = null;
  if (propValue) {
    try {
      firstFormId = JSON.parse(propValue).formId;
    } catch (e) {}
  }

  var form = createConfirmationForm(year, month);
  sendConfirmationFormToMembers(form.getPublishedUrl(), year, month, firstFormId);

  console.log('手動2回目送信完了: ' + year + '年' + month + '月');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 回答集計シート生成
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 回答集計シートを生成・更新
 * 日程×メンバーのマトリクス表示
 */
function generateSummarySheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // シート取得または作成
  var sheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(CONFIG.SUMMARY_SHEET_NAME);
  }

  // 日程設定シートからデータ取得
  var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!scheduleSheet) {
    sheet.getRange(1, 1).setValue('日程設定シートが見つかりません');
    return;
  }

  var scheduleData = scheduleSheet.getDataRange().getValues();

  // スケジュール対象メンバー取得
  var members = getScheduleMembers();
  var memberNames = members.map(function(m) { return m.name; });

  // 月ごとにグループ化
  var monthGroups = {};
  for (var i = 1; i < scheduleData.length; i++) {
    var dateVal = scheduleData[i][SCHEDULE_COLUMNS.DATE];
    if (!dateVal) continue;

    var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
    var monthKey = date.getFullYear() + '年' + (date.getMonth() + 1) + '月';

    if (!monthGroups[monthKey]) {
      monthGroups[monthKey] = [];
    }

    var timeVal = scheduleData[i][SCHEDULE_COLUMNS.TIME];
    var timeStr;
    if (timeVal instanceof Date) {
      timeStr = Utilities.formatDate(timeVal, 'Asia/Tokyo', 'HH:mm');
    } else {
      timeStr = String(timeVal);
    }

    var dow = date.getDay();
    var dayNames = ['日', '月', '火', '水', '木', '金', '土'];
    var dateLabel = (date.getMonth() + 1) + '/' + date.getDate() + '(' + dayNames[dow] + ')';

    var membersStr = scheduleData[i][SCHEDULE_COLUMNS.MEMBERS]
      ? scheduleData[i][SCHEDULE_COLUMNS.MEMBERS].toString() : '';
    var registeredMembers = membersStr
      ? membersStr.split(',').map(function(n) { return n.trim(); })
      : [];

    monthGroups[monthKey].push({
      dateLabel: dateLabel,
      time: timeStr,
      method: scheduleData[i][SCHEDULE_COLUMNS.METHOD] || '',
      available: scheduleData[i][SCHEDULE_COLUMNS.AVAILABLE] || '',
      booking: scheduleData[i][SCHEDULE_COLUMNS.BOOKING_STATUS] || '',
      score: scheduleData[i][SCHEDULE_COLUMNS.SCORE] || 0,
      bookable: scheduleData[i][SCHEDULE_COLUMNS.BOOKABLE] || '×',
      registeredMembers: registeredMembers
    });
  }

  // フォーム回答状況を取得
  var props = PropertiesService.getScriptProperties().getProperties();
  var formResponses = {};
  Object.keys(props).forEach(function(key) {
    if (key.indexOf('polling_form_') !== 0 && key.indexOf('polling_confirm_form_') !== 0) return;
    try {
      var formInfo = JSON.parse(props[key]);
      if (formInfo.formId) {
        var form = FormApp.openById(formInfo.formId);
        var responses = form.getResponses();
        responses.forEach(function(resp) {
          var email = resp.getRespondentEmail();
          if (email) {
            formResponses[email.toLowerCase()] = {
              submittedAt: resp.getTimestamp(),
              round: formInfo.round
            };
          }
        });
      }
    } catch (e) {}
  });

  // シート書き込み
  var currentRow = 1;

  var monthKeys = Object.keys(monthGroups).sort();
  monthKeys.forEach(function(monthKey) {
    var slots = monthGroups[monthKey];

    // === 月タイトル ===
    sheet.getRange(currentRow, 1).setValue(monthKey + ' 回答集計');
    sheet.getRange(currentRow, 1, 1, memberNames.length + 5).merge();
    sheet.getRange(currentRow, 1).setFontSize(14).setFontWeight('bold')
      .setBackground('#0f2350').setFontColor('#ffffff');
    currentRow++;

    // === ヘッダー行 ===
    var headers = ['日付', '時間', '方法', '予約状況', '判定'];
    memberNames.forEach(function(name) {
      // 苗字のみ表示
      headers.push(name.split(' ')[0] || name);
    });
    headers.push('参加数', '配置点数');

    sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(currentRow, 1, 1, headers.length)
      .setBackground('#34495e').setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center');
    currentRow++;

    // === データ行 ===
    slots.forEach(function(slot) {
      var row = [
        slot.dateLabel,
        slot.time,
        slot.method === 'オンライン' ? 'Zoom' : slot.method === '両方' ? '対面/Zoom' : slot.method,
        slot.booking || '空き',
        slot.booking === '予約済み' ? '×' : slot.bookable
      ];

      var participantCount = 0;
      memberNames.forEach(function(name) {
        if (slot.registeredMembers.indexOf(name) !== -1) {
          row.push('○');
          participantCount++;
        } else {
          row.push('×');
        }
      });

      row.push(participantCount);
      row.push(slot.score);

      sheet.getRange(currentRow, 1, 1, row.length).setValues([row]);

      // セル書式設定
      for (var col = 0; col < memberNames.length; col++) {
        var cell = sheet.getRange(currentRow, 6 + col);
        if (slot.registeredMembers.indexOf(memberNames[col]) !== -1) {
          cell.setBackground('#d4edda').setHorizontalAlignment('center');
        } else {
          cell.setBackground('#f8d7da').setHorizontalAlignment('center');
        }
      }

      // 予約状況の色
      var bookingCell = sheet.getRange(currentRow, 4);
      if (slot.booking === '予約済み') {
        bookingCell.setBackground('#cce5ff').setFontWeight('bold');
      }

      // 判定の色（予約済みも×扱い）
      var bookableCell = sheet.getRange(currentRow, 5);
      bookableCell.setHorizontalAlignment('center');
      if (slot.booking === '予約済み') {
        bookableCell.setBackground('#cce5ff');
      } else if (slot.bookable === '○') {
        bookableCell.setBackground('#d4edda');
      } else {
        bookableCell.setBackground('#f8d7da');
      }

      currentRow++;
    });

    // === メンバー別 参加枠数サマリ ===
    currentRow++;
    var summaryRow = ['', '', '', '', '参加枠数'];
    memberNames.forEach(function(name) {
      var count = 0;
      slots.forEach(function(slot) {
        if (slot.registeredMembers.indexOf(name) !== -1) count++;
      });
      summaryRow.push(count + '/' + slots.length);
    });
    summaryRow.push('', '');
    sheet.getRange(currentRow, 1, 1, summaryRow.length).setValues([summaryRow]);
    sheet.getRange(currentRow, 1, 1, summaryRow.length)
      .setBackground('#eee').setFontWeight('bold').setHorizontalAlignment('center');
    currentRow++;

    // === フォーム回答状況 ===
    var responseRow = ['', '', '', '', '回答状況'];
    memberNames.forEach(function(name) {
      var member = members.find(function(m) { return m.name === name; });
      if (member && member.email) {
        var resp = formResponses[member.email.toLowerCase()];
        if (resp) {
          var d = resp.submittedAt;
          responseRow.push('済 ' + (d.getMonth() + 1) + '/' + d.getDate());
        } else {
          responseRow.push('未回答');
        }
      } else {
        responseRow.push('メール未設定');
      }
    });
    responseRow.push('', '');
    sheet.getRange(currentRow, 1, 1, responseRow.length).setValues([responseRow]);

    // 未回答を赤く
    for (var col = 0; col < memberNames.length; col++) {
      var respCell = sheet.getRange(currentRow, 6 + col);
      respCell.setHorizontalAlignment('center');
      var val = responseRow[5 + col];
      if (val === '未回答') {
        respCell.setBackground('#f8d7da').setFontColor('#721c24');
      } else if (val.indexOf('済') === 0) {
        respCell.setBackground('#d4edda').setFontColor('#155724');
      }
    }

    currentRow += 2;
  });

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 日付
  sheet.setColumnWidth(2, 60);   // 時間
  sheet.setColumnWidth(3, 80);   // 方法
  sheet.setColumnWidth(4, 80);   // 予約状況
  sheet.setColumnWidth(5, 60);   // 判定
  for (var c = 0; c < memberNames.length; c++) {
    sheet.setColumnWidth(6 + c, 70);
  }
  sheet.setColumnWidth(6 + memberNames.length, 60);     // 参加数
  sheet.setColumnWidth(6 + memberNames.length + 1, 70); // 配置点数

  // 行固定なし（月ごとにヘッダーがあるため）
  sheet.setFrozenRows(0);

  console.log('回答集計シートを生成しました');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ポーリング状況確認・回答集計
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * ポーリングの全体状況を取得
 * トリガー設定状況、フォーム作成状況、回答状況を一覧で返す
 * @returns {Object} ポーリング状況
 */
function getPollingStatus() {
  var result = {
    success: true,
    timestamp: new Date().toISOString(),
    triggers: [],
    forms: [],
    schedule_summary: [],
    members: []
  };

  // 1. トリガー設定状況
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    result.triggers.push({
      handler: trigger.getHandlerFunction(),
      type: trigger.getEventType().toString(),
      sourceId: trigger.getTriggerSourceId() || ''
    });
  });

  // 2. 保存済みフォーム情報を検索（直近6ヶ月分）
  var props = PropertiesService.getScriptProperties().getProperties();
  var formKeys = Object.keys(props).filter(function(k) {
    return k.indexOf('polling_form_') === 0 || k.indexOf('polling_confirm_form_') === 0;
  });

  formKeys.sort().forEach(function(key) {
    try {
      var formInfo = JSON.parse(props[key]);
      var formEntry = {
        key: key,
        targetYear: formInfo.targetYear,
        targetMonth: formInfo.targetMonth,
        round: formInfo.round,
        createdAt: formInfo.createdAt,
        formUrl: formInfo.formUrl || '',
        responseCount: 0,
        responses: []
      };

      // フォームの回答を取得
      if (formInfo.formId) {
        try {
          var form = FormApp.openById(formInfo.formId);
          var responses = form.getResponses();
          formEntry.responseCount = responses.length;

          responses.forEach(function(resp) {
            var respData = {
              email: resp.getRespondentEmail() || '不明',
              submittedAt: resp.getTimestamp().toISOString(),
              answers: {}
            };

            resp.getItemResponses().forEach(function(ir) {
              respData.answers[ir.getItem().getTitle()] = ir.getResponse();
            });

            formEntry.responses.push(respData);
          });
        } catch (formErr) {
          formEntry.error = 'フォーム取得失敗: ' + formErr.message;
        }
      }

      result.forms.push(formEntry);
    } catch (e) {
      result.forms.push({ key: key, error: e.message });
    }
  });

  // 3. 日程設定シートの集計（直近3ヶ月分）
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      var monthSummary = {};

      for (var i = 1; i < data.length; i++) {
        var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
        if (!dateVal) continue;

        var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
        var monthKey = date.getFullYear() + '年' + (date.getMonth() + 1) + '月';

        if (!monthSummary[monthKey]) {
          monthSummary[monthKey] = { total: 0, bookable: 0, booked: 0, closed: 0, members_set: {} };
        }

        monthSummary[monthKey].total++;

        var bookable = data[i][SCHEDULE_COLUMNS.BOOKABLE];
        var booking = data[i][SCHEDULE_COLUMNS.BOOKING_STATUS];
        var available = data[i][SCHEDULE_COLUMNS.AVAILABLE];

        if (bookable === '○') monthSummary[monthKey].bookable++;
        if (booking === '予約済み') monthSummary[monthKey].booked++;
        if (available === '×') monthSummary[monthKey].closed++;

        // 参加メンバーの集計
        var members = data[i][SCHEDULE_COLUMNS.MEMBERS];
        if (members) {
          members.toString().split(',').forEach(function(name) {
            var n = name.trim();
            if (n) {
              if (!monthSummary[monthKey].members_set[n]) {
                monthSummary[monthKey].members_set[n] = 0;
              }
              monthSummary[monthKey].members_set[n]++;
            }
          });
        }
      }

      Object.keys(monthSummary).sort().forEach(function(monthKey) {
        var s = monthSummary[monthKey];
        var memberList = [];
        Object.keys(s.members_set).forEach(function(name) {
          memberList.push({ name: name, slots: s.members_set[name] });
        });
        memberList.sort(function(a, b) { return b.slots - a.slots; });

        result.schedule_summary.push({
          month: monthKey,
          total_slots: s.total,
          bookable_slots: s.bookable,
          booked_slots: s.booked,
          closed_slots: s.closed,
          members: memberList
        });
      });
    }
  } catch (e) {
    result.schedule_error = e.message;
  }

  // 4. メンバー一覧
  var members = getAllMembers();
  members.forEach(function(m) {
    result.members.push({
      name: m.name,
      term: m.term,
      type: m.type,
      email: m.email || '未設定',
      isScheduleTarget: m.type !== '顧問'
    });
  });

  return result;
}
