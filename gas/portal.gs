/**
 * スタッフポータル 追加API（Phase 5）
 *
 * auth.gs のAPIに加え、以下の機能を提供:
 *   - シフト参加/取消
 *   - プロフィール取得
 *   - ニュース投稿
 *   - メンバー一覧（管理者）
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// シフト管理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * シフト参加/取消
 * @param {Object} session - セッションデータ
 * @param {string|number} rowStr - 行番号
 * @param {boolean} join - 参加する場合 true
 * @returns {Object} { success, message }
 */
function toggleShiftParticipation(session, rowStr, join) {
  if (!rowStr) return { success: false, message: '行番号が指定されていません' };

  // 取消は警告確認済みの前提で許可

  var row = parseInt(rowStr);
  if (isNaN(row) || row < 2) return { success: false, message: '無効な行番号です' };

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!sheet) return { success: false, message: '日程設定シートが見つかりません' };

  // MEMBERS列 (参加者リスト)
  var membersCol = SCHEDULE_COLUMNS.MEMBERS + 1; // 1-indexed
  var currentVal = (sheet.getRange(row, membersCol).getValue() || '').toString();
  var memberName = session.name;

  var members = currentVal ? currentVal.split(',').map(function(m) { return m.trim(); }).filter(Boolean) : [];
  var idx = members.indexOf(memberName);

  if (join) {
    if (idx >= 0) return { success: true, message: '既に参加登録済みです' };
    members.push(memberName);
  } else {
    if (idx < 0) return { success: true, message: '参加登録されていません' };
    members.splice(idx, 1);
  }

  sheet.getRange(row, membersCol).setValue(members.join(', '));

  // 配置点数を再計算
  try {
    recalculateScheduleScore(sheet, row);
  } catch (e) {
    console.log('スコア再計算スキップ:', e.message);
  }

  console.log('シフト' + (join ? '参加' : '取消') + ': ' + memberName + ' (行' + row + ')');
  return { success: true, message: join ? '参加を登録しました' : '参加を取り消しました' };
}

/**
 * 日程の配置点数を再計算
 * @param {Sheet} sheet - 日程設定シート
 * @param {number} row - 行番号
 */
function recalculateScheduleScore(sheet, row) {
  var membersCol = SCHEDULE_COLUMNS.MEMBERS + 1;
  var scoreCol = SCHEDULE_COLUMNS.SCORE + 1;
  var membersStr = (sheet.getRange(row, membersCol).getValue() || '').toString();

  if (!membersStr) {
    sheet.getRange(row, scoreCol).setValue(0);
    return;
  }

  var names = membersStr.split(',').map(function(n) { return n.trim(); }).filter(Boolean);
  var totalScore = 0;

  // メンバーマスタから点数を取得
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var memberSheet = ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME);
  if (!memberSheet || memberSheet.getLastRow() < 2) {
    sheet.getRange(row, scoreCol).setValue(names.length);
    return;
  }

  var memberData = memberSheet.getDataRange().getValues();
  for (var n = 0; n < names.length; n++) {
    var found = false;
    for (var m = 1; m < memberData.length; m++) {
      if (memberData[m][MEMBER_COLUMNS.NAME] === names[n]) {
        var term = (memberData[m][MEMBER_COLUMNS.TERM] || '').toString();
        var cert = (memberData[m][MEMBER_COLUMNS.CERT] || '').toString();
        // 1期/2期の診断士=2pt, 3期/4期やオブザーバー=1pt
        if ((term === '1期' || term === '2期') && cert.indexOf('診断士') >= 0) {
          totalScore += 2;
        } else {
          totalScore += 1;
        }
        found = true;
        break;
      }
    }
    if (!found) totalScore += 1; // マスタにない場合は1pt
  }

  sheet.getRange(row, scoreCol).setValue(totalScore);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// プロフィール
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * プロフィール情報を取得
 * @param {Object} session - セッションデータ
 * @returns {Object} プロフィールデータ
 */
function getPortalProfile(session) {
  var member = getMemberByEmail(session.email);
  if (!member) return { success: false, message: 'メンバー情報が見つかりません' };

  return {
    success: true,
    profile: {
      name: member.name,
      email: session.email,
      term: member.term,
      cert: member.cert,
      type: member.type,
      phone: member.phone,
      specialties: member.specialties,
      themes: member.themes,
      titles: member.titles,
      role: session.role
    }
  };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ニュース投稿
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * ニュースを投稿（リーダー/管理者）
 * @param {Object} session - セッションデータ
 * @param {Object} params - { title, body, category }
 * @returns {Object} { success, message }
 */
function postPortalNews(session, params) {
  if (!params.title || !params.body) {
    return { success: false, message: 'タイトルと本文は必須です' };
  }

  try {
    var result = addNews({
      title: params.title,
      body: params.body,
      category: params.category || 'お知らせ',
      author: session.name
    });
    return { success: true, message: 'ニュースを投稿しました', id: result.id || '' };
  } catch (e) {
    return { success: false, message: 'ニュース投稿に失敗しました: ' + e.message };
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// プロフィール変更依頼
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * プロフィール変更依頼を管理者にメール送信
 * @param {Object} session - セッションデータ
 * @param {string} detail - 変更依頼の内容
 * @returns {Object} { success, message }
 */
function requestProfileChange(session, detail) {
  if (!detail || !detail.trim()) {
    return { success: false, message: '変更内容を入力してください' };
  }

  var subject = '【ポータル】プロフィール変更依頼 - ' + session.name;
  var body = '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'プロフィール変更依頼\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '依頼者: ' + session.name + '\n' +
    'メール: ' + session.email + '\n' +
    'ロール: ' + session.role + '\n' +
    '日時: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
    '■ 変更内容:\n' + detail.trim() + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'スタッフポータルから自動送信';

  try {
    CONFIG.ADMIN_EMAILS.forEach(function(adminEmail) {
      GmailApp.sendEmail(adminEmail, subject, body, { name: CONFIG.SENDER_NAME });
    });
    console.log('プロフィール変更依頼送信: ' + session.name);
    return { success: true, message: '変更依頼を管理者に送信しました' };
  } catch (e) {
    console.error('プロフィール変更依頼エラー:', e);
    return { success: false, message: '送信に失敗しました: ' + e.message };
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// メンバー一覧（管理者向け）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 全メンバー一覧を取得（管理者専用）
 * @returns {Object} メンバーリスト
 */
function getPortalMembers() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: true, members: [] };
  }

  var data = sheet.getDataRange().getValues();
  var members = [];

  for (var i = 1; i < data.length; i++) {
    var active = data[i][MEMBER_COLUMNS.ACTIVE];
    members.push({
      name: data[i][MEMBER_COLUMNS.NAME],
      term: data[i][MEMBER_COLUMNS.TERM],
      cert: data[i][MEMBER_COLUMNS.CERT],
      type: data[i][MEMBER_COLUMNS.TYPE],
      email: data[i][MEMBER_COLUMNS.EMAIL],
      specialties: data[i][MEMBER_COLUMNS.SPECIALTIES],
      themes: data[i][MEMBER_COLUMNS.THEMES],
      titles: data[i][MEMBER_COLUMNS.TITLES],
      active: active !== false && active !== 'FALSE' && active !== false
    });
  }

  return { success: true, members: members, total: members.length };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 会場設定 + 予約確定
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 会場を設定しステータスを確定に変更（対面相談用）
 * @param {Object} session - セッションデータ
 * @param {Object} params - { row, venue }
 * @returns {Object} { success, message }
 */
function setVenueAndConfirm(session, params) {
  var row = parseInt(params.row);
  var venue = (params.venue || '').trim();
  if (!row || isNaN(row) || row < 2) return { success: false, message: '無効な行番号です' };
  if (!venue) return { success: false, message: '会場を選択してください' };

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return { success: false, message: 'シートが見つかりません' };

  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var currentStatus = data[COLUMNS.STATUS];
  var method = (data[COLUMNS.METHOD] || '').toString();

  // Zoom相談は会場不要
  if (method === 'オンライン' || method === 'zoom' || method === 'Zoom') {
    return { success: false, message: 'オンライン相談は会場設定不要です' };
  }

  // 既に確定済みの場合
  if (currentStatus === STATUS.CONFIRMED) {
    return { success: false, message: '既に確定済みです' };
  }

  // 会場をN列にセット
  sheet.getRange(row, COLUMNS.LOCATION + 1).setValue(venue);
  data.location = venue;

  // ── 確定日時の解決 ──
  // Q列（確定日時）が空ならK列+日程設定シートから補完
  if (!data[COLUMNS.CONFIRMED_DATE]) {
    var fullDateTime = resolveConfirmedDateTime(row, sheet);
    if (fullDateTime) {
      sheet.getRange(row, COLUMNS.CONFIRMED_DATE + 1).setValue(fullDateTime);
    }
  }

  // ── 担当者自動選定（P列が未設定の場合のみ） ──
  var existingStaff = (data[COLUMNS.STAFF] || '').toString().trim();
  var staffSelectionResult = null;

  if (!existingStaff) {
    var consultTheme = (data[COLUMNS.THEME] || '').toString();
    var specialFlag = false;
    // 日程設定シートから該当日の特別対応フラグを取得
    try {
      var schedSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
      if (schedSheet) {
        var schedData = schedSheet.getDataRange().getValues();
        // 確定日時 or 希望日時1 から日付を特定
        var confirmedOrDate1 = data[COLUMNS.CONFIRMED_DATE] || data[COLUMNS.DATE1];
        if (confirmedOrDate1) {
          var targetDateStr = confirmedOrDate1 instanceof Date
            ? Utilities.formatDate(confirmedOrDate1, 'Asia/Tokyo', 'yyyy-MM-dd')
            : confirmedOrDate1.toString().substring(0, 10);
          for (var sd = 1; sd < schedData.length; sd++) {
            if (schedData[sd][SCHEDULE_COLUMNS.SPECIAL_FLAG] === true) {
              var schedDate = schedData[sd][SCHEDULE_COLUMNS.DATE];
              var schedDateStr = schedDate instanceof Date
                ? Utilities.formatDate(schedDate, 'Asia/Tokyo', 'yyyy-MM-dd')
                : (schedDate || '').toString().substring(0, 10);
              if (schedDateStr === targetDateStr) {
                specialFlag = true;
                break;
              }
            }
          }
        }
      }
    } catch (schedErr) {
      console.log('特別対応フラグ確認スキップ:', schedErr.message);
    }

    staffSelectionResult = selectStaffMembers(consultTheme, specialFlag);
    if (!staffSelectionResult.success) {
      // 選定失敗 → 確定を中止（会場設定は元に戻さない）
      return { success: false, message: '担当者の自動選定に失敗しました: ' + staffSelectionResult.message };
    }

    // P列に確定メンバーを書き込み
    sheet.getRange(row, COLUMNS.STAFF + 1).setValue(staffSelectionResult.confirmed.join(', '));

    // T列（備考）に予備メンバーを追記
    if (staffSelectionResult.reserve.length > 0) {
      var currentNotes = (data[COLUMNS.NOTES] || '').toString();
      var reserveNote = '予備: ' + staffSelectionResult.reserve.join(', ');
      var newNotes = currentNotes ? currentNotes + '\n' + reserveNote : reserveNote;
      sheet.getRange(row, COLUMNS.NOTES + 1).setValue(newNotes);
    }

    console.log('担当者自動選定: 確定=' + staffSelectionResult.confirmed.join(',') +
      ' 予備=' + staffSelectionResult.reserve.join(',') +
      ' スコア=' + staffSelectionResult.score);
  }

  // ステータスを確定に変更
  sheet.getRange(row, COLUMNS.STATUS + 1).setValue(STATUS.CONFIRMED);

  // ── 以下、メール送信・日程更新をトリガーに依存せず直接実行 ──
  // getRowDataで最新データを取得（ステータス・確定日時・担当者反映済み）
  var rowData = getRowData(row);
  rowData.location = venue;

  // 日程設定シートの予約状況を「予約済み」に更新
  if (rowData.confirmedDate) {
    var parsed = parseConfirmedDateTime(rowData.confirmedDate);
    if (parsed.date) {
      var booked = markAsBooked(parsed.date, parsed.time);
      console.log('日程設定シート同期: ' + parsed.date + ' ' + (parsed.time || '') + ' → ' + (booked ? '予約済み' : '該当なし'));
    }
  }

  // リーダー自動選定（メール送信前に実行し、リーダー名をメールに反映）
  if (!rowData.leader) {
    try {
      autoSelectLeaderOnConfirm(row);
      // 選定結果を反映するためrowDataを再取得
      rowData = getRowData(row);
      rowData.location = venue;
    } catch (leaderErr) {
      console.error('リーダー選定エラー（ポータル確定時）:', leaderErr);
    }
  } else {
    // リーダー履歴に「予定」として記録
    var schedMembers = getParticipatingMembers(rowData.confirmedDate) || '';
    recordLeaderAssignment(rowData, rowData.leader, schedMembers, 0, '手動設定', '予定');
  }

  // 相談者に確定メール送信
  try {
    sendConfirmedEmail(rowData);
    console.log('確定メール送信完了: ' + rowData.email);
  } catch (mailErr) {
    console.error('確定メール送信エラー:', mailErr);
  }

  // 担当者（P列の確定メンバーのみ）に確定通知を送信
  try {
    var emailResult = buildStaffNotificationEmail_(rowData);
    var sentEmails = {};

    if (rowData.staff) {
      sendStaffNotifications(rowData.staff, emailResult.subject, emailResult.body);
      var staffNames = rowData.staff.split(',').map(function(n) { return n.trim(); }).filter(function(n) { return n; });
      staffNames.forEach(function(name) {
        var m = getMemberByName(name);
        if (m && m.email) sentEmails[m.email] = true;
      });
    }

    // 誰にも送れなかった場合は管理者にフォールバック
    if (Object.keys(sentEmails).length === 0) {
      CONFIG.ADMIN_EMAILS.forEach(function(adminEmail) {
        GmailApp.sendEmail(adminEmail, emailResult.subject, emailResult.body, {
          name: emailResult.senderName || CONFIG.SENDER_NAME
        });
      });
      console.log('確定通知: 担当者未設定のため管理者にフォールバック');
    }
  } catch (staffErr) {
    console.error('担当者通知エラー:', staffErr);
  }

  // 重複防止フラグをセット（onSheetEditトリガーでの二重送信防止）
  try {
    PropertiesService.getScriptProperties().setProperty('VENUE_CONFIRMED_ROW_' + row, String(Date.now()));
  } catch (propErr) {
    console.error('PropertiesServiceエラー:', propErr);
  }

  // レスポンスに選定結果を含める
  var resultMsg = '会場を「' + venue + '」に設定し、予約を確定しました';
  if (staffSelectionResult) {
    resultMsg += '\n確定: ' + staffSelectionResult.confirmed.join(', ') +
      '（' + staffSelectionResult.confirmed.length + '名/' + staffSelectionResult.score + '点）';
    if (staffSelectionResult.reserve.length > 0) {
      resultMsg += '\n予備: ' + staffSelectionResult.reserve.join(', ');
    }
  }

  console.log('会場設定+確定（メール送信済み）: ' + venue + ' (行' + row + ') by ' + session.name);
  return { success: true, message: resultMsg };
}
