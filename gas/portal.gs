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
