/**
 * スタッフポータル認証システム（Phase 5）
 *
 * メールリンク認証（マジックリンク）方式:
 *   1. スタッフがメールアドレスを入力
 *   2. メンバーマスタで検証 → ログインリンクをメール送信
 *   3. リンクをクリック → セッション発行
 *   4. CacheService でセッション管理（最大6時間）
 *
 * ロール:
 *   - admin: 管理者（全機能利用可）
 *   - leader: リーダー（案件管理・レポート管理）
 *   - member: 一般メンバー（シフト管理・閲覧）
 */

/**
 * ポータル設定
 */
var PORTAL_CONFIG = {
  SESSION_HOURS: 6,       // セッション有効期間（時間）
  TOKEN_EXPIRY_MIN: 30,   // マジックリンクの有効期限（分）
  ADMIN_EMAILS: (CONFIG.ADMIN_EMAILS || []).map(function(e) { return e.toLowerCase(); })
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// マジックリンク認証
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * ログインリクエスト処理
 * メールアドレスを検証し、マジックリンクを送信
 * @param {string} email - メールアドレス
 * @returns {Object} { success, message }
 */
function requestPortalLogin(email) {
  if (!email) {
    return { success: false, message: 'メールアドレスを入力してください' };
  }

  email = email.trim().toLowerCase();

  // メンバーマスタで検索
  var member = getMemberByEmail(email);
  if (!member) {
    // セキュリティ上、存在しないメールでも同じメッセージを返す
    console.log('Portal login attempt - email not found: ' + email);
    return { success: true, message: 'ログインリンクをメールに送信しました。メールをご確認ください。' };
  }

  // アクティブでないメンバーはログイン不可
  if (member.active === false || member.active === 'FALSE') {
    console.log('Portal login attempt - inactive member: ' + email);
    return { success: true, message: 'ログインリンクをメールに送信しました。メールをご確認ください。' };
  }

  // マジックリンクトークン生成
  var token = Utilities.getUuid();
  var props = PropertiesService.getScriptProperties();
  var tokenData = JSON.stringify({
    email: email,
    name: member.name,
    createdAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + PORTAL_CONFIG.TOKEN_EXPIRY_MIN * 60 * 1000).toISOString()
  });
  props.setProperty('portal_login_' + token, tokenData);

  // マジックリンクURL
  var loginUrl = CONFIG.CONSENT.WEB_APP_URL + '?action=portal-verify&token=' + token;

  // メール送信
  var subject = '【ログイン】スタッフポータル - 関西学院大学 中小企業経営診断研究会';
  var body = member.name + ' 様\n\n' +
    'スタッフポータルへのログインリクエストを受け付けました。\n' +
    '以下のリンクをクリックしてログインしてください。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ ログインリンク\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    loginUrl + '\n\n' +
    '※ このリンクは' + PORTAL_CONFIG.TOKEN_EXPIRY_MIN + '分間有効です。\n' +
    '※ このリクエストに心当たりがない場合は、このメールを無視してください。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    CONFIG.ORG.NAME + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

  GmailApp.sendEmail(email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });

  console.log('Portal login link sent: ' + member.name + ' (' + email + ')');
  return { success: true, message: 'ログインリンクをメールに送信しました。メールをご確認ください。' };
}

/**
 * マジックリンクのトークンを検証してセッションを発行
 * @param {string} token - マジックリンクトークン
 * @returns {Object} { success, sessionId, member, role }
 */
function verifyPortalLogin(token) {
  if (!token) {
    return { success: false, message: '無効なリンクです' };
  }

  var props = PropertiesService.getScriptProperties();
  var tokenDataStr = props.getProperty('portal_login_' + token);
  if (!tokenDataStr) {
    return { success: false, message: 'このリンクは無効または使用済みです' };
  }

  var tokenData;
  try {
    tokenData = JSON.parse(tokenDataStr);
  } catch (e) {
    return { success: false, message: 'トークンの解析に失敗しました' };
  }

  // 有効期限チェック
  if (new Date() > new Date(tokenData.expiresAt)) {
    props.deleteProperty('portal_login_' + token);
    return { success: false, message: 'このリンクの有効期限が切れています。再度ログインしてください。' };
  }

  // トークンを無効化（1回のみ使用可能）
  props.deleteProperty('portal_login_' + token);

  // メンバー情報取得
  var member = getMemberByEmail(tokenData.email);
  if (!member) {
    return { success: false, message: 'メンバー情報が見つかりません' };
  }

  // ロール判定
  var role = determineRole(member, tokenData.email);

  // セッション発行
  var sessionId = createSession(tokenData.email, member.name, role);

  // 最終ログイン日時を更新
  updateLastLogin(tokenData.email);

  console.log('Portal login verified: ' + member.name + ' (role=' + role + ')');

  return {
    success: true,
    sessionId: sessionId,
    member: {
      name: member.name,
      email: tokenData.email,
      term: member.term,
      type: member.type,
      role: role
    }
  };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// セッション管理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * セッションを作成
 * @returns {string} セッションID
 */
function createSession(email, name, role) {
  var sessionId = Utilities.getUuid();
  var cache = CacheService.getScriptCache();
  var sessionData = JSON.stringify({
    email: email,
    name: name,
    role: role,
    createdAt: new Date().toISOString()
  });

  // CacheService: 最大6時間（21600秒）
  cache.put('portal_session_' + sessionId, sessionData, PORTAL_CONFIG.SESSION_HOURS * 3600);

  return sessionId;
}

/**
 * セッションを検証
 * @param {string} sessionId - セッションID
 * @returns {Object|null} セッションデータ（無効時はnull）
 */
function validateSession(sessionId) {
  if (!sessionId) return null;

  var cache = CacheService.getScriptCache();
  var sessionDataStr = cache.get('portal_session_' + sessionId);
  if (!sessionDataStr) return null;

  try {
    return JSON.parse(sessionDataStr);
  } catch (e) {
    return null;
  }
}

/**
 * セッションを破棄（ログアウト）
 * @param {string} sessionId - セッションID
 */
function destroySession(sessionId) {
  if (!sessionId) return;
  var cache = CacheService.getScriptCache();
  cache.remove('portal_session_' + sessionId);
}

/**
 * セッション検証ミドルウェア
 * @param {Object} e - GETリクエスト
 * @returns {Object|null} セッションデータ（無効時はnull）
 */
function requireAuth(e) {
  var sessionId = e.parameter.sessionId || e.parameter.session;
  return validateSession(sessionId);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ロール管理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * メンバーのロールを判定
 * @param {Object} member - メンバーデータ
 * @param {string} email - メールアドレス
 * @returns {string} ロール ('admin', 'leader', 'member')
 */
function determineRole(member, email) {
  // 管理者メールアドレスの場合
  if (PORTAL_CONFIG.ADMIN_EMAILS.indexOf(email.toLowerCase()) >= 0) {
    return 'admin';
  }

  // メンバーマスタの区分で判定
  var type = (member.type || '').toString();
  if (type === '顧問' || type === '管理者') {
    return 'admin';
  }

  // 1期/2期の診断士はリーダー候補
  var term = (member.term || '').toString();
  var cert = (member.cert || '').toString();
  if ((term === '1期' || term === '2期') && cert.indexOf('診断士') >= 0) {
    return 'leader';
  }

  return 'member';
}

/**
 * ロール権限チェック
 * @param {string} role - 現在のロール
 * @param {string} requiredRole - 必要なロール
 * @returns {boolean} 権限があるか
 */
function hasRole(role, requiredRole) {
  var hierarchy = { 'admin': 3, 'leader': 2, 'member': 1 };
  return (hierarchy[role] || 0) >= (hierarchy[requiredRole] || 0);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ヘルパー関数
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * メールアドレスでメンバーを検索
 * @param {string} email - メールアドレス
 * @returns {Object|null} メンバーデータ
 */
function getMemberByEmail(email) {
  if (!email) return null;
  email = email.trim().toLowerCase();

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var memberEmail = (data[i][MEMBER_COLUMNS.EMAIL] || '').toString().trim().toLowerCase();
    if (memberEmail === email) {
      return {
        name: data[i][MEMBER_COLUMNS.NAME],
        term: data[i][MEMBER_COLUMNS.TERM],
        cert: data[i][MEMBER_COLUMNS.CERT],
        type: data[i][MEMBER_COLUMNS.TYPE],
        email: data[i][MEMBER_COLUMNS.EMAIL],
        phone: data[i][MEMBER_COLUMNS.PHONE],
        lineId: data[i][MEMBER_COLUMNS.LINE_ID],
        specialties: data[i][MEMBER_COLUMNS.SPECIALTIES],
        themes: data[i][MEMBER_COLUMNS.THEMES],
        active: data[i][MEMBER_COLUMNS.ACTIVE],
        titles: data[i][MEMBER_COLUMNS.TITLES],
        row: i + 1
      };
    }
  }
  return null;
}

/**
 * 最終ログイン日時を更新
 * メンバーマスタに「最終ログイン」列がある場合に更新
 * @param {string} email - メールアドレス
 */
function updateLastLogin(email) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME);
    if (!sheet) return;

    var data = sheet.getDataRange().getValues();
    var headers = data[0];

    // 「最終ログイン」列を検索
    var loginCol = -1;
    for (var c = 0; c < headers.length; c++) {
      if (headers[c] === '最終ログイン') {
        loginCol = c;
        break;
      }
    }

    if (loginCol < 0) return;  // 列がなければスキップ

    for (var i = 1; i < data.length; i++) {
      var memberEmail = (data[i][MEMBER_COLUMNS.EMAIL] || '').toString().trim().toLowerCase();
      if (memberEmail === email.trim().toLowerCase()) {
        sheet.getRange(i + 1, loginCol + 1).setValue(new Date());
        break;
      }
    }
  } catch (e) {
    console.log('最終ログイン更新スキップ:', e.message);
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ポータルAPI（認証後のデータ取得）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * ダッシュボードデータを取得
 * @param {Object} session - セッションデータ
 * @returns {Object} ダッシュボードデータ
 */
function getPortalDashboard(session) {
  var result = {
    member: {
      name: session.name,
      role: session.role
    },
    upcomingConsultations: [],
    recentNews: [],
    pendingTasks: []
  };

  // 直近の相談予定（確定済み）
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet && sheet.getLastRow() > 1) {
    var data = sheet.getDataRange().getValues();
    var now = new Date();

    for (var i = 1; i < data.length; i++) {
      var status = data[i][COLUMNS.STATUS];
      if (status !== STATUS.CONFIRMED) continue;

      var confDate = data[i][COLUMNS.CONFIRMED_DATE];
      if (!confDate) continue;

      var dateObj = confDate instanceof Date ? confDate : new Date(confDate);
      if (dateObj < now) continue;  // 過去は除外

      result.upcomingConsultations.push({
        id: data[i][COLUMNS.ID],
        company: data[i][COLUMNS.COMPANY],
        date: Utilities.formatDate(dateObj, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
        method: data[i][COLUMNS.METHOD],
        leader: data[i][COLUMNS.LEADER],
        theme: data[i][COLUMNS.THEME]
      });
    }

    // 日付順ソート（近い順）
    result.upcomingConsultations.sort(function(a, b) {
      return new Date(a.date) - new Date(b.date);
    });
    result.upcomingConsultations = result.upcomingConsultations.slice(0, 5);
  }

  // 最新のお知らせ
  try {
    result.recentNews = getLatestNews().slice(0, 3);
  } catch (e) {}

  // 未対応タスク（リーダー/管理者向け）
  if (hasRole(session.role, 'leader')) {
    try {
      var reportSheet = ss.getSheetByName(CONFIG.REPORT_SHEET_NAME);
      if (reportSheet && reportSheet.getLastRow() > 1) {
        var reportData = reportSheet.getDataRange().getValues();
        for (var j = 1; j < reportData.length; j++) {
          var rStatus = reportData[j][REPORT_COLUMNS.STATUS];
          if (rStatus === REPORT_STATUS.REQUESTED || rStatus === REPORT_STATUS.OVERDUE) {
            var leaderEmail = (reportData[j][REPORT_COLUMNS.LEADER_EMAIL] || '').toLowerCase();
            if (session.role === 'admin' || leaderEmail === session.email.toLowerCase()) {
              result.pendingTasks.push({
                type: 'report',
                label: 'レポート提出待ち',
                applicationId: reportData[j][REPORT_COLUMNS.APP_ID],
                company: reportData[j][REPORT_COLUMNS.COMPANY],
                deadline: reportData[j][REPORT_COLUMNS.DEADLINE] instanceof Date
                  ? Utilities.formatDate(reportData[j][REPORT_COLUMNS.DEADLINE], 'Asia/Tokyo', 'yyyy/MM/dd')
                  : '',
                status: rStatus
              });
            }
          }
        }
      }
    } catch (e) {}
  }

  return result;
}

/**
 * 自分のシフト（参加可能日）を取得/更新
 * @param {Object} session - セッションデータ
 * @param {string} yearMonth - 'YYYY/MM' 形式（オプション）
 * @returns {Object} シフトデータ
 */
function getMyShifts(session, yearMonth) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var schedSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!schedSheet) return { success: false, message: '日程設定シートが見つかりません' };

  var data = schedSheet.getDataRange().getValues();
  var shifts = [];
  var memberName = session.name;

  // 対象月フィルタ
  var targetYear = null;
  var targetMonth = null;
  if (yearMonth) {
    var parts = yearMonth.split('/');
    targetYear = parseInt(parts[0]);
    targetMonth = parseInt(parts[1]);
  }

  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
    if (!dateVal) continue;

    var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
    if (isNaN(date.getTime())) continue;

    // 月フィルタ
    if (targetYear && (date.getFullYear() !== targetYear || (date.getMonth() + 1) !== targetMonth)) continue;

    var members = (data[i][SCHEDULE_COLUMNS.MEMBERS] || '').toString();
    var isParticipating = members.indexOf(memberName) >= 0;

    shifts.push({
      row: i + 1,
      date: Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd'),
      dayOfWeek: ['日','月','火','水','木','金','土'][date.getDay()],
      time: data[i][SCHEDULE_COLUMNS.TIME] instanceof Date
        ? Utilities.formatDate(data[i][SCHEDULE_COLUMNS.TIME], 'Asia/Tokyo', 'HH:mm')
        : String(data[i][SCHEDULE_COLUMNS.TIME]),
      method: data[i][SCHEDULE_COLUMNS.METHOD] || '',
      bookingStatus: data[i][SCHEDULE_COLUMNS.BOOKING_STATUS] || '',
      participating: isParticipating,
      score: data[i][SCHEDULE_COLUMNS.SCORE] || 0,
      bookable: data[i][SCHEDULE_COLUMNS.BOOKABLE] || ''
    });
  }

  return { success: true, shifts: shifts, memberName: memberName };
}

/**
 * 自分の担当案件一覧を取得
 * @param {Object} session - セッションデータ
 * @returns {Object} 案件リスト
 */
function getMyCases(session) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return { success: false, message: 'シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  var cases = [];
  var memberName = session.name;

  for (var i = 1; i < data.length; i++) {
    var leader = data[i][COLUMNS.LEADER] || '';
    var staff = data[i][COLUMNS.STAFF] || '';

    // 自分がリーダーまたは担当者の案件
    var isMyCase = leader === memberName || staff.indexOf(memberName) >= 0;
    // 管理者は全案件
    if (session.role === 'admin') isMyCase = true;

    if (!isMyCase) continue;

    cases.push({
      id: data[i][COLUMNS.ID],
      company: data[i][COLUMNS.COMPANY],
      name: data[i][COLUMNS.NAME],
      industry: data[i][COLUMNS.INDUSTRY],
      theme: data[i][COLUMNS.THEME],
      status: data[i][COLUMNS.STATUS],
      confirmedDate: data[i][COLUMNS.CONFIRMED_DATE] instanceof Date
        ? Utilities.formatDate(data[i][COLUMNS.CONFIRMED_DATE], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
        : String(data[i][COLUMNS.CONFIRMED_DATE] || ''),
      method: data[i][COLUMNS.METHOD],
      leader: leader,
      reportStatus: data[i][COLUMNS.REPORT_STATUS] || '',
      transcriptStatus: data[i][COLUMNS.TRANSCRIPT_STATUS] || ''
    });
  }

  // 日付降順
  cases.reverse();

  return { success: true, cases: cases, total: cases.length };
}
