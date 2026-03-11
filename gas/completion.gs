/**
 * 相談完了確認システム
 * 確定日時+4時間後にリーダーへ完了確認メールを送信
 * リーダーの回答に応じてステータス更新・アンケート配信を制御
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// トークン管理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 完了確認トークンを生成
 * @param {string} applicationId - 申込ID
 * @param {string} leaderName - リーダー名
 * @param {number} rowIndex - 行番号
 * @returns {string} トークン
 */
function generateCompletionToken(applicationId, leaderName, rowIndex) {
  var token = Utilities.getUuid();
  var props = PropertiesService.getScriptProperties();
  var tokenData = JSON.stringify({
    applicationId: applicationId,
    leaderName: leaderName,
    rowIndex: rowIndex,
    createdAt: new Date().toISOString()
  });
  props.setProperty('completion_token_' + token, tokenData);
  return token;
}

/**
 * 完了確認トークンを検証
 * @param {string} token - トークン
 * @returns {Object|null} トークンデータ
 */
function validateCompletionToken(token) {
  if (!token) return null;
  var props = PropertiesService.getScriptProperties();
  var tokenDataStr = props.getProperty('completion_token_' + token);
  if (!tokenDataStr) return null;
  try {
    return JSON.parse(tokenDataStr);
  } catch (e) {
    return null;
  }
}

/**
 * 完了確認トークンを無効化
 * @param {string} token - トークン
 */
function invalidateCompletionToken(token) {
  if (!token) return;
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('completion_token_' + token);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 毎時チェック＆メール送信
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 完了確認メールを送信（トリガーから毎時実行）
 * 条件:
 *   1. ステータスが「確定」
 *   2. メールアドレスあり
 *   3. 確定日時 + DELAY_HOURS が経過
 *   4. 備考に「完了確認メール送信済」を含まない
 *   5. リーダー列（Y列）に名前あり
 */
function sendCompletionConfirmEmails() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var now = new Date();
  var data = sheet.getDataRange().getValues();
  var delayMs = CONFIG.COMPLETION.DELAY_HOURS * 60 * 60 * 1000;

  for (var i = 1; i < data.length; i++) {
    var status = data[i][COLUMNS.STATUS];
    var email = data[i][COLUMNS.EMAIL];
    var confirmedDate = data[i][COLUMNS.CONFIRMED_DATE];
    var notes = data[i][COLUMNS.NOTES] || '';
    var leaderName = data[i][COLUMNS.LEADER];
    var appId = data[i][COLUMNS.ID];

    // 条件チェック
    if (status !== STATUS.CONFIRMED) continue;
    if (!email) continue;
    if (!confirmedDate) continue;
    if (notes.indexOf('完了確認メール送信済') >= 0) continue;
    if (!leaderName) continue;

    // 確定日時 + DELAY_HOURS が経過しているか
    var confirmDate;
    if (confirmedDate instanceof Date || (typeof confirmedDate === 'object' && confirmedDate && typeof confirmedDate.getTime === 'function')) {
      confirmDate = confirmedDate;
    } else {
      confirmDate = new Date(confirmedDate);
    }
    if (isNaN(confirmDate.getTime())) continue;

    var sendTime = new Date(confirmDate.getTime() + delayMs);
    if (now < sendTime) continue;

    // リーダーのメールアドレスを取得
    var leaderMember = getMemberByName(leaderName);
    if (!leaderMember || !leaderMember.email) {
      console.log('完了確認: リーダー "' + leaderName + '" のメール未設定 (行' + (i + 1) + ')');
      continue;
    }

    // 行データ取得
    var rowIndex = i + 1;
    var rowData = getRowData(rowIndex);

    // トークン生成 → 確認ページURL構築
    var token = generateCompletionToken(appId, leaderName, rowIndex);
    var confirmUrl = CONFIG.CONSENT.WEB_APP_URL + '?action=completion&token=' + token;

    // メール送信
    sendCompletionConfirmEmail(leaderMember.email, leaderName, rowData, confirmUrl);

    // 備考に送信済みマーク
    var currentNotes = sheet.getRange(rowIndex, COLUMNS.NOTES + 1).getValue() || '';
    sheet.getRange(rowIndex, COLUMNS.NOTES + 1).setValue(
      currentNotes + (currentNotes ? '\n' : '') + '完了確認メール送信済(' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ')'
    );

    console.log('完了確認メール送信: ' + appId + ' → ' + leaderName + ' (' + leaderMember.email + ')');
  }
}

/**
 * リーダー宛の完了確認メール送信
 * @param {string} email - リーダーのメールアドレス
 * @param {string} leaderName - リーダー名
 * @param {Object} data - 行データ
 * @param {string} confirmUrl - 確認ページURL
 */
function sendCompletionConfirmEmail(email, leaderName, data, confirmUrl) {
  var subject = '【相談完了確認】' + (data.company || '') + '様 - ' + (data.id || '');

  var body = leaderName + ' 様\n\n' +
    'お疲れ様です。\n' +
    '下記の経営相談について、完了確認をお願いいたします。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 相談概要\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '申込ID　：' + (data.id || '') + '\n' +
    '相談日時：' + (data.confirmedDate || '') + '\n' +
    '相談者　：' + (data.name || '') + ' 様\n' +
    '企業名　：' + (data.company || '') + '\n' +
    '相談方法：' + (data.method || '') + '\n' +
    'テーマ　：' + (data.theme || '') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '以下のリンクから相談が完了したかどうかを回答してください。\n\n' +
    '■ 完了確認ページ\n' +
    confirmUrl + '\n\n' +
    '「完了」を選択すると、相談者にアンケートメールが自動送信されます。\n' +
    '「未完了」を選択すると、管理者に通知されます。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    CONFIG.ORG.NAME + '\n' +
    'Email: ' + CONFIG.ORG.EMAIL + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

  GmailApp.sendEmail(email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// フォーム送信処理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 完了確認フォームの送信処理（Webページから呼ばれる）
 * @param {Object} formData - { token, result, comment }
 * @returns {Object} { success, message }
 */
function submitCompletionConfirm(formData) {
  try {
    // トークン検証
    var tokenData = validateCompletionToken(formData.token);
    if (!tokenData) {
      return { success: false, message: '無効なリンクです。既に回答済みか、リンクが期限切れです。' };
    }

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    var rowIndex = tokenData.rowIndex;
    var appId = tokenData.applicationId;
    var leaderName = tokenData.leaderName;
    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

    if (formData.result === 'completed') {
      // ━━━ 完了 ━━━

      // ステータスを「完了」に更新
      sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(STATUS.COMPLETED);

      // リーダー履歴更新（予定→完了）＆欠席時フォールバック再選定
      try {
        autoSelectLeaderOnComplete(rowIndex);
      } catch (leaderError) {
        console.error('完了確認: リーダー選定エラー:', leaderError);
      }

      // 備考に記録
      var currentNotes = sheet.getRange(rowIndex, COLUMNS.NOTES + 1).getValue() || '';
      sheet.getRange(rowIndex, COLUMNS.NOTES + 1).setValue(
        currentNotes + (currentNotes ? '\n' : '') + '完了確認済(' + nowStr + ') ' + leaderName + ' 回答'
      );

      console.log('完了確認: 完了 - ' + appId + ' (リーダー: ' + leaderName + ')');

    } else if (formData.result === 'incomplete') {
      // ━━━ 未完了 ━━━

      // 管理者にメール通知
      var comment = formData.comment || '（コメントなし）';
      var adminSubject = '【未完了報告】' + appId + ' - リーダー回答';
      var adminBody = 'リーダー「' + leaderName + '」が「未完了」と回答しました。\n\n' +
        '申込ID：' + appId + '\n' +
        'リーダー：' + leaderName + '\n' +
        'コメント：' + comment + '\n\n' +
        '手動で対応してください。';

      CONFIG.ADMIN_EMAILS.forEach(function(adminEmail) {
        GmailApp.sendEmail(adminEmail, adminSubject, adminBody, {
          name: CONFIG.SENDER_NAME
        });
      });

      // 備考に記録
      var currentNotes2 = sheet.getRange(rowIndex, COLUMNS.NOTES + 1).getValue() || '';
      sheet.getRange(rowIndex, COLUMNS.NOTES + 1).setValue(
        currentNotes2 + (currentNotes2 ? '\n' : '') + '未完了報告(' + nowStr + ') ' + leaderName + ' 回答: ' + comment
      );

      console.log('完了確認: 未完了 - ' + appId + ' (リーダー: ' + leaderName + ', コメント: ' + comment + ')');
    }

    // トークン無効化
    invalidateCompletionToken(formData.token);

    return { success: true, message: '回答を受け付けました。ありがとうございます。' };

  } catch (error) {
    console.error('完了確認送信エラー:', error);
    return { success: false, message: 'エラーが発生しました: ' + error.toString() };
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 確認ページ生成
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 完了確認ページを生成
 * @param {Object} e - GETリクエストパラメータ
 * @returns {HtmlOutput} HTMLページ
 */
function generateCompletionConfirmPage(e) {
  var token = e.parameter.token;
  var tokenData = validateCompletionToken(token);

  if (!tokenData) {
    return HtmlService.createHtmlOutput(
      '<html><body><h2>無効なリンクです</h2><p>この完了確認リンクは無効または既に回答済みです。</p></body></html>'
    )
      .setTitle('エラー - 完了確認')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 行データを取得
  var rowData = getRowData(tokenData.rowIndex);

  var html = getCompletionConfirmPageHtml(tokenData, rowData, token);
  return HtmlService.createHtmlOutput(html)
    .setTitle('相談完了確認 - 関西学院大学 中小企業経営診断研究会無料経営相談分科会')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// トリガー設定
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 完了確認メール送信トリガーのセットアップ（1時間おき）
 */
function setupCompletionTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendCompletionConfirmEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('sendCompletionConfirmEmails')
    .timeBased()
    .everyHours(1)
    .create();

  return { success: true, message: '完了確認メール送信トリガーをセットアップしました（1時間おき）' };
}
