/**
 * アンケートシステム
 * 相談後アンケートの表示・回答保存・自動配信
 */

/**
 * アンケートシートのセットアップ
 */
function setupSurveySheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SURVEY_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SURVEY_SHEET_NAME);
  }

  var headers = [
    '回答日時', '申込ID', '氏名', '企業名',
    'Q1:きっかけ', 'Q1:SNS種別',
    'Q2:手続き(5段階)', 'Q2:コメント',
    'Q3:感想',
    'Q4:時間',
    'Q5:説明(5段階)', 'Q6:課題解決(5段階)', 'Q7:対応(5段階)', 'Q8:アドバイス(5段階)',
    'Q9:また受けたい(5段階)', 'Q9:理由',
    'Q10:勧めたい(5段階)', 'Q10:理由',
    'Q11:レポート希望'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#0F2350')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  return { success: true, message: 'アンケートシートをセットアップしました' };
}

/**
 * アンケートページを生成
 */
function generateSurveyPage(e) {
  var token = e.parameter.token;
  var tokenData = validateSurveyToken(token);

  if (!tokenData) {
    return HtmlService.createHtmlOutput('<html><body><h2>無効なリンクです</h2><p>このアンケートリンクは無効または期限切れです。</p></body></html>')
      .setTitle('エラー - アンケート')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var html = getSurveyPageHtml(tokenData);
  return HtmlService.createHtmlOutput(html)
    .setTitle('相談後アンケート - 関西学院大学 中小企業経営診断研究会')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * アンケート回答を保存
 */
function submitSurveyResponse(formData) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SURVEY_SHEET_NAME);

    if (!sheet) {
      setupSurveySheet();
      sheet = ss.getSheetByName(CONFIG.SURVEY_SHEET_NAME);
    }

    var row = [
      new Date(),
      formData.applicationId || '',
      formData.name || '',
      formData.company || '',
      formData.q1 || '',
      formData.q1Sns || '',
      formData.q2 || '',
      formData.q2Comment || '',
      formData.q3 || '',
      formData.q4 || '',
      formData.q5 || '',
      formData.q6 || '',
      formData.q7 || '',
      formData.q8 || '',
      formData.q9 || '',
      formData.q9Reason || '',
      formData.q10 || '',
      formData.q10Reason || '',
      formData.q11 || ''
    ];

    sheet.appendRow(row);

    // 管理者に通知
    notifySurveyResponse(formData);

    return { success: true, message: 'アンケートのご回答ありがとうございました' };
  } catch (error) {
    console.error('アンケート保存エラー:', error);
    return { success: false, message: 'エラーが発生しました: ' + error.toString() };
  }
}

/**
 * アンケート回答の管理者通知
 */
function notifySurveyResponse(formData) {
  var subject = '【アンケート回答】' + (formData.name || '') + '様 - ' + (formData.company || '');
  var body = 'アンケートの回答がありました。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '申込ID：' + (formData.applicationId || '') + '\n' +
    'お名前：' + (formData.name || '') + '\n' +
    '企業名：' + (formData.company || '') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Q2 手続き：' + (formData.q2 || '') + '\n' +
    'Q4 時間：' + (formData.q4 || '') + '\n' +
    'Q9 また受けたい：' + (formData.q9 || '') + '\n' +
    'Q10 勧めたい：' + (formData.q10 || '') + '\n' +
    'Q11 レポート希望：' + (formData.q11 || '') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    'スプレッドシートで詳細をご確認ください。';

  CONFIG.ADMIN_EMAILS.forEach(function(email) {
    GmailApp.sendEmail(email, subject, body, { name: CONFIG.SENDER_NAME });
  });
}

/**
 * アンケートトークンを生成
 */
function generateSurveyToken(applicationId, name, company, ndaFileId) {
  var token = Utilities.getUuid();
  var props = PropertiesService.getScriptProperties();
  var tokenData = JSON.stringify({
    applicationId: applicationId,
    name: name,
    company: company,
    ndaFileId: ndaFileId || '',
    createdAt: new Date().toISOString()
  });
  props.setProperty('survey_token_' + token, tokenData);
  return token;
}

/**
 * アンケートトークンを検証
 */
function validateSurveyToken(token) {
  if (!token) return null;
  var props = PropertiesService.getScriptProperties();
  var tokenDataStr = props.getProperty('survey_token_' + token);
  if (!tokenDataStr) return null;
  try {
    return JSON.parse(tokenDataStr);
  } catch (e) {
    return null;
  }
}

/**
 * 相談終了後にアンケートメールを送信
 * （トリガーから呼ばれる or 手動実行）
 */
function sendSurveyEmails() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);

  if (!sheet || !scheduleSheet) return;

  var now = new Date();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var status = data[i][COLUMNS.STATUS];
    var confirmedDate = data[i][COLUMNS.CONFIRMED_DATE];
    var email = data[i][COLUMNS.EMAIL];
    var name = data[i][COLUMNS.NAME];
    var company = data[i][COLUMNS.COMPANY];
    var appId = data[i][COLUMNS.ID];
    var notes = data[i][COLUMNS.NOTES] || '';

    // 確定済みかつアンケート未送信のもの
    if (status === '完了' && email && !notes.includes('アンケート送信済')) {
      // 確定日時から終了時刻を推定し、2時間後を過ぎているか確認
      if (confirmedDate) {
        var confirmDate = new Date(confirmedDate);
        var sendTime = new Date(confirmDate.getTime() + CONFIG.SURVEY.DELAY_HOURS * 60 * 60 * 1000);

        if (now >= sendTime) {
          // アンケートトークン生成
          var surveyToken = generateSurveyToken(appId, name, company, '');
          var surveyUrl = CONFIG.CONSENT.WEB_APP_URL + '?action=survey&token=' + surveyToken;

          // メール送信
          sendSurveyEmail(email, name, company, appId, surveyUrl);

          // 送信済みマーク
          var currentNotes = sheet.getRange(i + 1, COLUMNS.NOTES + 1).getValue() || '';
          sheet.getRange(i + 1, COLUMNS.NOTES + 1).setValue(
            currentNotes + (currentNotes ? '\n' : '') + 'アンケート送信済(' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ')'
          );
        }
      }
    }
  }
}

/**
 * アンケートメール送信
 */
function sendSurveyEmail(email, name, company, appId, surveyUrl) {
  var subject = '【アンケートのお願い】無料経営相談のご感想をお聞かせください';
  var body = name + ' 様\n\n' +
    '先日は、関西学院大学 中小企業経営診断研究会の無料経営相談をご利用いただき、\n' +
    '誠にありがとうございました。\n\n' +
    '今後のサービス向上のため、簡単なアンケートにご協力をお願いいたします。\n' +
    '（所要時間：約3分）\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ アンケート回答ページ\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    surveyUrl + '\n\n' +
    'ご不明な点がございましたら、お気軽にお問い合わせください。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    CONFIG.ORG.NAME + '\n' +
    'Email: ' + CONFIG.ORG.EMAIL + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

  GmailApp.sendEmail(email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });
}

/**
 * アンケート送信トリガーのセットアップ
 */
function setupSurveyTrigger() {
  // 既存のアンケートトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendSurveyEmails') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 1時間おきにチェック
  ScriptApp.newTrigger('sendSurveyEmails')
    .timeBased()
    .everyHours(1)
    .create();

  return { success: true, message: 'アンケート送信トリガーをセットアップしました' };
}
