/**
 * メイン処理ファイル（拡張版）
 * Webフォームからの送信を受け付け、各種処理を実行
 * 同意書ルーティング、企業URL、当日受付対応を追加
 */

/**
 * POSTリクエスト処理（フォーム送信時）
 */
function doPost(e) {
  try {
    // フォームデータ取得
    const data = parseFormData(e);

    // 申込ID生成
    const applicationId = generateApplicationId();
    data.id = applicationId;

    // 当日受付判定（備考欄の9999コード or LP側のwalkInFlag）
    const isWalkIn = (data.remarks && data.remarks.includes(CONFIG.WALK_IN_CODE)) || data.walkInFlag === 'TRUE';

    // スプレッドシートに記録（拡張版: 企業URL、当日受付フラグ付き）
    const rowIndex = saveToSpreadsheet(data, isWalkIn);

    // 予約した日程を「予約済み」に更新
    if (data.date1 && data.time) {
      const dateStr = convertJapaneseDateToSlash(data.date1);
      if (dateStr) {
        markAsBooked(dateStr, data.time);
      }
    }

    // 同意書トークン生成
    const consentToken = generateNdaToken(applicationId, data.email);
    const consentUrl = CONFIG.CONSENT.WEB_APP_URL !== 'ここにGASデプロイURLを入力'
      ? CONFIG.CONSENT.WEB_APP_URL + '?action=nda&token=' + consentToken
      : '';

    // 申込者に自動返信メール送信
    if (isWalkIn) {
      // 当日受付用メール
      sendWalkInConfirmationEmail(data, consentUrl);
    } else {
      // 通常の確認メール（ヒアリングシート添付 + 同意書リンク）
      sendConfirmationEmail(data, consentUrl);
    }

    // 担当者にメール通知
    sendAdminNotification(data);

    // LINE通知
    sendLineNotification(data);

    // Notion連携（有効な場合）
    if (CONFIG.NOTION.ENABLED) {
      createNotionEntry(data);
    }

    // 成功レスポンス
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'お申し込みを受け付けました',
        applicationId: applicationId
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Error in doPost:', error);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'エラーが発生しました。お手数ですがお問い合わせください。',
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * GETリクエスト処理（日程取得API + 同意書ページ）
 */
function doGet(e) {
  try {
    const action = e.parameter.action || 'status';

    // 同意書ページ
    if (action === 'nda') {
      return generateNdaPage(e);
    }

    // 同意処理（POSTの代替としてGETで受ける場合）
    if (action === 'nda-submit') {
      const result = processNdaConsent(e);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 【一時ルート】日程データクリーンアップ（使用後削除すること）
    if (action === 'cleanup-schedule') {
      try {
        cleanupAndRegenerateMonth(2026, 3);
        cleanupAndRegenerateMonth(2026, 4);
        return ContentService
          .createTextOutput(JSON.stringify({
            success: true,
            message: '3月・4月の日程データをクリーンアップし再生成しました'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            error: err.toString()
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // お知らせ管理ページ
    if (action === 'news-admin') {
      return generateNewsAdminPage();
    }

    // お知らせ取得API（LP用）
    if (action === 'news') {
      var news = getLatestNews();
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          news: news
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // お知らせ追加
    if (action === 'news-add') {
      addNews(e.parameter.date, e.parameter.content);
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=news-admin";</script>'
      );
    }

    // お知らせ表示/非表示切替
    if (action === 'news-toggle') {
      toggleNewsVisibility(e.parameter.row);
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=news-admin";</script>'
      );
    }

    // お知らせ削除
    if (action === 'news-delete') {
      deleteNews(e.parameter.row);
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=news-admin";</script>'
      );
    }

    // 日程取得
    if (action === 'schedule') {
      const method = e.parameter.method || 'both';
      const schedule = getAvailableSchedule(method);

      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          schedule: schedule
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // デフォルト: ステータス確認
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'OK',
        message: '無料経営相談予約システム API',
        endpoints: {
          'GET ?action=schedule&method=visit': '対面可能な日程を取得',
          'GET ?action=schedule&method=zoom': 'オンライン可能な日程を取得',
          'GET ?action=nda&token=xxx': '同意書確認ページを表示',
          'GET ?action=news': 'お知らせ取得（LP用）',
          'GET ?action=news-admin': 'お知らせ管理ページ',
          'POST': '予約申込'
        }
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error('Error in doGet:', error);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * フォームデータをパース（JSON形式とフォーム形式の両方に対応・拡張版）
 */
function parseFormData(e) {
  let params;

  if (e.postData && e.postData.contents) {
    try {
      params = JSON.parse(e.postData.contents);
    } catch (jsonError) {
      params = e.parameter || {};
    }
  } else {
    params = e.parameter || {};
  }

  return {
    timestamp: new Date(),
    name: params.name || '',
    company: params.company || '',
    email: params.email || '',
    phone: params.phone || '',
    position: params.position || '',
    industry: params.industry || '',
    theme: params.theme || params.topic || '',
    content: params.content || '',
    date1: params.date1 || params.date || '',
    date2: params.date2 || '',
    time: params.time || '',
    method: params.method || '',
    notes: params.notes || '',
    remarks: params.remarks || '',
    companyUrl: params.companyUrl || '',
    walkInFlag: params.walkInFlag || 'FALSE'
  };
}

/**
 * 日本語日付をスラッシュ形式に変換
 */
function convertJapaneseDateToSlash(jaDate) {
  if (!jaDate) return null;

  const match = jaDate.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (match) {
    const year = match[1];
    const month = String(match[2]).padStart(2, '0');
    const day = String(match[3]).padStart(2, '0');
    return `${year}/${month}/${day}`;
  }

  if (/^\d{4}\/\d{2}\/\d{2}$/.test(jaDate)) {
    return jaDate;
  }

  return null;
}

/**
 * 申込ID生成（YYYYMMDD-001形式）
 */
function generateApplicationId() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  const data = sheet.getDataRange().getValues();
  let todayCount = 0;

  for (let i = 1; i < data.length; i++) {
    const id = data[i][COLUMNS.ID];
    if (id && id.toString().startsWith(dateStr)) {
      todayCount++;
    }
  }

  const seq = String(todayCount + 1).padStart(3, '0');
  return `${dateStr}-${seq}`;
}

/**
 * 申込者への確認メール送信（同意書確認リンク付き）
 */
function sendConfirmationEmail(data, consentUrl) {
  const subject = '【受付完了】無料経営相談のお申し込みありがとうございます';
  const body = getConfirmationEmailBody(data, consentUrl);

  // ヒアリングシートを添付
  let attachments = [];
  try {
    if (CONFIG.HEARING_SHEET_FILE_ID && CONFIG.HEARING_SHEET_FILE_ID !== 'ここにヒアリングシートのファイルIDを入力') {
      const file = DriveApp.getFileById(CONFIG.HEARING_SHEET_FILE_ID);
      attachments.push(file.getAs(MimeType.PDF));
    }
  } catch (e) {
    console.log('ヒアリングシートの添付に失敗:', e);
  }

  const options = {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO,
    attachments: attachments
  };

  GmailApp.sendEmail(data.email, subject, body, options);
}

/**
 * 当日受付用確認メール送信
 */
function sendWalkInConfirmationEmail(data, consentUrl) {
  const subject = '【当日受付】無料経営相談へようこそ';
  const body = getWalkInConfirmationEmailBody(data, consentUrl);

  GmailApp.sendEmail(data.email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });
}

/**
 * 担当者への通知メール送信
 */
function sendAdminNotification(data) {
  const subject = `【新規申込】${data.name}様 - ${data.theme}`;
  const body = getAdminNotificationBody(data);

  CONFIG.ADMIN_EMAILS.forEach(email => {
    GmailApp.sendEmail(email, subject, body, {
      name: CONFIG.SENDER_NAME
    });
  });
}

/**
 * 予約確定メール送信
 */
function sendConfirmedEmail(data) {
  const subject = '【予約確定】無料経営相談のご予約が確定しました';
  const body = getConfirmedEmailBody(data);

  GmailApp.sendEmail(data.email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });
}

