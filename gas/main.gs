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
      // 通常の確認メール（同意書リンク付き）
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

    // 回答集計シート生成（管理用）
    if (action === 'generate-summary') {
      generateSummarySheet();
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          message: '回答集計シートを生成しました'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ポーリング状況確認（管理用）
    if (action === 'polling-status') {
      const result = getPollingStatus();
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // PDF更新（管理用: URLからPDFを取得してDriveを更新）
    if (action === 'update-pdf') {
      const pdfUrl = e.parameter.url;
      if (!pdfUrl) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'urlパラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const result = updateConsentPdfFromUrl(pdfUrl);
      return ContentService
        .createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 日程枠追加（管理用）
    if (action === 'add-schedule') {
      var date = e.parameter.date;
      var time = e.parameter.time;
      var method = e.parameter.method || 'オンライン';
      if (!date || !time) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'date, timeパラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      var schedSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
      var newRow = [date, time, '○', method, '', '空き', '', '', 0, 'FALSE', '○'];
      schedSheet.appendRow(newRow);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, message: date + ' ' + time + ' (' + method + ') を追加しました' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // メンバー一括登録（管理用）
    if (action === 'register-members') {
      var regYear = parseInt(e.parameter.year) || new Date().getFullYear();
      var regMonth = parseInt(e.parameter.month) || (new Date().getMonth() + 1);
      registerAllMembersForMonth(regYear, regMonth);
      // 回答集計も再生成
      generateSummarySheet();
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          message: regYear + '年' + regMonth + '月の全メンバー登録＋回答集計再生成完了'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 回答集計→日程設定 一括同期（管理用）
    if (action === 'sync-summary') {
      var syncResult = syncAllSummaryToSchedule();
      return ContentService
        .createTextOutput(JSON.stringify(syncResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 回答集計onEditトリガー設定（管理用）
    if (action === 'setup-summary-trigger') {
      var triggerResult = setupSummaryEditTrigger();
      return ContentService
        .createTextOutput(JSON.stringify(triggerResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 場所列マイグレーション（管理用）
    if (action === 'migrate-location') {
      var migrateResult = migrateAddLocationColumn();
      return ContentService
        .createTextOutput(JSON.stringify(migrateResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // メンバーマスタセットアップ（管理用）
    if (action === 'setup-members') {
      setupMemberSheet();
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          message: 'メンバーマスタシートのセットアップが完了しました'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // アンケートシステム
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // アンケートページ表示
    if (action === 'survey') {
      return generateSurveyPage(e);
    }

    // アンケートシートセットアップ（管理用）
    if (action === 'setup-survey') {
      var surveyResult = setupSurveySheet();
      return ContentService
        .createTextOutput(JSON.stringify(surveyResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // アンケートトリガーセットアップ（管理用）
    if (action === 'setup-survey-trigger') {
      var triggerResult = setupSurveyTrigger();
      return ContentService
        .createTextOutput(JSON.stringify(triggerResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // オブザーバー専用ページ
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // オブザーバー専用ページ表示
    if (action === 'observer') {
      return generateObserverPage(e);
    }

    // オブザーバーNDAシートセットアップ（管理用）
    if (action === 'setup-observer-nda') {
      var observerNdaResult = setupObserverNdaSheet();
      return ContentService
        .createTextOutput(JSON.stringify(observerNdaResult))
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
          'GET ?action=survey&token=xxx': 'アンケートページ',
          'GET ?action=observer': 'オブザーバー専用ページ',
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
 * 日本語日付をスラッシュ形式に変換（時間帯を保持）
 * "2026年3月7日（土） 14:00" → "2026/03/07 14:00"
 * "2026年3月7日（土）" → "2026/03/07"
 */
function convertJapaneseDateToSlashWithTime(jaDate) {
  if (!jaDate) return null;
  var str = String(jaDate);

  var dateMatch = str.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (!dateMatch) {
    // yyyy/MM/dd HH:mm形式ならそのまま
    if (/^\d{4}\/\d{2}\/\d{2}/.test(str)) return str;
    return null;
  }

  var year = dateMatch[1];
  var month = String(dateMatch[2]).padStart(2, '0');
  var day = String(dateMatch[3]).padStart(2, '0');
  var datePart = year + '/' + month + '/' + day;

  // 時間帯を抽出（例: 14:00, 19:00 etc）
  var timeMatch = str.match(/(\d{1,2}:\d{2})/);
  if (timeMatch) {
    return datePart + ' ' + timeMatch[1];
  }

  return datePart;
}

/**
 * 確定日時を取得（K列 + 日程設定シートの時間帯から補完）
 * @param {number} rowIndex - 行番号（1-based）
 * @param {Object} sheet - 予約管理シート
 * @returns {string|null} 確定日時文字列
 */
function resolveConfirmedDateTime(rowIndex, sheet) {
  var date1 = sheet.getRange(rowIndex, COLUMNS.DATE1 + 1).getValue();
  if (!date1) return null;

  var dateStr = String(date1);
  var result = convertJapaneseDateToSlashWithTime(dateStr);

  // 時間が含まれていればそのまま返す
  if (result && /\d{1,2}:\d{2}/.test(result)) return result;

  // 時間がない場合、日程設定シートから予約済みの時間帯を検索
  var slashDate = convertJapaneseDateToSlash(dateStr);
  if (slashDate) {
    try {
      var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      var schedSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
      if (schedSheet) {
        var schedData = schedSheet.getDataRange().getValues();
        for (var i = 1; i < schedData.length; i++) {
          var sDate = schedData[i][SCHEDULE_COLUMNS.DATE];
          var sDateStr = sDate instanceof Date
            ? Utilities.formatDate(sDate, 'Asia/Tokyo', 'yyyy/MM/dd')
            : String(sDate);
          if (sDateStr === slashDate && schedData[i][SCHEDULE_COLUMNS.BOOKING_STATUS] === '予約済み') {
            var sTime = schedData[i][SCHEDULE_COLUMNS.TIME];
            var timeStr = sTime instanceof Date
              ? Utilities.formatDate(sTime, 'Asia/Tokyo', 'HH:mm')
              : String(sTime);
            return slashDate + ' ' + timeStr;
          }
        }
      }
    } catch (e) { /* fallback */ }
  }

  return result || dateStr;
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

  const options = {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
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

