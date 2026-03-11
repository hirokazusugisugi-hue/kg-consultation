/**
 * 相談者同意書：確認・同意処理
 * ※相談者が署名する「相談同意書」の処理
 * ※オブザーバーのNDA（秘密保持誓約書）は observer.gs で処理
 * GAS doGet()で動的生成する同意書ページ、同意処理、トークン管理
 */

/**
 * NDA同意用トークンを生成し、スプレッドシートに保存
 * @param {string} applicationId - 申込ID
 * @param {string} email - メールアドレス
 * @returns {string} トークン文字列
 */
function generateNdaToken(applicationId, email) {
  const token = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  const tokenData = JSON.stringify({
    applicationId: applicationId,
    email: email,
    createdAt: new Date().toISOString()
  });
  props.setProperty('nda_token_' + token, tokenData);
  return token;
}

/**
 * NDAトークンを検証
 * @param {string} token - トークン文字列
 * @returns {Object|null} トークンデータ or null
 */
function validateNdaToken(token) {
  if (!token) return null;

  const props = PropertiesService.getScriptProperties();
  const tokenDataStr = props.getProperty('nda_token_' + token);

  if (!tokenDataStr) return null;

  try {
    return JSON.parse(tokenDataStr);
  } catch (e) {
    return null;
  }
}

/**
 * 同意書ページをHTML出力
 * @param {Object} e - GETリクエストパラメータ
 * @returns {HtmlOutput} 同意書ページ
 */
function generateNdaPage(e) {
  const token = e.parameter.token;
  const tokenData = validateNdaToken(token);

  if (!tokenData) {
    return HtmlService.createHtmlOutput(getConsentErrorPageHtml())
      .setTitle('エラー - 同意書確認')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 申込者情報を取得
  const rowIndex = findRowByApplicationId(tokenData.applicationId);
  if (!rowIndex) {
    return HtmlService.createHtmlOutput(getConsentErrorPageHtml())
      .setTitle('エラー - 同意書確認')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const data = getRowData(rowIndex);

  // 既に同意済みの場合
  if (data.ndaStatus === '済') {
    return HtmlService.createHtmlOutput(getConsentAlreadyAgreedPageHtml(data))
      .setTitle('同意済 - 関西学院大学 中小企業経営相談研究会')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const html = getConsentPageHtml(data, token);
  return HtmlService.createHtmlOutput(html)
    .setTitle('相談同意書のご確認 - 関西学院大学 中小企業経営相談研究会')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 同意書同意処理
 * フォームからPOSTされた同意データを処理
 * @param {Object} e - POSTリクエストパラメータ
 * @returns {Object} 処理結果
 */
function processNdaConsent(e) {
  try {
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

    const token = params.token;
    const signature = params.signature;
    const agreed = params.agreed;

    console.log('相談者同意処理開始: token=' + (token ? token.substring(0, 8) + '...' : 'null'));

    if (!token || !signature || agreed !== 'true') {
      console.log('相談者同意: 必須項目不足 token=' + !!token + ' signature=' + !!signature + ' agreed=' + agreed);
      return { success: false, message: '必須項目が入力されていません' };
    }

    // トークン検証
    const tokenData = validateNdaToken(token);
    if (!tokenData) {
      console.log('相談者同意: トークン無効');
      return { success: false, message: '無効なトークンです' };
    }

    console.log('相談者同意: 申込ID=' + tokenData.applicationId);

    // スプレッドシートの行を特定
    const rowIndex = findRowByApplicationId(tokenData.applicationId);
    if (!rowIndex) {
      console.log('相談者同意: 申込データ未発見 ID=' + tokenData.applicationId);
      return { success: false, message: '申込データが見つかりません' };
    }

    // 同意ステータスを更新（U列/V列 + O列を「同意済」に）
    updateNdaStatus(rowIndex, signature);
    console.log('相談者同意: ステータス更新完了 行=' + rowIndex);

    // トークンを無効化
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('nda_token_' + token);

    // データ取得
    var data = getRowData(rowIndex);
    const isOnline = data.method && (data.method.indexOf('オンライン') >= 0 || data.method.indexOf('Zoom') >= 0 || data.method.indexOf('zoom') >= 0);

    console.log('相談者同意: method=' + data.method + ' isOnline=' + isOnline + ' confirmedDate=' + data.confirmedDate);

    if (isOnline && data.confirmedDate) {
      // ===== Zoom相談：自動確定フロー =====
      const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
        .getSheetByName(CONFIG.SHEET_NAME);

      // ステータスを「確定」に変更
      sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(STATUS.CONFIRMED);
      console.log('相談者同意[Zoom]: ステータス→確定');

      // Zoomミーティングを自動作成してR列に保存
      if (!data.zoomUrl) {
        try {
          createAndSaveZoomMeeting(data, rowIndex);
          console.log('相談者同意[Zoom]: Zoomミーティング作成完了');
        } catch (zoomErr) {
          console.error('相談者同意[Zoom]: Zoom作成失敗:', zoomErr);
        }
      }

      // Zoom作成後にデータ再取得（zoomUrl反映のため）
      data = getRowData(rowIndex);

      // 日程設定シートの予約状況を「予約済み」に更新
      var parsed = parseConfirmedDateTime(data.confirmedDate);
      if (parsed && parsed.date) {
        markAsBooked(parsed.date, parsed.time);
      }

      // 相談者に同意完了＋確定メール送信
      try {
        sendConsentAndConfirmedEmail(data);
        console.log('相談者同意[Zoom]: 確定メール送信完了 → ' + data.email);
      } catch (mailErr) {
        console.error('相談者同意[Zoom]: 確定メール送信失敗:', mailErr);
        // フォールバック: 管理者に通知
        notifyConsentEmailFailure_(data, mailErr);
      }

      // 管理者に通知（Zoom自動確定）
      notifyConsentAgreedAutoConfirmed(data, signature);

      // 担当者への確定通知（リーダー情報付き）
      var staffTarget = data.leader || data.staff;
      if (staffTarget) {
        var emailResult = buildStaffNotificationEmail_(data);
        sendStaffNotifications(
          staffTarget,
          '✅ Zoom予約自動確定\n\n申込ID: ' + data.id + '\nお名前: ' + data.name + '様\n貴社名: ' + data.company + '\n日時: ' + data.confirmedDate,
          emailResult.subject,
          emailResult.body
        );
      }

      console.log('相談者同意[Zoom]: 自動確定完了 ' + data.email);
      return { success: true, message: '同意が完了し、予約が確定しました' };

    } else {
      // ===== 対面相談：会場確保待ちフロー =====
      console.log('相談者同意[対面]: 会場確保待ちフロー開始');

      // 管理者・担当者・参加メンバーに通知（会場確保依頼）
      notifyConsentAgreed(data, signature);

      // 相談者に同意完了確認メール送信（会場確保待ち案内）
      try {
        sendConsentConfirmationToApplicant(data);
        console.log('相談者同意[対面]: 確認メール送信完了 → ' + data.email);
      } catch (mailErr) {
        console.error('相談者同意[対面]: 確認メール送信失敗:', mailErr);
        notifyConsentEmailFailure_(data, mailErr);
      }

      return { success: true, message: '同意が完了しました' };
    }

  } catch (error) {
    console.error('相談者同意処理エラー:', error);
    // エラー時も管理者にフォールバック通知
    try {
      CONFIG.ADMIN_EMAILS.forEach(function(email) {
        GmailApp.sendEmail(email, '【エラー】相談者同意処理失敗', '相談者同意処理でエラーが発生しました。\n\n' + error.toString() + '\n\nスタックトレース:\n' + error.stack, {
          name: CONFIG.SENDER_NAME
        });
      });
    } catch (notifyErr) {
      console.error('エラー通知も失敗:', notifyErr);
    }
    return { success: false, message: 'エラーが発生しました: ' + error.toString() };
  }
}

/**
 * 相談者へのメール送信失敗時の管理者フォールバック通知
 */
function notifyConsentEmailFailure_(data, error) {
  try {
    CONFIG.ADMIN_EMAILS.forEach(function(email) {
      GmailApp.sendEmail(email,
        '【要対応】相談者同意メール送信失敗 - ' + (data.id || '不明'),
        '相談者への同意完了メールの送信に失敗しました。手動で連絡してください。\n\n' +
        '申込ID：' + data.id + '\nお名前：' + data.name + '\nメール：' + data.email + '\n\nエラー：' + error.toString(),
        { name: CONFIG.SENDER_NAME }
      );
    });
  } catch (e) {
    console.error('フォールバック通知も失敗:', e);
  }
}

/**
 * 同意書同意完了時の管理者通知
 * @param {Object} data - 申込データ
 * @param {string} signature - 電子署名（氏名）
 */
/**
 * 対面相談：同意完了 → 管理者・担当者・参加メンバーに会場確保依頼通知
 */
function notifyConsentAgreed(data, signature) {
  var nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

  var subject = '【同意完了・会場確保依頼】' + data.name + '様 - ' + data.id;
  var body = '相談同意書への同意が完了しました（対面相談）。\n' +
    '会場の予約をお願いいたします。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 申込情報\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '申込ID：' + data.id + '\n' +
    'お名前：' + data.name + '\n' +
    '貴社名：' + data.company + '\n' +
    '業種：' + (data.industry || '未入力') + '\n' +
    '相談方法：' + data.method + '\n' +
    'テーマ：' + (data.theme || '未入力') + '\n' +
    '希望日時：' + data.date1 + (data.date2 ? '\n第二希望：' + data.date2 : '') + '\n' +
    (data.companyUrl ? '企業URL：' + data.companyUrl + '\n' : '') +
    '電子署名：' + signature + '\n' +
    '同意日時：' + nowStr + '\n' +
    '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 対応事項：会場の予約\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '以下の手順で会場を予約し、予約を確定してください。\n\n' +
    '手順：\n' +
    '1. スタッフポータルにログイン\n' +
    '   ' + CONFIG.PORTAL.SITE_URL + '\n' +
    '2.「案件」ページを開く\n' +
    '3. 該当案件の「会場未設定」をタップ\n' +
    '4. 会場を選択して「確定」をタップ\n\n' +
    '確定すると相談者に予約確定メールが自動送信されます。\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

  var lineMessage = '📋 同意完了（対面・会場確保依頼）\n\n' +
    '申込ID: ' + data.id + '\n' +
    'お名前: ' + data.name + '様\n' +
    '貴社名: ' + data.company + '\n' +
    '方法: ' + data.method + '\n' +
    'テーマ: ' + (data.theme || '-') + '\n' +
    '希望日時: ' + data.date1 + '\n' +
    (data.companyUrl ? '企業URL: ' + data.companyUrl + '\n' : '') +
    '同意日時: ' + nowStr + '\n\n' +
    '→ 会場の予約をお願いします。\n' +
    '→ ポータル「案件」から会場を選択して確定してください。\n' +
    '→ ' + CONFIG.PORTAL.SITE_URL;

  // 管理者に通知
  CONFIG.ADMIN_EMAILS.forEach(function(email) {
    GmailApp.sendEmail(email, subject, body, { name: CONFIG.SENDER_NAME });
  });
  var sentEmails = {};
  CONFIG.ADMIN_EMAILS.forEach(function(email) { sentEmails[email] = true; });

  // 担当者（P列）に通知
  if (data.staff) {
    sendStaffNotifications(data.staff, lineMessage, subject, body);
    var staffNames = data.staff.split(',').map(function(n) { return n.trim(); }).filter(function(n) { return n; });
    staffNames.forEach(function(name) {
      var m = getMemberByName(name);
      if (m && m.email) sentEmails[m.email] = true;
    });
  }

  // 日程設定シートの参加メンバー全員にも通知（重複除外）
  var memberList = getScheduleMembersForDate_(data.date1 || data.confirmedDate);
  if (memberList && memberList.length > 0) {
    memberList.forEach(function(cm) {
      var m = getMemberByName(cm.name);
      if (m && m.email && !sentEmails[m.email]) {
        GmailApp.sendEmail(m.email, subject, body, { name: CONFIG.SENDER_NAME });
        if (m.lineId) sendLineMessage(m.lineId, lineMessage);
        sentEmails[m.email] = true;
        console.log('同意完了通知（参加メンバー）: ' + cm.name + ' (' + m.email + ')');
      }
    });
  }

  console.log('同意完了・会場確保依頼通知: 送信先 ' + Object.keys(sentEmails).length + '名');
}

/**
 * Zoom相談：同意完了 → 自動確定通知
 */
function notifyConsentAgreedAutoConfirmed(data, signature) {
  const subject = `【自動確定・Zoom】${data.name}様 - ${data.id}`;
  const body = `Zoom相談のため、相談者同意書の同意完了時に自動確定しました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 確定情報
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
お名前：${data.name}
貴社名：${data.company}
日時：${data.confirmedDate}
相談方法：${data.method}
担当：${data.staff || '（未割当）'}
${data.zoomUrl ? 'Zoom URL：' + data.zoomUrl + '（自動作成済み）' : '※Zoom URLの自動作成に失敗しました。手動で設定してください。'}
電子署名：${signature}
同意日時：${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;

  CONFIG.ADMIN_EMAILS.forEach(email => {
    GmailApp.sendEmail(email, subject, body, { name: CONFIG.SENDER_NAME });
  });

  const lineMessage = `✅ Zoom自動確定

申込ID: ${data.id}
お名前: ${data.name}様
貴社名: ${data.company}
日時: ${data.confirmedDate}
同意日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}

※Zoom URLが未設定の場合は前日までに設定してください。`;

  sendLineMessage(CONFIG.LINE.GROUP_ID, lineMessage);
}

/**
 * 申込IDから行番号を取得
 * @param {string} applicationId - 申込ID
 * @returns {number|null} 行番号（1-based）
 */
function findRowByApplicationId(applicationId) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.ID] === applicationId) {
      return i + 1; // 1-based row number
    }
  }

  return null;
}

/**
 * 同意書同意ステータスを更新
 * @param {number} rowIndex - 行番号（1-based）
 * @param {string} signature - 電子署名
 */
function updateNdaStatus(rowIndex, signature) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);

  // U列: 同意書同意 = 「済」
  sheet.getRange(rowIndex, COLUMNS.NDA_STATUS + 1).setValue('済');

  // V列: 同意日時
  sheet.getRange(rowIndex, COLUMNS.NDA_DATE + 1).setValue(new Date());

  // O列: ステータスを「同意済」に更新（相談者同意書）
  const currentStatus = sheet.getRange(rowIndex, COLUMNS.STATUS + 1).getValue();
  if (currentStatus === STATUS.PENDING || currentStatus === '' || !currentStatus ||
      currentStatus === STATUS.NDA_AGREED) {
    sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(STATUS.CONSENT_AGREED);
  }

  // Q列: 確定日時をK列（希望日時1）+ 日程設定シートの時間帯から自動設定
  const confirmedDate = sheet.getRange(rowIndex, COLUMNS.CONFIRMED_DATE + 1).getValue();
  if (!confirmedDate || confirmedDate === '') {
    const fullDateTime = resolveConfirmedDateTime(rowIndex, sheet);
    if (fullDateTime) {
      sheet.getRange(rowIndex, COLUMNS.CONFIRMED_DATE + 1).setValue(fullDateTime);
    }
  }

  console.log('相談者同意ステータス更新: 行=' + rowIndex + ' 旧ステータス=' + currentStatus + ' → 同意済');
}

/**
 * 同意書ページで同意送信時にGAS側で呼ばれる関数
 * （HtmlServiceのgoogle.script.runから呼び出し用）
 * @param {Object} formData - フォームデータ
 * @returns {Object} 処理結果
 */
function submitNdaConsent(formData) {
  return processNdaConsent({
    parameter: formData,
    postData: null
  });
}

/**
 * 同意完了後に相談者へ確認メールを送信
 * @param {Object} data - 申込データ
 */
/**
 * 同意書PDFをGoogle Driveで更新
 * @param {string} base64Data - base64エンコードされたPDFデータ
 * @returns {Object} 処理結果
 */
function updateConsentPdf(base64Data) {
  try {
    const fileId = CONFIG.CONSENT.PDF_FILE_ID;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', '経営相談同意書.pdf');
    const file = DriveApp.getFileById(fileId);

    // 既存ファイルの内容を更新（Drive API v2を使用）
    const url = 'https://www.googleapis.com/upload/drive/v2/files/' + fileId + '?uploadType=media';
    const options = {
      method: 'put',
      contentType: 'application/pdf',
      payload: blob.getBytes(),
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (response.getResponseCode() === 200) {
      return { success: true, message: 'PDFを更新しました', fileId: fileId };
    } else {
      return { success: false, message: 'Drive API エラー: ' + response.getContentText() };
    }
  } catch (error) {
    console.error('PDF更新エラー:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * URLからPDFを取得してGoogle Driveに新しいファイルとして保存
 * 旧ファイルは削除し、新ファイルのIDを返す
 * @param {string} pdfUrl - PDFファイルのURL
 * @returns {Object} 処理結果
 */
function updateConsentPdfFromUrl(pdfUrl) {
  try {
    var oldFileId = CONFIG.CONSENT.PDF_FILE_ID;

    // URLからPDFをダウンロード
    var dlResponse = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true });
    if (dlResponse.getResponseCode() !== 200) {
      return { success: false, message: 'PDFダウンロード失敗: HTTP ' + dlResponse.getResponseCode() };
    }

    var pdfBlob = dlResponse.getBlob().setName('経営相談同意書.pdf');

    // 旧ファイルの親フォルダを取得
    var oldFile = DriveApp.getFileById(oldFileId);
    var folders = oldFile.getParents();
    var parentFolder = folders.hasNext() ? folders.next() : DriveApp.getRootFolder();

    // 旧ファイルの共有設定を取得
    var oldAccess = oldFile.getSharingAccess();
    var oldPermission = oldFile.getSharingPermission();

    // 新しいファイルを同じフォルダに作成
    var newFile = parentFolder.createFile(pdfBlob);

    // 共有設定を引き継ぐ
    newFile.setSharing(oldAccess, oldPermission);

    // 旧ファイルをゴミ箱へ
    oldFile.setTrashed(true);

    var newFileId = newFile.getId();
    return {
      success: true,
      message: 'PDFを更新しました。新しいファイルID: ' + newFileId,
      newFileId: newFileId,
      oldFileId: oldFileId
    };
  } catch (error) {
    console.error('PDF更新エラー:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * 同意完了確認メール送信（対面用：会場確保待ち案内）
 */
function sendConsentConfirmationToApplicant(data) {
  try {
    const subject = `【同意完了】相談同意書への同意を受領しました - ${data.id}`;
    const body = `${data.name} 様

相談同意書への同意を受領いたしました。ありがとうございます。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 同意内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
お名前：${data.name}
貴社名：${data.company || '（個人）'}
同意日時：${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

現在、会場の確保を行っております。
会場が確定次第、3日以内に予約確定メールをお送りいたします。
しばらくお待ちください。

※会場の都合により、ご希望の日程での開催が難しい場合は
  別日程へのご変更をお願いする場合がございます。
  あらかじめご了承ください。

ご不明な点がございましたら、お気軽にお問い合わせください。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
URL: ${CONFIG.ORG.URL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;

    GmailApp.sendEmail(data.email, subject, body, {
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.REPLY_TO
    });
    console.log(`同意完了確認メール（対面・会場確保待ち）を ${data.email} に送信しました`);
  } catch (e) {
    console.error('同意完了確認メール送信エラー:', e);
  }
}

/**
 * Zoom相談：同意完了＋自動確定メール送信
 * NDA同意完了時に自動で予約確定とし、確定メールを送信する
 */
function sendConsentAndConfirmedEmail(data) {
  try {
    const zoomInfo = data.zoomUrl
      ? `Zoom URL：${data.zoomUrl}\n\n※開始時刻の5分前を目安にご参加ください`
      : `Zoom URLについては、前日までに改めてメールでお送りいたします。\n\n※開始時刻の5分前を目安にご参加ください`;

    const subject = `【予約確定】無料経営相談のご予約が確定しました - ${data.id}`;
    const body = `${data.name} 様

相談同意書への同意を受領いたしました。ありがとうございます。
オンライン（Zoom）相談のため、ご予約が確定しましたのでお知らせいたします。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご予約内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
日時：${data.confirmedDate}
相談方法：${data.method}
${data.staff ? '担当：' + data.staff : ''}

【オンライン相談】
${zoomInfo}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【当日の流れ】
1. 現状のヒアリング（15分程度）
   - お話を伺います

2. 課題の整理・ディスカッション（30〜45分）
   - 課題を整理し、解決の方向性を一緒に考えます

3. 今後のアクション整理（15分程度）
   - 次のステップを明確にします

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご準備いただくもの
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
・関連資料（決算書、事業計画書等）があればご準備ください
  ※必須ではありません

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ キャンセル・変更について
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ご都合が悪くなった場合は、できるだけ早めにご連絡ください。
連絡先：${CONFIG.ORG.EMAIL}

ご不明な点がございましたら、お気軽にお問い合わせください。
当日お会いできることを楽しみにしております。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
URL: ${CONFIG.ORG.URL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;

    GmailApp.sendEmail(data.email, subject, body, {
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.REPLY_TO
    });
    console.log(`同意完了＋自動確定メール（Zoom）を ${data.email} に送信しました`);
  } catch (e) {
    console.error('同意完了＋自動確定メール送信エラー:', e);
  }
}
