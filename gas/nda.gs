/**
 * 同意書確認・同意処理
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
      .setTitle('同意済 - 関西学院大学 中小企業経営診断研究会')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const html = getConsentPageHtml(data, token);
  return HtmlService.createHtmlOutput(html)
    .setTitle('相談同意書のご確認 - 関西学院大学 中小企業経営診断研究会')
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

    if (!token || !signature || agreed !== 'true') {
      return { success: false, message: '必須項目が入力されていません' };
    }

    // トークン検証
    const tokenData = validateNdaToken(token);
    if (!tokenData) {
      return { success: false, message: '無効なトークンです' };
    }

    // スプレッドシートの行を特定
    const rowIndex = findRowByApplicationId(tokenData.applicationId);
    if (!rowIndex) {
      return { success: false, message: '申込データが見つかりません' };
    }

    // 同意ステータスを更新（T列/U列に記録のみ）
    updateNdaStatus(rowIndex, signature);

    // トークンを無効化
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('nda_token_' + token);

    // データ取得
    const data = getRowData(rowIndex);
    const isOnline = data.method && (data.method.indexOf('オンライン') >= 0 || data.method.indexOf('Zoom') >= 0 || data.method.indexOf('zoom') >= 0);

    if (isOnline && data.confirmedDate) {
      // ===== Zoom相談：自動確定フロー =====
      const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
        .getSheetByName(CONFIG.SHEET_NAME);

      // ステータスを「確定」に変更
      sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(STATUS.CONFIRMED);

      // 日程設定シートの予約状況を「予約済み」に更新
      var parsed = parseConfirmedDateTime(data.confirmedDate);
      if (parsed && parsed.date) {
        markAsBooked(parsed.date, parsed.time);
      }

      // 相談者に同意完了＋確定メール送信
      sendConsentAndConfirmedEmail(data);

      // 管理者に通知（Zoom自動確定）
      notifyConsentAgreedAutoConfirmed(data, signature);

      // 担当者への通知
      if (data.staff) {
        const staffLineMsg = `✅ Zoom予約自動確定\n\n申込ID: ${data.id}\nお名前: ${data.name}様\n貴社名: ${data.company}\n日時: ${data.confirmedDate}\n方法: ${data.method}\nテーマ: ${data.theme}\n${data.companyUrl ? '企業URL: ' + data.companyUrl + '\n※事前リサーチをお願いします' : ''}`;
        const staffEmailSubject = `【予約確定・Zoom自動】${data.name}様 - ${data.confirmedDate}`;
        const staffEmailBody = `Zoom相談のため、NDA同意完了時に自動確定しました。\n\n申込ID：${data.id}\nお名前：${data.name}様\n貴社名：${data.company}\n日時：${data.confirmedDate}\n相談方法：${data.method}\nテーマ：${data.theme}\n${data.companyUrl ? '\n企業URL：' + data.companyUrl + '\n※事前リサーチにAIツールの活用を推奨します' : ''}\n\n事前準備をお願いいたします。`;
        sendStaffNotifications(data.staff, staffLineMsg, staffEmailSubject, staffEmailBody);
      }

      console.log(`Zoom自動確定完了: ${data.email}`);
      return { success: true, message: '同意が完了し、予約が確定しました' };

    } else {
      // ===== 対面相談：会場確保待ちフロー =====
      // 管理者に通知（会場確保依頼）
      notifyConsentAgreed(data, signature);

      // 相談者に同意完了確認メール送信（会場確保待ち案内）
      sendConsentConfirmationToApplicant(data);

      return { success: true, message: '同意が完了しました' };
    }

  } catch (error) {
    console.error('同意処理エラー:', error);
    return { success: false, message: 'エラーが発生しました: ' + error.toString() };
  }
}

/**
 * 同意書同意完了時の管理者通知
 * @param {Object} data - 申込データ
 * @param {string} signature - 電子署名（氏名）
 */
/**
 * 対面相談：同意完了 → 管理者に会場確保依頼通知
 */
function notifyConsentAgreed(data, signature) {
  const subject = `【同意完了・会場確保依頼】${data.name}様 - ${data.id}`;
  const body = `同意書への同意が完了しました（対面相談）。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 同意情報
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
お名前：${data.name}
貴社名：${data.company}
相談方法：${data.method}
希望日時：${data.date1}${data.date2 ? '\n第二希望：' + data.date2 : ''}
電子署名：${signature}
同意日時：${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}

【対応事項】
会場を確保し、スプレッドシートの場所（N列）を設定後、
ステータスを「確定」に変更してください。
確定メールが自動送信されます。
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;

  CONFIG.ADMIN_EMAILS.forEach(email => {
    GmailApp.sendEmail(email, subject, body, { name: CONFIG.SENDER_NAME });
  });

  const lineMessage = `📋 同意完了（対面・会場確保依頼）

申込ID: ${data.id}
お名前: ${data.name}様
貴社名: ${data.company}
方法: ${data.method}
希望日時: ${data.date1}
同意日時: ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}

→ 会場を確保してステータスを「確定」にしてください。`;

  sendLineMessage(CONFIG.LINE.GROUP_ID, lineMessage);
}

/**
 * Zoom相談：同意完了 → 自動確定通知
 */
function notifyConsentAgreedAutoConfirmed(data, signature) {
  const subject = `【自動確定・Zoom】${data.name}様 - ${data.id}`;
  const body = `Zoom相談のため、NDA同意完了時に自動確定しました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 確定情報
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
お名前：${data.name}
貴社名：${data.company}
日時：${data.confirmedDate}
相談方法：${data.method}
担当：${data.staff || '（未割当）'}
電子署名：${signature}
同意日時：${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}

※Zoom URLが未設定の場合は、前日までに設定してください。
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

  // T列: 同意書同意 = 「済」
  sheet.getRange(rowIndex, COLUMNS.NDA_STATUS + 1).setValue('済');

  // U列: 同意日時
  sheet.getRange(rowIndex, COLUMNS.NDA_DATE + 1).setValue(new Date());

  // N列: ステータスを「NDA同意済」に更新
  const currentStatus = sheet.getRange(rowIndex, COLUMNS.STATUS + 1).getValue();
  if (currentStatus === STATUS.PENDING || currentStatus === '' || !currentStatus) {
    sheet.getRange(rowIndex, COLUMNS.STATUS + 1).setValue(STATUS.NDA_AGREED);
  }

  // P列: 確定日時をK列（希望日時1）+ 日程設定シートの時間帯から自動設定
  const confirmedDate = sheet.getRange(rowIndex, COLUMNS.CONFIRMED_DATE + 1).getValue();
  if (!confirmedDate || confirmedDate === '') {
    const fullDateTime = resolveConfirmedDateTime(rowIndex, sheet);
    if (fullDateTime) {
      sheet.getRange(rowIndex, COLUMNS.CONFIRMED_DATE + 1).setValue(fullDateTime);
    }
  }
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
