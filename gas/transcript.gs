/**
 * 文字起こし・報告書自動作成管理（Phase 3）
 *
 * パイプライン:
 *   Zoom MP4 → [CF1: zoom_to_transcript] Speech-to-Text（話者分離）
 *            → [CF2: transcript_to_report] Claude API → Google Docs
 *            → [CF3: report_to_notion] Docs→PDF → Notion API（オプション）
 *
 * 事前設定:
 *   ScriptProperties に以下を設定:
 *   - TRANSCRIPT_CF_URL: zoom_to_transcript Cloud Function URL
 *   - TRANSCRIPT_CF_SECRET: 共有シークレット
 *   - REPORT_CF_URL: transcript_to_report Cloud Function URL
 *   - REPORT_CF_SECRET: 共有シークレット
 *   - NOTION_CF_URL: report_to_notion Cloud Function URL（オプション）
 *   - NOTION_CF_SECRET: 共有シークレット（オプション）
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 文字起こし処理
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * Zoom録画の文字起こしを開始
 * processZoomRecordings() から呼ばれる
 * @param {Object} mp4File - Zoom APIのrecording_fileオブジェクト（MP4）
 * @param {Object} rowData - getRowData形式の予約データ
 * @param {number} rowIndex - スプレッドシート行番号（1-based）
 * @param {string} zoomToken - Zoomアクセストークン
 */
function startTranscription(mp4File, rowData, rowIndex, zoomToken) {
  var props = PropertiesService.getScriptProperties();
  var cfUrl = props.getProperty('TRANSCRIPT_CF_URL') || (CONFIG.TRANSCRIPT && CONFIG.TRANSCRIPT.CLOUD_FUNCTION_URL) || '';
  var cfSecret = props.getProperty('TRANSCRIPT_CF_SECRET') || (CONFIG.TRANSCRIPT && CONFIG.TRANSCRIPT.CLOUD_FUNCTION_SECRET) || '';

  if (!cfUrl) {
    console.log('文字起こしCloud Function URLが未設定のためスキップ');
    return;
  }

  // AD列: 処理中に更新
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
  sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_STATUS + 1).setValue(TRANSCRIPT_STATUS.PROCESSING);

  var payload = {
    secret: cfSecret,
    zoom_token: zoomToken,
    download_url: mp4File.download_url,
    application_id: rowData.id || '',
    meeting_topic: (rowData.company || rowData.name || '') + '様 経営相談'
  };

  console.log('文字起こしリクエスト送信: ' + rowData.id + ' -> ' + cfUrl);

  try {
    var response = UrlFetchApp.fetch(cfUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 600  // 10分（長時間音声対応）
    });

    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      console.error('文字起こしエラー (' + code + '): ' + body);
      sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_STATUS + 1).setValue(TRANSCRIPT_STATUS.ERROR);
      return;
    }

    var result = JSON.parse(body);
    if (!result.success || !result.transcript) {
      console.error('文字起こし結果エラー: ' + body);
      sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_STATUS + 1).setValue(TRANSCRIPT_STATUS.ERROR);
      return;
    }

    // 文字起こし結果をDriveに保存
    var transcriptFileId = saveTranscriptToDrive(rowData, result.transcript, result.duration_sec, result.speaker_count);

    // スプレッドシート更新
    sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_STATUS + 1).setValue(TRANSCRIPT_STATUS.COMPLETED);
    sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_FILE_ID + 1).setValue(transcriptFileId);

    console.log('文字起こし完了: ' + rowData.id + ', ' + result.duration_sec + '秒, ' + result.speaker_count + '話者');

    // リーダーに文字起こし完了通知
    notifyTranscriptComplete(rowData, transcriptFileId, result.duration_sec, result.speaker_count);

    // 報告書自動生成が有効な場合、続けてレポート生成
    if (CONFIG.AUTO_REPORT && CONFIG.AUTO_REPORT.ENABLED) {
      startAutoReportGeneration(result.transcript, rowData, rowIndex);
    }

  } catch (e) {
    console.error('文字起こし処理エラー:', e);
    sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_STATUS + 1).setValue(TRANSCRIPT_STATUS.ERROR);
  }
}

/**
 * 文字起こし結果をDriveに保存
 * @param {Object} rowData - 予約データ
 * @param {string} transcript - 文字起こしテキスト
 * @param {number} durationSec - 録画時間（秒）
 * @param {number} speakerCount - 話者数
 * @returns {string} DriveファイルID
 */
function saveTranscriptToDrive(rowData, transcript, durationSec, speakerCount) {
  var folder = getDriveFolder('DRIVE_FOLDER_TRANSCRIPT', '');

  var header = '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '経営相談 文字起こし\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
    '申込ID：' + (rowData.id || '') + '\n' +
    '相談日時：' + (rowData.confirmedDate || '') + '\n' +
    '相談者：' + (rowData.name || '') + '（' + (rowData.company || '') + '）\n' +
    '業種：' + (rowData.industry || '') + '\n' +
    'テーマ：' + (rowData.theme || '') + '\n' +
    'リーダー：' + (rowData.leader || '') + '\n' +
    '録画時間：' + Math.round(durationSec / 60) + '分\n' +
    '話者数：' + speakerCount + '名\n' +
    '文字起こし日時：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

  var content = header + transcript;
  var fileName = (rowData.id || 'unknown') + '_transcript.txt';
  var file = folder.createFile(fileName, content, MimeType.PLAIN_TEXT);

  console.log('文字起こし保存: ' + fileName + ' (ID: ' + file.getId() + ')');
  return file.getId();
}

/**
 * リーダーに文字起こし完了を通知
 * @param {Object} rowData - 予約データ
 * @param {string} transcriptFileId - 文字起こしファイルID
 * @param {number} durationSec - 録画時間（秒）
 * @param {number} speakerCount - 話者数
 */
function notifyTranscriptComplete(rowData, transcriptFileId, durationSec, speakerCount) {
  if (!rowData.leader) {
    console.log('リーダー未設定のため文字起こし通知スキップ: ' + rowData.id);
    return;
  }

  var leaderMember = getMemberByName(rowData.leader);
  if (!leaderMember || !leaderMember.email) {
    console.log('リーダーメールアドレス不明: ' + rowData.leader);
    return;
  }

  var transcriptUrl = 'https://drive.google.com/file/d/' + transcriptFileId;

  var subject = '【文字起こし完了】' + (rowData.company || rowData.name) + '様 - 経営相談';

  var body = rowData.leader + ' 様\n\n' +
    'お疲れ様です。\n' +
    '経営相談の録画文字起こしが完了しましたのでお知らせいたします。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 相談概要\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '申込ID　：' + (rowData.id || '') + '\n' +
    '相談日時：' + (rowData.confirmedDate || '') + '\n' +
    '相談者　：' + (rowData.name || '') + ' 様（' + (rowData.company || '') + '）\n' +
    'テーマ　：' + (rowData.theme || '') + '\n' +
    '録画時間：約' + Math.round(durationSec / 60) + '分\n' +
    '話者数　：' + speakerCount + '名（自動検出）\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 文字起こしファイル\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    transcriptUrl + '\n\n' +
    '※ Google Driveで閲覧・ダウンロードが可能です。\n' +
    '※ 話者分離が適用されています（【話者1】【話者2】等で区別）。\n';

  // 報告書自動生成が有効な場合の追加メッセージ
  if (CONFIG.AUTO_REPORT && CONFIG.AUTO_REPORT.ENABLED) {
    body += '\n報告書ドラフトの自動生成も開始されています。\n' +
      '完了次第、別途ご連絡いたします。\n';
  } else {
    body += '\n文字起こし内容を参考に、診断報告書の作成をお願いいたします。\n';
  }

  body += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    CONFIG.ORG.NAME + '\n' +
    'Email: ' + CONFIG.ORG.EMAIL + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

  GmailApp.sendEmail(leaderMember.email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });

  // 管理者にも通知
  CONFIG.ADMIN_EMAILS.forEach(function(adminEmail) {
    if (adminEmail !== leaderMember.email) {
      GmailApp.sendEmail(adminEmail, subject, '※管理者控え※\n\n' + body, {
        name: CONFIG.SENDER_NAME
      });
    }
  });

  console.log('文字起こし通知送信完了: ' + rowData.leader + ', ' + rowData.id);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 報告書自動生成
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 文字起こしから報告書ドラフトを自動生成
 * @param {string} transcript - 文字起こしテキスト
 * @param {Object} rowData - 予約データ
 * @param {number} rowIndex - 行番号
 */
function startAutoReportGeneration(transcript, rowData, rowIndex) {
  var props = PropertiesService.getScriptProperties();
  var cfUrl = props.getProperty('REPORT_CF_URL') || (CONFIG.AUTO_REPORT && CONFIG.AUTO_REPORT.CLOUD_FUNCTION_URL) || '';
  var cfSecret = props.getProperty('REPORT_CF_SECRET') || (CONFIG.AUTO_REPORT && CONFIG.AUTO_REPORT.CLOUD_FUNCTION_SECRET) || '';

  if (!cfUrl) {
    console.log('報告書生成Cloud Function URLが未設定のためスキップ');
    return;
  }

  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);

  var payload = {
    secret: cfSecret,
    transcript: transcript,
    application_id: rowData.id || '',
    company: rowData.company || '',
    industry: rowData.industry || '',
    theme: rowData.theme || '',
    name: rowData.name || '',
    leader: rowData.leader || '',
    confirmed_date: rowData.confirmedDate ? rowData.confirmedDate.toString() : ''
  };

  console.log('報告書生成リクエスト送信: ' + rowData.id + ' -> ' + cfUrl);

  try {
    var response = UrlFetchApp.fetch(cfUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 300  // 5分
    });

    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      console.error('報告書生成エラー (' + code + '): ' + body);
      return;
    }

    var result = JSON.parse(body);
    if (!result.success || !result.doc_id) {
      console.error('報告書生成結果エラー: ' + body);
      return;
    }

    // AF列: ドラフトGoogle Docs IDを保存
    sheet.getRange(rowIndex, COLUMNS.REPORT_DRAFT_ID + 1).setValue(result.doc_id);

    console.log('報告書ドラフト生成完了: ' + rowData.id + ' -> DocID=' + result.doc_id);

    // リーダーにドラフト完了通知
    notifyReportDraftComplete(rowData, result.doc_id, result.doc_url);

    // Notion連携（オプション）
    if (CONFIG.NOTION_CF && CONFIG.NOTION_CF.ENABLED) {
      startNotionIntegration(result, rowData, rowIndex);
    }

  } catch (e) {
    console.error('報告書自動生成エラー:', e);
  }
}

/**
 * リーダーに報告書ドラフト完了を通知
 * @param {Object} rowData - 予約データ
 * @param {string} docId - Google Docs ID
 * @param {string} docUrl - Google Docs URL
 */
function notifyReportDraftComplete(rowData, docId, docUrl) {
  if (!rowData.leader) return;

  var leaderMember = getMemberByName(rowData.leader);
  if (!leaderMember || !leaderMember.email) return;

  var subject = '【報告書ドラフト完了】' + (rowData.company || rowData.name) + '様 - レビューをお願いします';

  var body = rowData.leader + ' 様\n\n' +
    'お疲れ様です。\n' +
    '経営相談の診断報告書ドラフトが自動生成されました。\n' +
    'レビュー・修正をお願いいたします。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 相談概要\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '申込ID　：' + (rowData.id || '') + '\n' +
    '相談日時：' + (rowData.confirmedDate || '') + '\n' +
    '相談者　：' + (rowData.name || '') + ' 様（' + (rowData.company || '') + '）\n' +
    'テーマ　：' + (rowData.theme || '') + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 報告書ドラフト（Google Docs）\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    (docUrl || 'https://docs.google.com/document/d/' + docId + '/edit') + '\n\n' +
    '【お願い事項】\n' +
    '1. ドラフト内容を確認し、必要に応じて修正してください\n' +
    '2. AI生成のため、事実誤認・誤字脱字がある場合があります\n' +
    '3. 修正完了後、レポートアップロード機能で最終版をアップロードしてください\n' +
    '   （別途、アップロード用リンクをお送りします）\n\n' +
    '※ このドラフトはAI（Claude）により自動生成されたものです。\n' +
    '※ 必ずリーダーの確認・修正を経てから相談者への配信をお願いします。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    CONFIG.ORG.NAME + '\n' +
    'Email: ' + CONFIG.ORG.EMAIL + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

  GmailApp.sendEmail(leaderMember.email, subject, body, {
    name: CONFIG.SENDER_NAME,
    replyTo: CONFIG.REPLY_TO
  });

  // 管理者にも通知
  CONFIG.ADMIN_EMAILS.forEach(function(adminEmail) {
    if (adminEmail !== leaderMember.email) {
      GmailApp.sendEmail(adminEmail, subject, '※管理者控え※\n\n' + body, {
        name: CONFIG.SENDER_NAME
      });
    }
  });

  console.log('報告書ドラフト通知送信: ' + rowData.leader + ', DocID=' + docId);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// Notion連携（オプション）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * Notion連携を開始（報告書ドラフト完了後に呼ばれる）
 * @param {Object} reportResult - transcript_to_reportの結果
 * @param {Object} rowData - 予約データ
 * @param {number} rowIndex - 行番号
 */
function startNotionIntegration(reportResult, rowData, rowIndex) {
  var props = PropertiesService.getScriptProperties();
  var cfUrl = props.getProperty('NOTION_CF_URL') || (CONFIG.NOTION_CF && CONFIG.NOTION_CF.CLOUD_FUNCTION_URL) || '';
  var cfSecret = props.getProperty('NOTION_CF_SECRET') || (CONFIG.NOTION_CF && CONFIG.NOTION_CF.CLOUD_FUNCTION_SECRET) || '';

  if (!cfUrl) {
    console.log('Notion Cloud Function URLが未設定のためスキップ');
    return;
  }

  var payload = {
    secret: cfSecret,
    doc_id: reportResult.doc_id,
    doc_url: reportResult.doc_url,
    application_id: rowData.id || '',
    company: rowData.company || '',
    industry: rowData.industry || '',
    theme: rowData.theme || '',
    leader: rowData.leader || '',
    confirmed_date: rowData.confirmedDate ? rowData.confirmedDate.toString() : '',
    report_summary: (reportResult.report_text || '').substring(0, 1000),
    drive_folder_id: CONFIG.AUTO_REPORT.DRIVE_FOLDER_ID || ''
  };

  console.log('Notion連携リクエスト送信: ' + rowData.id);

  try {
    var response = UrlFetchApp.fetch(cfUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 120
    });

    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      console.error('Notion連携エラー (' + code + '): ' + body);
      return;
    }

    var result = JSON.parse(body);
    if (!result.success) {
      console.error('Notion連携結果エラー: ' + body);
      return;
    }

    // AG列: Notion Page IDを保存
    if (result.notion_page_id) {
      var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
      sheet.getRange(rowIndex, COLUMNS.NOTION_PAGE_ID + 1).setValue(result.notion_page_id);
    }

    console.log('Notion連携完了: ' + rowData.id + ', pageId=' + (result.notion_page_id || 'N/A'));

  } catch (e) {
    console.error('Notion連携エラー:', e);
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 手動操作 & 管理用関数
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 指定行の予約に対して手動で文字起こしを開始
 * @param {number} rowIndex - 行番号（1-based）
 * @returns {Object} 結果
 */
function manualStartTranscription(rowIndex) {
  try {
    var rowData = getRowData(rowIndex);
    if (!rowData.id) {
      return { success: false, message: '行 ' + rowIndex + ' にデータがありません' };
    }

    // Zoom録画URLが必要
    if (!rowData.recordingUrl && !rowData.zoomUrl) {
      return { success: false, message: '録画URLが未設定です（AB列 or R列）' };
    }

    // Zoomトークン取得
    var token = getZoomAccessToken();

    // Zoom URLからミーティングIDを取得
    var zoomUrl = rowData.zoomUrl || '';
    var match = zoomUrl.match(/\/j\/(\d+)/);
    if (!match) {
      return { success: false, message: 'Zoom URLからミーティングIDを抽出できません' };
    }
    var meetingId = match[1];

    // 録画一覧を取得
    var recResponse = UrlFetchApp.fetch(
      'https://api.zoom.us/v2/meetings/' + meetingId + '/recordings',
      {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    if (recResponse.getResponseCode() !== 200) {
      return { success: false, message: '録画情報の取得に失敗: ' + recResponse.getContentText() };
    }

    var recData = JSON.parse(recResponse.getContentText());
    var mp4File = findMp4Recording(recData.recording_files || []);

    if (!mp4File) {
      return { success: false, message: 'MP4録画ファイルが見つかりません' };
    }

    // 文字起こし開始
    startTranscription(mp4File, rowData, rowIndex, token);
    return { success: true, message: '文字起こしを開始しました', applicationId: rowData.id };

  } catch (e) {
    console.error('手動文字起こし開始エラー:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * 文字起こしテキストを取得（DriveファイルIDから読み込み）
 * @param {string} fileId - DriveファイルID
 * @returns {string} 文字起こしテキスト
 */
function getTranscriptText(fileId) {
  if (!fileId) return '';
  try {
    var file = DriveApp.getFileById(fileId);
    return file.getBlob().getDataAsString('UTF-8');
  } catch (e) {
    console.error('文字起こし読み込みエラー:', e);
    return '';
  }
}

/**
 * 指定行に対して手動で報告書自動生成を開始
 * @param {number} rowIndex - 行番号
 * @returns {Object} 結果
 */
function manualStartReportGeneration(rowIndex) {
  try {
    var rowData = getRowData(rowIndex);
    if (!rowData.id) {
      return { success: false, message: '行 ' + rowIndex + ' にデータがありません' };
    }

    var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    var transcriptFileId = sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_FILE_ID + 1).getValue();

    if (!transcriptFileId) {
      return { success: false, message: '文字起こしファイルが未設定です（AE列）' };
    }

    var transcript = getTranscriptText(transcriptFileId);
    if (!transcript) {
      return { success: false, message: '文字起こしテキストの読み込みに失敗しました' };
    }

    startAutoReportGeneration(transcript, rowData, rowIndex);
    return { success: true, message: '報告書自動生成を開始しました', applicationId: rowData.id };

  } catch (e) {
    console.error('手動報告書生成エラー:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * 文字起こし・報告書パイプラインの状態一覧を取得
 * @returns {Object} パイプライン状態
 */
function getTranscriptPipelineStatus() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return { success: false, message: 'シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  var pipelines = [];

  for (var i = 1; i < data.length; i++) {
    var status = data[i][COLUMNS.STATUS];
    // 完了済み案件のみ
    if (status !== STATUS.COMPLETED) continue;

    var transcriptStatus = data[i][COLUMNS.TRANSCRIPT_STATUS] || '';
    var transcriptFileId = data[i][COLUMNS.TRANSCRIPT_FILE_ID] || '';
    var reportDraftId = data[i][COLUMNS.REPORT_DRAFT_ID] || '';
    var notionPageId = data[i][COLUMNS.NOTION_PAGE_ID] || '';

    pipelines.push({
      row: i + 1,
      id: data[i][COLUMNS.ID] || '',
      company: data[i][COLUMNS.COMPANY] || '',
      leader: data[i][COLUMNS.LEADER] || '',
      confirmedDate: data[i][COLUMNS.CONFIRMED_DATE] || '',
      recordingUrl: data[i][COLUMNS.RECORDING_URL] ? '有' : '無',
      youtubeUrl: data[i][COLUMNS.YOUTUBE_URL] ? '有' : '無',
      transcriptStatus: transcriptStatus,
      transcriptFileId: transcriptFileId ? '有' : '無',
      reportDraftId: reportDraftId ? '有' : '無',
      reportDraftUrl: reportDraftId ? 'https://docs.google.com/document/d/' + reportDraftId + '/edit' : '',
      notionPageId: notionPageId ? '有' : '無',
      reportStatus: data[i][COLUMNS.REPORT_STATUS] || ''
    });
  }

  return {
    success: true,
    count: pipelines.length,
    pipelines: pipelines,
    config: {
      transcriptEnabled: !!(CONFIG.TRANSCRIPT && CONFIG.TRANSCRIPT.ENABLED),
      autoReportEnabled: !!(CONFIG.AUTO_REPORT && CONFIG.AUTO_REPORT.ENABLED),
      notionEnabled: !!(CONFIG.NOTION_CF && CONFIG.NOTION_CF.ENABLED)
    }
  };
}

/**
 * 文字起こし・報告書設定のセットアップ確認
 * @returns {Object} 設定状態
 */
function checkTranscriptSetup() {
  var props = PropertiesService.getScriptProperties();
  return {
    success: true,
    transcript: {
      enabled: !!(CONFIG.TRANSCRIPT && CONFIG.TRANSCRIPT.ENABLED),
      cfUrl: !!(props.getProperty('TRANSCRIPT_CF_URL') || (CONFIG.TRANSCRIPT && CONFIG.TRANSCRIPT.CLOUD_FUNCTION_URL)),
      cfSecret: !!(props.getProperty('TRANSCRIPT_CF_SECRET') || (CONFIG.TRANSCRIPT && CONFIG.TRANSCRIPT.CLOUD_FUNCTION_SECRET))
    },
    autoReport: {
      enabled: !!(CONFIG.AUTO_REPORT && CONFIG.AUTO_REPORT.ENABLED),
      cfUrl: !!(props.getProperty('REPORT_CF_URL') || (CONFIG.AUTO_REPORT && CONFIG.AUTO_REPORT.CLOUD_FUNCTION_URL)),
      cfSecret: !!(props.getProperty('REPORT_CF_SECRET') || (CONFIG.AUTO_REPORT && CONFIG.AUTO_REPORT.CLOUD_FUNCTION_SECRET))
    },
    notion: {
      enabled: !!(CONFIG.NOTION_CF && CONFIG.NOTION_CF.ENABLED),
      cfUrl: !!(props.getProperty('NOTION_CF_URL') || (CONFIG.NOTION_CF && CONFIG.NOTION_CF.CLOUD_FUNCTION_URL)),
      cfSecret: !!(props.getProperty('NOTION_CF_SECRET') || (CONFIG.NOTION_CF && CONFIG.NOTION_CF.CLOUD_FUNCTION_SECRET))
    }
  };
}

/**
 * 文字起こし・報告書 Cloud Function URL/Secretの一括設定
 * GASエディタで実行、またはAPIから呼び出し
 */
function setupTranscriptCredentials(transcriptCfUrl, transcriptCfSecret, reportCfUrl, reportCfSecret) {
  var props = PropertiesService.getScriptProperties();
  var updates = {};

  if (transcriptCfUrl) updates['TRANSCRIPT_CF_URL'] = transcriptCfUrl;
  if (transcriptCfSecret) updates['TRANSCRIPT_CF_SECRET'] = transcriptCfSecret;
  if (reportCfUrl) updates['REPORT_CF_URL'] = reportCfUrl;
  if (reportCfSecret) updates['REPORT_CF_SECRET'] = reportCfSecret;

  if (Object.keys(updates).length > 0) {
    props.setProperties(updates);
  }

  console.log('文字起こし・報告書設定を保存しました');
  return checkTranscriptSetup();
}
