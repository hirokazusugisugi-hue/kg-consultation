/**
 * Zoom API連携
 * Server-to-Server OAuth でミーティングを自動作成・削除
 *
 * 事前設定:
 *   GASエディタ「プロジェクトの設定」→「スクリプト プロパティ」に以下を追加
 *   - ZOOM_ACCOUNT_ID
 *   - ZOOM_CLIENT_ID
 *   - ZOOM_CLIENT_SECRET
 */

/**
 * Zoom API用アクセストークンを取得（有効期限1時間、都度取得方式）
 * @returns {string} アクセストークン
 */
function getZoomAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const accountId = props.getProperty('ZOOM_ACCOUNT_ID') || (CONFIG.ZOOM && CONFIG.ZOOM.ACCOUNT_ID) || '';
  const clientId = props.getProperty('ZOOM_CLIENT_ID') || (CONFIG.ZOOM && CONFIG.ZOOM.CLIENT_ID) || '';
  const clientSecret = props.getProperty('ZOOM_CLIENT_SECRET') || (CONFIG.ZOOM && CONFIG.ZOOM.CLIENT_SECRET) || '';

  if (!accountId || !clientId || !clientSecret) {
    throw new Error('Zoom API認証情報が未設定です。CONFIG.ZOOMまたはスクリプトプロパティにZOOM_ACCOUNT_ID, ZOOM_CLIENT_ID, ZOOM_CLIENT_SECRETを設定してください。');
  }

  const credentials = Utilities.base64Encode(clientId + ':' + clientSecret);

  const response = UrlFetchApp.fetch('https://zoom.us/oauth/token', {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + credentials,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: 'grant_type=account_credentials&account_id=' + encodeURIComponent(accountId),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error('Zoomトークン取得エラー:', response.getContentText());
    throw new Error('Zoomアクセストークンの取得に失敗しました: ' + response.getContentText());
  }

  const result = JSON.parse(response.getContentText());
  return result.access_token;
}

/**
 * Zoomミーティングを作成し、Join URLを返す
 * @param {Object} data - 予約データ（getRowData()の戻り値）
 * @returns {string|null} Join URL（失敗時はnull）
 */
function createZoomMeeting(data) {
  try {
    const token = getZoomAccessToken();
    const startTime = convertToZoomDateTime(data.confirmedDate);

    if (!startTime) {
      console.error('Zoom作成エラー: 確定日時の変換に失敗しました - ' + data.confirmedDate);
      return null;
    }

    const payload = {
      topic: '【無料経営相談】' + (data.company || data.name) + ' 様',
      type: 2,  // スケジュール済みミーティング
      start_time: startTime,
      duration: 90,  // 1時間半
      timezone: 'Asia/Tokyo',
      settings: {
        waiting_room: false,
        join_before_host: true,
        mute_upon_entry: true,
        audio: 'voip',
        auto_recording: 'cloud',
        meeting_authentication: false
      }
    };

    const response = UrlFetchApp.fetch('https://api.zoom.us/v2/users/me/meetings', {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 201) {
      console.error('Zoomミーティング作成エラー:', response.getContentText());
      return null;
    }

    const meeting = JSON.parse(response.getContentText());
    console.log('Zoomミーティング作成成功: ID=' + meeting.id + ', URL=' + meeting.join_url + ', passcode=' + (meeting.password || ''));
    return meeting.join_url;

  } catch (e) {
    console.error('Zoomミーティング作成エラー:', e);
    return null;
  }
}

/**
 * Zoomミーティングを削除（キャンセル時）
 * @param {string} zoomUrl - Join URL
 * @returns {boolean} 削除成功かどうか
 */
function deleteZoomMeeting(zoomUrl) {
  if (!zoomUrl) return false;

  try {
    // Join URLからミーティングIDを抽出
    var match = zoomUrl.match(/\/j\/(\d+)/);
    if (!match) {
      console.log('Zoom URL からミーティングIDを抽出できませんでした: ' + zoomUrl);
      return false;
    }

    var meetingId = match[1];
    var token = getZoomAccessToken();

    var response = UrlFetchApp.fetch('https://api.zoom.us/v2/meetings/' + meetingId, {
      method: 'delete',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 204) {
      console.log('Zoomミーティング削除成功: ID=' + meetingId);
      return true;
    } else {
      console.error('Zoomミーティング削除エラー:', response.getResponseCode(), response.getContentText());
      return false;
    }
  } catch (e) {
    console.error('Zoomミーティング削除エラー:', e);
    return false;
  }
}

/**
 * Zoomミーティングを作成してスプレッドシートのR列に保存
 * @param {Object} data - 予約データ
 * @param {number} rowIndex - 行番号（1-based）
 * @returns {string|null} Join URL
 */
function createAndSaveZoomMeeting(data, rowIndex) {
  var joinUrl = createZoomMeeting(data);

  if (joinUrl) {
    var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
      .getSheetByName(CONFIG.SHEET_NAME);
    sheet.getRange(rowIndex, COLUMNS.ZOOM_URL + 1).setValue(joinUrl);
    data.zoomUrl = joinUrl;
    console.log('Zoom URL をR列に保存: 行' + rowIndex + ' = ' + joinUrl);
  } else {
    console.error('Zoomミーティング作成失敗のため、R列は未更新（行' + rowIndex + '）');
  }

  return joinUrl;
}

/**
 * 確定日時文字列をZoom API用ISO 8601形式に変換
 * 例: "2026/03/15 10:00〜11:30" → "2026-03-15T10:00:00"
 * 例: "2026-03-15 10:00-11:30" → "2026-03-15T10:00:00"
 * @param {string|Date} confirmedDate - 確定日時
 * @returns {string|null} ISO 8601形式の日時文字列
 */
function convertToZoomDateTime(confirmedDate) {
  if (!confirmedDate) return null;

  // Date型の場合（instanceof が失敗する場合に備え、getTime() の存在もチェック）
  if (confirmedDate instanceof Date || (typeof confirmedDate === 'object' && typeof confirmedDate.getTime === 'function')) {
    return Utilities.formatDate(confirmedDate, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss");
  }

  var str = confirmedDate.toString();

  // "YYYY/MM/DD HH:mm" or "YYYY-MM-DD HH:mm" パターン
  var match = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})\s*(\d{1,2}):(\d{2})/);
  if (match) {
    return match[1] + '-' +
      match[2].padStart(2, '0') + '-' +
      match[3].padStart(2, '0') + 'T' +
      match[4].padStart(2, '0') + ':' + match[5] + ':00';
  }

  // "Day Mon DD YYYY HH:mm:ss GMT+0900" パターン（Date.toString()形式）
  var monthMap = {Jan:'01',Feb:'02',Mar:'03',Apr:'04',May:'05',Jun:'06',Jul:'07',Aug:'08',Sep:'09',Oct:'10',Nov:'11',Dec:'12'};
  var dateStrMatch = str.match(/\w+\s+(\w+)\s+(\d{1,2})\s+(\d{4})\s+(\d{2}):(\d{2})/);
  if (dateStrMatch && monthMap[dateStrMatch[1]]) {
    return dateStrMatch[3] + '-' +
      monthMap[dateStrMatch[1]] + '-' +
      dateStrMatch[2].padStart(2, '0') + 'T' +
      dateStrMatch[4] + ':' + dateStrMatch[5] + ':00';
  }

  // 日付のみの場合は10:00をデフォルトに
  var dateMatch = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (dateMatch) {
    return dateMatch[1] + '-' +
      dateMatch[2].padStart(2, '0') + '-' +
      dateMatch[3].padStart(2, '0') + 'T10:00:00';
  }

  return null;
}

/**
 * Zoom認証情報をスクリプトプロパティに一括設定
 * GASエディタで実行してください
 * @param {string} accountId
 * @param {string} clientId
 * @param {string} clientSecret
 */
function setupZoomCredentials(accountId, clientId, clientSecret) {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    'ZOOM_ACCOUNT_ID': accountId,
    'ZOOM_CLIENT_ID': clientId,
    'ZOOM_CLIENT_SECRET': clientSecret
  });
  console.log('Zoom認証情報を設定しました');
  // 接続テストを実行
  return testZoomConnection();
}

/**
 * Zoom APIの接続テスト
 */
function testZoomConnection() {
  try {
    var token = getZoomAccessToken();
    console.log('Zoomアクセストークン取得成功');

    var response = UrlFetchApp.fetch('https://api.zoom.us/v2/users/me', {
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      var user = JSON.parse(response.getContentText());
      console.log('Zoom接続テスト成功: ' + user.email + ' (' + user.first_name + ' ' + user.last_name + ')');
      return { success: true, email: user.email };
    } else {
      console.error('Zoom APIエラー:', response.getContentText());
      return { success: false, error: response.getContentText() };
    }
  } catch (e) {
    console.error('Zoom接続テストエラー:', e);
    return { success: false, error: e.toString() };
  }
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 録画リンク自動取得・通知
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * Zoom録画一覧を取得し、スプレッドシートに未登録の録画を処理
 * 定期トリガー（1時間ごと）で実行
 */
function processZoomRecordings() {
  try {
    var token = getZoomAccessToken();
    var checkHours = (CONFIG.ZOOM && CONFIG.ZOOM.RECORDING && CONFIG.ZOOM.RECORDING.CHECK_HOURS) || 48;

    // 検索期間: 過去N時間
    var fromDate = new Date();
    fromDate.setHours(fromDate.getHours() - checkHours);
    var from = Utilities.formatDate(fromDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    var to = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

    var response = UrlFetchApp.fetch(
      'https://api.zoom.us/v2/users/me/recordings?from=' + from + '&to=' + to + '&page_size=30',
      {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) {
      console.error('Zoom録画一覧取得エラー:', response.getContentText());
      return { success: false, error: response.getContentText() };
    }

    var result = JSON.parse(response.getContentText());
    var meetings = result.meetings || [];
    if (meetings.length === 0) {
      console.log('処理対象の録画はありません');
      return { success: true, processed: 0 };
    }

    // 予約管理シートからZoom URLとミーティングIDのマッピングを構築
    var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var meetingMap = {};  // meetingId → { rowIndex, data }

    for (var i = 1; i < data.length; i++) {
      var zoomUrl = data[i][COLUMNS.ZOOM_URL];
      var recordingUrl = data[i][COLUMNS.RECORDING_URL];
      if (!zoomUrl || recordingUrl) continue;  // URL未設定 or 既に録画URL登録済み

      var match = zoomUrl.toString().match(/\/j\/(\d+)/);
      if (match) {
        meetingMap[match[1]] = { rowIndex: i + 1, zoomUrl: zoomUrl };
      }
    }

    var processedCount = 0;

    for (var m = 0; m < meetings.length; m++) {
      var meeting = meetings[m];
      var meetingId = String(meeting.id);

      // スプレッドシートに該当するミーティングがあるか
      if (!meetingMap[meetingId]) continue;

      var rowInfo = meetingMap[meetingId];
      var shareUrl = meeting.share_url || '';
      var passcode = meeting.recording_play_passcode || '';

      if (!shareUrl) {
        // share_urlがない場合、recordings settingsから取得を試みる
        shareUrl = getRecordingShareUrl(meetingId, token);
      }

      if (!shareUrl) continue;

      // 録画URLをスプレッドシートに保存
      var urlWithPasscode = shareUrl + (passcode ? '\nパスワード: ' + passcode : '');
      sheet.getRange(rowInfo.rowIndex, COLUMNS.RECORDING_URL + 1).setValue(urlWithPasscode);

      // リーダーに通知
      var rowData = getRowData(rowInfo.rowIndex);
      notifyRecordingToLeader(rowData, shareUrl, passcode);

      processedCount++;
      console.log('録画URL登録: 行' + rowInfo.rowIndex + ', meetingId=' + meetingId);
    }

    console.log('録画処理完了: ' + processedCount + '件');
    return { success: true, processed: processedCount, totalMeetings: meetings.length };

  } catch (e) {
    console.error('録画処理エラー:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * ミーティングの録画共有URLを取得
 * @param {string} meetingId - ミーティングID
 * @param {string} token - アクセストークン
 * @returns {string|null} 共有URL
 */
function getRecordingShareUrl(meetingId, token) {
  try {
    var response = UrlFetchApp.fetch(
      'https://api.zoom.us/v2/meetings/' + meetingId + '/recordings',
      {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );

    if (response.getResponseCode() !== 200) return null;

    var data = JSON.parse(response.getContentText());
    return data.share_url || null;
  } catch (e) {
    console.error('録画共有URL取得エラー:', e);
    return null;
  }
}

/**
 * リーダーに録画リンクをメール通知
 * @param {Object} rowData - getRowData形式のデータ
 * @param {string} shareUrl - 録画共有URL
 * @param {string} passcode - 録画閲覧パスワード
 */
function notifyRecordingToLeader(rowData, shareUrl, passcode) {
  if (!rowData.leader) {
    console.log('リーダー未設定のため録画通知をスキップ: ' + rowData.id);
    return;
  }

  var leaderMember = getMemberByName(rowData.leader);
  if (!leaderMember || !leaderMember.email) {
    console.log('リーダーのメールアドレスが見つかりません: ' + rowData.leader);
    return;
  }

  var subject = '【録画共有】' + (rowData.company || rowData.name) + '様 - 経営相談録画';

  var body = rowData.leader + ' 様\n\n' +
    'お疲れ様です。\n' +
    '下記の経営相談の録画が利用可能になりましたのでお知らせいたします。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 相談概要\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '申込ID　：' + (rowData.id || '') + '\n' +
    '相談日時：' + (rowData.confirmedDate || '') + '\n' +
    '相談者　：' + (rowData.name || '') + ' 様（' + (rowData.company || '') + '）\n' +
    'テーマ　：' + (rowData.theme || '') + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '■ 録画リンク\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    '視聴URL：' + shareUrl + '\n' +
    (passcode ? 'パスワード：' + passcode + '\n' : '') +
    '\n' +
    '※ 上記リンクからブラウザで視聴・ダウンロードが可能です。\n' +
    '※ Zoomアカウントは不要です。\n\n' +
    '録画内容は診断報告書の作成にご活用ください。\n\n' +
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
      GmailApp.sendEmail(adminEmail, subject,
        '※管理者控え※\n\n' + body,
        { name: CONFIG.SENDER_NAME }
      );
    }
  });

  console.log('録画通知送信完了: リーダー=' + rowData.leader + ', 申込ID=' + rowData.id);
}

/**
 * 録画チェックトリガーのセットアップ（1時間ごと）
 */
function setupRecordingCheckTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'processZoomRecordings') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('processZoomRecordings')
    .timeBased()
    .everyHours(1)
    .create();

  console.log('録画チェックトリガーをセットアップしました（1時間ごと）');
  return { success: true, message: '録画チェックトリガーを設定しました（1時間ごと）' };
}
