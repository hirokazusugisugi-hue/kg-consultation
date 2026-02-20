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
        waiting_room: true,
        join_before_host: false,
        mute_upon_entry: true,
        audio: 'voip',
        auto_recording: 'cloud'
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
    console.log('Zoomミーティング作成成功: ID=' + meeting.id + ', URL=' + meeting.join_url);
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
