/**
 * レポート配信システム
 * リーダーへのレポート依頼、アップロード受付、相談者への配信
 */

/**
 * レポート依頼を開始
 * @param {string} applicationId - 申込ID
 */
function initiateReportRequest(applicationId) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // 申込IDから行を特定
  var data = mainSheet.getDataRange().getValues();
  var rowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.ID] === applicationId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    console.log('申込ID not found: ' + applicationId);
    return;
  }

  var rowData = getRowData(rowIndex);

  // リーダーが未設定の場合はスキップ
  if (!rowData.leader) {
    console.log('リーダー未設定のためレポート依頼スキップ: ' + applicationId);
    return;
  }

  // 既にレポート依頼済みの場合はスキップ
  if (rowData.reportStatus && rowData.reportStatus !== REPORT_STATUS.NOT_REQUESTED) {
    console.log('レポート既に依頼済み: ' + applicationId + ' (' + rowData.reportStatus + ')');
    return;
  }

  // リーダーのメールアドレスを取得
  var leaderMember = getMemberByName(rowData.leader);
  if (!leaderMember || !leaderMember.email) {
    console.log('リーダーのメールアドレスが見つかりません: ' + rowData.leader);
    return;
  }

  // レポートトークン生成
  var token = generateReportToken(applicationId, rowData.leader, leaderMember.email);

  // アップロードURL
  var uploadUrl = CONFIG.CONSENT.WEB_APP_URL + '?action=report-upload&token=' + token;

  // 期限計算
  var now = new Date();
  var deadline = new Date(now.getTime() + CONFIG.REPORT.DEADLINE_DAYS * 24 * 60 * 60 * 1000);
  var deadlineStr = Utilities.formatDate(deadline, 'Asia/Tokyo', 'yyyy/MM/dd');

  // レポート管理シートに記録
  var reportSheet = ss.getSheetByName(CONFIG.REPORT_SHEET_NAME);
  if (!reportSheet) {
    setupReportSheet();
    reportSheet = ss.getSheetByName(CONFIG.REPORT_SHEET_NAME);
  }

  reportSheet.appendRow([
    applicationId,                   // A: 申込ID
    rowData.confirmedDate || '',     // B: 相談日
    rowData.company,                 // C: 相談企業
    rowData.leader,                  // D: リーダー
    leaderMember.email,              // E: リーダーメール
    now,                             // F: 依頼日時
    deadline,                        // G: 期限
    '',                              // H: アップロード日時
    '',                              // I: ファイルID
    '',                              // J: ファイルURL
    '',                              // K: 配信日時
    REPORT_STATUS.REQUESTED          // L: ステータス
  ]);

  // 予約管理シートのZ列を更新
  mainSheet.getRange(rowIndex, COLUMNS.REPORT_STATUS + 1).setValue(REPORT_STATUS.REQUESTED);

  // リーダーに依頼メール送信
  var emailBody = getReportRequestEmailBody({
    leaderName: rowData.leader,
    company: rowData.company,
    industry: rowData.industry,
    theme: rowData.theme,
    confirmedDate: rowData.confirmedDate,
    applicationId: applicationId,
    uploadUrl: uploadUrl,
    deadlineStr: deadlineStr
  });

  GmailApp.sendEmail(leaderMember.email,
    '【レポート作成依頼】' + rowData.company + '様 - 診断報告書',
    emailBody,
    { name: CONFIG.SENDER_NAME, replyTo: CONFIG.REPLY_TO }
  );

  console.log('レポート依頼送信: ' + rowData.leader + ' (' + leaderMember.email + ')');
}

/**
 * レポートトークン生成
 * @param {string} applicationId - 申込ID
 * @param {string} leaderName - リーダー名
 * @param {string} leaderEmail - リーダーメール
 * @returns {string} トークン
 */
function generateReportToken(applicationId, leaderName, leaderEmail) {
  var token = Utilities.getUuid();
  var props = PropertiesService.getScriptProperties();
  var tokenData = JSON.stringify({
    applicationId: applicationId,
    leaderName: leaderName,
    leaderEmail: leaderEmail,
    createdAt: new Date().toISOString()
  });
  props.setProperty('report_token_' + token, tokenData);
  return token;
}

/**
 * レポートトークンを検証
 * @param {string} token - トークン
 * @returns {Object|null} トークンデータ
 */
function validateReportToken(token) {
  if (!token) return null;
  var props = PropertiesService.getScriptProperties();
  var tokenDataStr = props.getProperty('report_token_' + token);
  if (!tokenDataStr) return null;
  try {
    return JSON.parse(tokenDataStr);
  } catch (e) {
    return null;
  }
}

/**
 * レポートアップロードページ生成
 * @param {Object} e - GETリクエストイベント
 * @returns {HtmlOutput} アップロードページHTML
 */
function generateReportUploadPage(e) {
  var token = e.parameter.token;
  var tokenData = validateReportToken(token);

  if (!tokenData) {
    return HtmlService.createHtmlOutput(
      '<html><body><h2>無効なリンクです</h2><p>このアップロードリンクは無効または期限切れです。</p></body></html>'
    ).setTitle('エラー - レポートアップロード')
     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 申込IDから相談情報を取得
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var allData = mainSheet.getDataRange().getValues();
  var consultData = null;
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][COLUMNS.ID] === tokenData.applicationId) {
      consultData = {
        id: allData[i][COLUMNS.ID],
        company: allData[i][COLUMNS.COMPANY],
        name: allData[i][COLUMNS.NAME],
        industry: allData[i][COLUMNS.INDUSTRY],
        theme: allData[i][COLUMNS.THEME]
      };
      break;
    }
  }

  if (!consultData) {
    return HtmlService.createHtmlOutput(
      '<html><body><h2>データが見つかりません</h2></body></html>'
    ).setTitle('エラー')
     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var html = getReportUploadPageHtml(tokenData, consultData, token);
  return HtmlService.createHtmlOutput(html)
    .setTitle('レポートアップロード - 関西学院大学 中小企業経営診断研究会')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * レポートアップロードページHTML
 */
function getReportUploadPageHtml(tokenData, consultData, token) {
  return '<!DOCTYPE html>' +
'<html lang="ja">' +
'<head>' +
'  <meta charset="UTF-8">' +
'  <meta name="viewport" content="width=device-width, initial-scale=1.0">' +
'  <title>レポートアップロード</title>' +
'  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">' +
'  <style>' +
'    * { box-sizing: border-box; margin: 0; padding: 0; }' +
'    body { font-family: "Noto Sans JP", sans-serif; background: #f5f5f7; color: #1a1a1a; line-height: 1.8; }' +
'    .header { background: #0F2350; color: #fff; padding: 2rem 0; text-align: center; }' +
'    .header h1 { font-size: 1.3rem; font-weight: 700; }' +
'    .header p { font-size: 0.85rem; opacity: 0.8; }' +
'    .container { max-width: 700px; margin: 0 auto; padding: 2rem 1.5rem; }' +
'    .card { background: #fff; border-radius: 12px; padding: 2rem; margin-bottom: 1.5rem; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }' +
'    .card h2 { font-size: 1.1rem; font-weight: 700; margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid #0F2350; }' +
'    .info-grid { display: grid; grid-template-columns: 100px 1fr; gap: 0.5rem; font-size: 0.9rem; }' +
'    .info-grid dt { font-weight: 600; color: #666; }' +
'    .upload-area { border: 2px dashed #ccc; border-radius: 12px; padding: 2rem; text-align: center; cursor: pointer; transition: all 0.3s; margin: 1rem 0; }' +
'    .upload-area:hover { border-color: #0F2350; background: #f8f9ff; }' +
'    .upload-area.has-file { border-color: #28a745; background: #f0fff4; }' +
'    .upload-area input[type="file"] { display: none; }' +
'    .file-info { font-size: 0.85rem; color: #28a745; margin-top: 0.5rem; }' +
'    .file-limit { font-size: 0.8rem; color: #888; margin-top: 0.5rem; }' +
'    .btn { width: 100%; padding: 1rem; background: #0F2350; color: #fff; border: none; border-radius: 8px; font-size: 1rem; font-weight: 700; cursor: pointer; }' +
'    .btn:hover { opacity: 0.9; }' +
'    .btn:disabled { background: #ccc; cursor: not-allowed; }' +
'    .error-msg { color: #dc3545; font-size: 0.85rem; margin-top: 0.5rem; display: none; }' +
'    .success-overlay { display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 1000; justify-content: center; align-items: center; }' +
'    .success-overlay.active { display: flex; }' +
'    .success-box { background: #fff; border-radius: 12px; padding: 3rem 2rem; text-align: center; max-width: 500px; margin: 1rem; }' +
'    .success-icon { width: 60px; height: 60px; background: #d4edda; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 1rem; font-size: 1.5rem; color: #28a745; }' +
'  </style>' +
'</head>' +
'<body>' +
'  <div class="header">' +
'    <h1>診断報告書のアップロード</h1>' +
'    <p>関西学院大学 中小企業経営診断研究会</p>' +
'  </div>' +
'  <div class="container">' +
'    <div class="card">' +
'      <h2>相談情報</h2>' +
'      <dl class="info-grid">' +
'        <dt>申込ID</dt><dd>' + consultData.id + '</dd>' +
'        <dt>企業名</dt><dd>' + consultData.company + '</dd>' +
'        <dt>業種</dt><dd>' + (consultData.industry || '') + '</dd>' +
'        <dt>テーマ</dt><dd>' + (consultData.theme || '') + '</dd>' +
'        <dt>リーダー</dt><dd>' + tokenData.leaderName + '</dd>' +
'      </dl>' +
'    </div>' +
'    <div class="card">' +
'      <h2>報告書アップロード</h2>' +
'      <p style="font-size:0.9rem; color:#666; margin-bottom:1rem;">PDF または Word ファイルをアップロードしてください。</p>' +
'      <div class="upload-area" id="uploadArea" onclick="document.getElementById(\'fileInput\').click()">' +
'        <div id="uploadLabel">ファイルを選択またはドラッグ&ドロップ</div>' +
'        <div class="file-info" id="fileInfo" style="display:none;"></div>' +
'        <input type="file" id="fileInput" accept=".pdf,.doc,.docx" onchange="handleFileSelect(this)">' +
'      </div>' +
'      <div class="file-limit">対応形式: PDF, Word (.doc, .docx) / 上限: 5MB</div>' +
'      <div id="errorMsg" class="error-msg"></div>' +
'      <button class="btn" id="submitBtn" onclick="submitReport()" disabled>アップロードして送信</button>' +
'    </div>' +
'  </div>' +
'  <div id="successOverlay" class="success-overlay">' +
'    <div class="success-box">' +
'      <div class="success-icon">&#10003;</div>' +
'      <h3>アップロード完了</h3>' +
'      <p style="margin-top:1rem; color:#666;">報告書が正常にアップロードされました。<br>相談者様へ自動的に配信されます。<br>このページを閉じていただいて結構です。</p>' +
'    </div>' +
'  </div>' +
'  <script>' +
'    var selectedFile = null;' +
'    var maxSize = 5 * 1024 * 1024;' +
'' +
'    function handleFileSelect(input) {' +
'      var file = input.files[0];' +
'      if (!file) return;' +
'      if (file.size > maxSize) {' +
'        showError("ファイルサイズが5MBを超えています（" + (file.size / 1024 / 1024).toFixed(1) + "MB）");' +
'        return;' +
'      }' +
'      var ext = file.name.split(".").pop().toLowerCase();' +
'      if (["pdf", "doc", "docx"].indexOf(ext) < 0) {' +
'        showError("対応していないファイル形式です");' +
'        return;' +
'      }' +
'      selectedFile = file;' +
'      document.getElementById("uploadArea").classList.add("has-file");' +
'      document.getElementById("uploadLabel").textContent = file.name;' +
'      document.getElementById("fileInfo").style.display = "block";' +
'      document.getElementById("fileInfo").textContent = (file.size / 1024).toFixed(0) + " KB";' +
'      document.getElementById("submitBtn").disabled = false;' +
'      hideError();' +
'    }' +
'' +
'    function showError(msg) {' +
'      var el = document.getElementById("errorMsg");' +
'      el.textContent = msg;' +
'      el.style.display = "block";' +
'    }' +
'    function hideError() {' +
'      document.getElementById("errorMsg").style.display = "none";' +
'    }' +
'' +
'    function submitReport() {' +
'      if (!selectedFile) return;' +
'      var btn = document.getElementById("submitBtn");' +
'      btn.disabled = true;' +
'      btn.textContent = "アップロード中...";' +
'' +
'      var reader = new FileReader();' +
'      reader.onload = function(e) {' +
'        var base64 = e.target.result.split(",")[1];' +
'        google.script.run' +
'          .withSuccessHandler(function(result) {' +
'            if (result.success) {' +
'              document.getElementById("successOverlay").classList.add("active");' +
'            } else {' +
'              showError(result.message || "エラーが発生しました");' +
'              btn.disabled = false;' +
'              btn.textContent = "アップロードして送信";' +
'            }' +
'          })' +
'          .withFailureHandler(function(err) {' +
'            showError("エラー: " + err.message);' +
'            btn.disabled = false;' +
'            btn.textContent = "アップロードして送信";' +
'          })' +
'          .submitReportUpload({' +
'            token: "' + token + '",' +
'            fileName: selectedFile.name,' +
'            mimeType: selectedFile.type,' +
'            base64: base64' +
'          });' +
'      };' +
'      reader.readAsDataURL(selectedFile);' +
'    }' +
'' +
'    // ドラッグ&ドロップ' +
'    var area = document.getElementById("uploadArea");' +
'    area.addEventListener("dragover", function(e) { e.preventDefault(); area.style.borderColor = "#0F2350"; });' +
'    area.addEventListener("dragleave", function() { area.style.borderColor = "#ccc"; });' +
'    area.addEventListener("drop", function(e) {' +
'      e.preventDefault();' +
'      area.style.borderColor = "#ccc";' +
'      if (e.dataTransfer.files.length > 0) {' +
'        document.getElementById("fileInput").files = e.dataTransfer.files;' +
'        handleFileSelect(document.getElementById("fileInput"));' +
'      }' +
'    });' +
'  </script>' +
'</body>' +
'</html>';
}

/**
 * レポートファイルのアップロード処理
 * @param {Object} formData - { token, fileName, mimeType, base64 }
 * @returns {Object} { success, message }
 */
function submitReportUpload(formData) {
  try {
    var tokenData = validateReportToken(formData.token);
    if (!tokenData) {
      return { success: false, message: '無効なトークンです' };
    }

    // Base64デコードしてDriveに保存
    var blob = Utilities.newBlob(
      Utilities.base64Decode(formData.base64),
      formData.mimeType,
      formData.fileName
    );

    // サイズチェック
    if (blob.getBytes().length > CONFIG.REPORT.MAX_FILE_SIZE) {
      return { success: false, message: 'ファイルサイズが5MBを超えています' };
    }

    // Driveに保存
    var folder;
    if (CONFIG.REPORT.DRIVE_FOLDER_ID) {
      folder = DriveApp.getFolderById(CONFIG.REPORT.DRIVE_FOLDER_ID);
    } else {
      folder = DriveApp.getRootFolder();
    }

    var fileName = tokenData.applicationId + '_report_' + formData.fileName;
    var file = folder.createFile(blob.setName(fileName));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileId = file.getId();
    var fileUrl = file.getUrl();

    // レポート管理シートを更新
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var reportSheet = ss.getSheetByName(CONFIG.REPORT_SHEET_NAME);
    if (reportSheet) {
      var reportData = reportSheet.getDataRange().getValues();
      for (var i = 1; i < reportData.length; i++) {
        if (reportData[i][REPORT_COLUMNS.APP_ID] === tokenData.applicationId) {
          var row = i + 1;
          reportSheet.getRange(row, REPORT_COLUMNS.UPLOAD_DATE + 1).setValue(new Date());
          reportSheet.getRange(row, REPORT_COLUMNS.FILE_ID + 1).setValue(fileId);
          reportSheet.getRange(row, REPORT_COLUMNS.FILE_URL + 1).setValue(fileUrl);
          reportSheet.getRange(row, REPORT_COLUMNS.STATUS + 1).setValue(REPORT_STATUS.UPLOADED);
          break;
        }
      }
    }

    // 予約管理シートのZ列を更新
    var mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    var mainData = mainSheet.getDataRange().getValues();
    for (var j = 1; j < mainData.length; j++) {
      if (mainData[j][COLUMNS.ID] === tokenData.applicationId) {
        mainSheet.getRange(j + 1, COLUMNS.REPORT_STATUS + 1).setValue(REPORT_STATUS.UPLOADED);
        break;
      }
    }

    // 相談者にレポート配信
    deliverReportToConsultee(tokenData.applicationId, fileUrl);

    // トークンを無効化
    PropertiesService.getScriptProperties().deleteProperty('report_token_' + formData.token);

    console.log('レポートアップロード完了: ' + tokenData.applicationId + ' (' + fileName + ')');
    return { success: true, message: 'アップロード完了' };

  } catch (error) {
    console.error('レポートアップロードエラー:', error);
    return { success: false, message: 'エラーが発生しました: ' + error.toString() };
  }
}

/**
 * 相談者へレポート配信メール送信
 * @param {string} applicationId - 申込ID
 * @param {string} fileUrl - レポートファイルURL
 */
function deliverReportToConsultee(applicationId, fileUrl) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var allData = mainSheet.getDataRange().getValues();

  var rowData = null;
  var rowIndex = -1;
  for (var i = 1; i < allData.length; i++) {
    if (allData[i][COLUMNS.ID] === applicationId) {
      rowIndex = i + 1;
      rowData = getRowData(rowIndex);
      break;
    }
  }

  if (!rowData || !rowData.email) {
    console.log('相談者情報が見つかりません: ' + applicationId);
    return;
  }

  var emailBody = getReportDeliveryEmailBody({
    name: rowData.name,
    company: rowData.company,
    theme: rowData.theme,
    confirmedDate: rowData.confirmedDate,
    applicationId: applicationId,
    fileUrl: fileUrl
  });

  GmailApp.sendEmail(rowData.email,
    '【診断報告書のお届け】' + rowData.company + '様 - 経営相談レポート',
    emailBody,
    { name: CONFIG.SENDER_NAME, replyTo: CONFIG.REPLY_TO }
  );

  // レポート管理シート・予約管理シートのステータスを「配信済」に更新
  var reportSheet = ss.getSheetByName(CONFIG.REPORT_SHEET_NAME);
  if (reportSheet) {
    var reportData = reportSheet.getDataRange().getValues();
    for (var j = 1; j < reportData.length; j++) {
      if (reportData[j][REPORT_COLUMNS.APP_ID] === applicationId) {
        reportSheet.getRange(j + 1, REPORT_COLUMNS.DELIVERY_DATE + 1).setValue(new Date());
        reportSheet.getRange(j + 1, REPORT_COLUMNS.STATUS + 1).setValue(REPORT_STATUS.DELIVERED);
        break;
      }
    }
  }

  if (rowIndex > 0) {
    mainSheet.getRange(rowIndex, COLUMNS.REPORT_STATUS + 1).setValue(REPORT_STATUS.DELIVERED);
  }

  console.log('レポート配信完了: ' + rowData.email + ' (' + applicationId + ')');
}

/**
 * レポート期限チェック & リマインド送信（日次トリガーから呼ばれる）
 */
function checkReportDeadlines() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var reportSheet = ss.getSheetByName(CONFIG.REPORT_SHEET_NAME);

  if (!reportSheet || reportSheet.getLastRow() <= 1) return;

  var now = new Date();
  var data = reportSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var status = data[i][REPORT_COLUMNS.STATUS];
    if (status !== REPORT_STATUS.REQUESTED) continue;

    var deadline = data[i][REPORT_COLUMNS.DEADLINE];
    if (!deadline) continue;

    var deadlineDate = new Date(deadline);

    // 期限超過
    if (now > deadlineDate) {
      reportSheet.getRange(i + 1, REPORT_COLUMNS.STATUS + 1).setValue(REPORT_STATUS.OVERDUE);

      // 予約管理シートも更新
      var appId = data[i][REPORT_COLUMNS.APP_ID];
      updateMainSheetReportStatus(appId, REPORT_STATUS.OVERDUE);

      // リマインドメール送信
      sendReportReminder(data[i]);

      console.log('レポート期限超過: ' + appId);
      continue;
    }

    // 期限前日リマインド
    var oneDayBefore = new Date(deadlineDate.getTime() - 24 * 60 * 60 * 1000);
    var todayStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd');
    var reminderStr = Utilities.formatDate(oneDayBefore, 'Asia/Tokyo', 'yyyy/MM/dd');

    if (todayStr === reminderStr) {
      sendReportReminder(data[i]);
      console.log('レポート期限前日リマインド: ' + data[i][REPORT_COLUMNS.APP_ID]);
    }
  }
}

/**
 * リマインドメール送信（リーダー向け）
 * @param {Array} reportRow - レポート管理シートの行データ
 */
function sendReportReminder(reportRow) {
  var leaderEmail = reportRow[REPORT_COLUMNS.LEADER_EMAIL];
  if (!leaderEmail) return;

  var deadline = reportRow[REPORT_COLUMNS.DEADLINE];
  var deadlineStr = deadline instanceof Date
    ? Utilities.formatDate(deadline, 'Asia/Tokyo', 'yyyy/MM/dd')
    : String(deadline);

  var emailBody = getReportReminderEmailBody({
    leaderName: reportRow[REPORT_COLUMNS.LEADER],
    company: reportRow[REPORT_COLUMNS.COMPANY],
    applicationId: reportRow[REPORT_COLUMNS.APP_ID],
    deadlineStr: deadlineStr
  });

  GmailApp.sendEmail(leaderEmail,
    '【リマインド】診断報告書の提出期限が近づいています',
    emailBody,
    { name: CONFIG.SENDER_NAME, replyTo: CONFIG.REPLY_TO }
  );
}

/**
 * 予約管理シートのレポート状態を更新
 * @param {string} applicationId - 申込ID
 * @param {string} status - 新しいステータス
 */
function updateMainSheetReportStatus(applicationId, status) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var mainSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var data = mainSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][COLUMNS.ID] === applicationId) {
      mainSheet.getRange(i + 1, COLUMNS.REPORT_STATUS + 1).setValue(status);
      break;
    }
  }
}

/**
 * レポート管理シートのセットアップ
 */
function setupReportSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.REPORT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.REPORT_SHEET_NAME);
  }

  var headers = ['申込ID', '相談日', '相談企業', 'リーダー', 'リーダーメール', '依頼日時', '期限', 'アップロード日時', 'ファイルID', 'ファイルURL', '配信日時', 'ステータス'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 120);   // 申込ID
  sheet.setColumnWidth(2, 120);   // 相談日
  sheet.setColumnWidth(3, 150);   // 相談企業
  sheet.setColumnWidth(4, 100);   // リーダー
  sheet.setColumnWidth(5, 250);   // リーダーメール
  sheet.setColumnWidth(6, 150);   // 依頼日時
  sheet.setColumnWidth(7, 120);   // 期限
  sheet.setColumnWidth(8, 150);   // アップロード日時
  sheet.setColumnWidth(9, 200);   // ファイルID
  sheet.setColumnWidth(10, 250);  // ファイルURL
  sheet.setColumnWidth(11, 150);  // 配信日時
  sheet.setColumnWidth(12, 100);  // ステータス

  // ステータス列にプルダウン
  var statusRange = sheet.getRange(2, REPORT_COLUMNS.STATUS + 1, 1000, 1);
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      REPORT_STATUS.NOT_REQUESTED,
      REPORT_STATUS.REQUESTED,
      REPORT_STATUS.UPLOADED,
      REPORT_STATUS.DELIVERED,
      REPORT_STATUS.OVERDUE
    ])
    .build();
  statusRange.setDataValidation(statusRule);

  sheet.setFrozenRows(1);

  console.log('レポート管理シートのセットアップが完了しました');
  return { success: true, message: 'レポート管理シートをセットアップしました' };
}

/**
 * レポート期限チェック日次トリガーのセットアップ
 */
function setupReportDeadlineTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'checkReportDeadlines') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('checkReportDeadlines')
    .timeBased()
    .everyDays(1)
    .atHour(10)
    .create();

  console.log('レポート期限チェックトリガーを設定しました（毎日10時）');
}
