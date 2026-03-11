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
    .setTitle('レポートアップロード - 関西学院大学 中小企業経営診断研究会無料経営相談分科会')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * レポートアップロードページHTML
 */
function getReportUploadPageHtml(tokenData, consultData, token) {
  return '<!DOCTYPE html>\n' +
'<html lang="ja">\n' +
'<head>\n' +
'<meta charset="UTF-8">\n' +
'<meta name="viewport" content="width=device-width, initial-scale=1.0">\n' +
'<title>レポートアップロード</title>\n' +
'<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">\n' +
'<style>\n' +
'* { box-sizing: border-box; margin: 0; padding: 0; }\n' +
'body { font-family: "Noto Sans JP", sans-serif; background: #f5f5f7; color: #1a1a1a; line-height: 1.8; }\n' +
'.header { background: #0F2350; color: #fff; padding: 2rem 0; text-align: center; }\n' +
'.header h1 { font-size: 1.3rem; font-weight: 700; }\n' +
'.header p { font-size: 0.85rem; opacity: 0.8; }\n' +
'.container { max-width: 700px; margin: 0 auto; padding: 2rem 1.5rem; }\n' +
'.card { background: #fff; border-radius: 12px; padding: 2rem; margin-bottom: 1.5rem; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }\n' +
'.card h2 { font-size: 1.1rem; font-weight: 700; margin-bottom: 1rem; padding-bottom: 0.5rem; border-bottom: 2px solid #0F2350; }\n' +
'.info-grid { display: grid; grid-template-columns: 100px 1fr; gap: 0.5rem; font-size: 0.9rem; }\n' +
'.info-grid dt { font-weight: 600; color: #666; }\n' +
'.size-notice { background: #fff3cd; border: 1px solid #ffc107; border-radius: 8px; padding: 0.75rem 1rem; margin-bottom: 1rem; font-size: 0.85rem; color: #856404; }\n' +
'.size-notice strong { font-weight: 700; }\n' +
'.file-label { display: block; border: 2px dashed #ccc; border-radius: 12px; padding: 2rem; text-align: center; cursor: pointer; transition: all 0.3s; margin: 1rem 0; }\n' +
'.file-label:hover { border-color: #0F2350; background: #f8f9ff; }\n' +
'.file-label.has-file { border-color: #28a745; background: #f0fff4; }\n' +
'#fileInput { display: none; }\n' +
'.file-name { font-size: 0.95rem; font-weight: 500; }\n' +
'.file-size { font-size: 0.85rem; color: #28a745; margin-top: 0.3rem; }\n' +
'.btn-submit { width: 100%; padding: 1rem; border: none; border-radius: 8px; font-size: 1rem; font-weight: 700; cursor: pointer; color: #fff; }\n' +
'.btn-gray { background: #ccc; cursor: not-allowed; }\n' +
'.btn-blue { background: #0F2350; cursor: pointer; }\n' +
'.btn-blue:hover { opacity: 0.9; }\n' +
'.error-msg { color: #dc3545; font-size: 0.85rem; margin-top: 0.5rem; }\n' +
'.progress-box { background: #f8f9ff; border: 1px solid #0F2350; border-radius: 8px; padding: 1.5rem; margin: 1rem 0; text-align: center; }\n' +
'.progress-bar-wrap { background: #e9ecef; border-radius: 4px; height: 8px; margin: 0.75rem 0; overflow: hidden; }\n' +
'.progress-bar-fill { background: #0F2350; height: 100%; border-radius: 4px; transition: width 0.5s; width: 0%; }\n' +
'.progress-stage { font-size: 0.85rem; color: #0F2350; font-weight: 500; }\n' +
'.progress-note { font-size: 0.75rem; color: #888; margin-top: 0.5rem; }\n' +
'@keyframes spin { to { transform: rotate(360deg); } }\n' +
'.spinner { display: inline-block; width: 18px; height: 18px; border: 2px solid #ccc; border-top-color: #0F2350; border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 0.5rem; }\n' +
'.success-overlay { display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 1000; justify-content: center; align-items: center; }\n' +
'.success-overlay.active { display: flex; }\n' +
'.success-box { background: #fff; border-radius: 12px; padding: 3rem 2rem; text-align: center; max-width: 500px; margin: 1rem; }\n' +
'.success-icon { width: 60px; height: 60px; background: #d4edda; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 1rem; font-size: 1.5rem; color: #28a745; }\n' +
'</style>\n' +
'</head>\n' +
'<body>\n' +
'<div class="header">\n' +
'  <h1>診断報告書のアップロード</h1>\n' +
'  <p>関西学院大学 中小企業経営診断研究会無料経営相談分科会</p>\n' +
'</div>\n' +
'<div class="container">\n' +
'  <div class="card">\n' +
'    <h2>相談情報</h2>\n' +
'    <dl class="info-grid">\n' +
'      <dt>申込ID</dt><dd>' + consultData.id + '</dd>\n' +
'      <dt>企業名</dt><dd>' + consultData.company + '</dd>\n' +
'      <dt>業種</dt><dd>' + (consultData.industry || '') + '</dd>\n' +
'      <dt>テーマ</dt><dd>' + (consultData.theme || '') + '</dd>\n' +
'      <dt>リーダー</dt><dd>' + tokenData.leaderName + '</dd>\n' +
'    </dl>\n' +
'  </div>\n' +
'  <div class="card">\n' +
'    <h2>報告書アップロード</h2>\n' +
'    <div class="size-notice">\n' +
'      <strong>対応形式:</strong> PDF, Word (.doc, .docx)<br>\n' +
'      <strong>最大ファイルサイズ: 5MB</strong>（5MBを超える場合はファイルを圧縮するか、画像を縮小してください）\n' +
'    </div>\n' +
'    <form id="uploadForm">\n' +
'      <input type="hidden" name="token" value="' + token + '">\n' +
'      <label class="file-label" id="fileLabel" for="fileInput">\n' +
'        <div id="fileLabelText">ここをタップしてファイルを選択</div>\n' +
'        <div id="fileNameDisplay" class="file-name" style="display:none;"></div>\n' +
'        <div id="fileSizeDisplay" class="file-size" style="display:none;"></div>\n' +
'      </label>\n' +
'      <input type="file" id="fileInput" name="reportFile" accept=".pdf,.doc,.docx">\n' +
'    </form>\n' +
'    <div id="progressBox" class="progress-box" style="display:none;">\n' +
'      <div class="progress-stage"><span class="spinner"></span><span id="progressText">準備中...</span></div>\n' +
'      <div class="progress-bar-wrap"><div class="progress-bar-fill" id="progressBar"></div></div>\n' +
'      <div class="progress-note" id="progressNote">このページを閉じないでください</div>\n' +
'    </div>\n' +
'    <div id="errorMsg" class="error-msg" style="display:none;"></div>\n' +
'    <button type="button" class="btn-submit btn-gray" id="submitBtn" onclick="doSubmit()">アップロードして送信</button>\n' +
'  </div>\n' +
'</div>\n' +
'<div id="successOverlay" class="success-overlay">\n' +
'  <div class="success-box">\n' +
'    <div class="success-icon">&#10003;</div>\n' +
'    <h3>アップロード完了</h3>\n' +
'    <p style="margin-top:1rem; color:#666;">報告書が正常にアップロードされました。<br>相談者様へ自動的に配信されます。<br>このページを閉じていただいて結構です。</p>\n' +
'  </div>\n' +
'</div>\n' +
'<script>\n' +
'var MAXSIZE = 5 * 1024 * 1024;\n' +
'var uploading = false;\n' +
'var fileReady = false;\n' +
'var uploadTimer = null;\n' +
'\n' +
'// ポーリング: 500msごとにファイル選択状態を監視（イベントハンドラが動かない場合のフォールバック）\n' +
'setInterval(function() {\n' +
'  try {\n' +
'    if (uploading) return;\n' +
'    var fi = document.getElementById("fileInput");\n' +
'    if (!fi) return;\n' +
'    var hasFile = fi.files && fi.files.length > 0;\n' +
'    if (hasFile && !fileReady) {\n' +
'      fileReady = true;\n' +
'      var f = fi.files[0];\n' +
'      var label = document.getElementById("fileLabel");\n' +
'      var nameEl = document.getElementById("fileNameDisplay");\n' +
'      var sizeEl = document.getElementById("fileSizeDisplay");\n' +
'      var textEl = document.getElementById("fileLabelText");\n' +
'      var btn = document.getElementById("submitBtn");\n' +
'      if (label) label.className = "file-label has-file";\n' +
'      if (textEl) textEl.style.display = "none";\n' +
'      if (nameEl) { nameEl.textContent = f.name; nameEl.style.display = "block"; }\n' +
'      if (sizeEl) {\n' +
'        var mb = f.size / 1024 / 1024;\n' +
'        sizeEl.textContent = mb >= 1 ? mb.toFixed(1) + " MB" : Math.round(f.size / 1024) + " KB";\n' +
'        sizeEl.style.display = "block";\n' +
'      }\n' +
'      if (btn) { btn.className = "btn-submit btn-blue"; }\n' +
'    } else if (!hasFile && fileReady) {\n' +
'      fileReady = false;\n' +
'      var btn2 = document.getElementById("submitBtn");\n' +
'      if (btn2) { btn2.className = "btn-submit btn-gray"; }\n' +
'      var textEl2 = document.getElementById("fileLabelText");\n' +
'      if (textEl2) textEl2.style.display = "block";\n' +
'      var nameEl2 = document.getElementById("fileNameDisplay");\n' +
'      if (nameEl2) nameEl2.style.display = "none";\n' +
'      var sizeEl2 = document.getElementById("fileSizeDisplay");\n' +
'      if (sizeEl2) sizeEl2.style.display = "none";\n' +
'    }\n' +
'  } catch(e) {}\n' +
'}, 500);\n' +
'\n' +
'// 送信ボタンクリック（ポーリングで有効/無効が切り替わるので、ここでもバリデーション）\n' +
'function doSubmit() {\n' +
'  try {\n' +
'    if (uploading) return;\n' +
'    var fi = document.getElementById("fileInput");\n' +
'    if (!fi || !fi.files || !fi.files[0]) {\n' +
'      showErr("ファイルを選択してからボタンを押してください。");\n' +
'      return;\n' +
'    }\n' +
'    var f = fi.files[0];\n' +
'    if (f.size > MAXSIZE) {\n' +
'      showErr("ファイルサイズが5MBを超えています（" + (f.size/1024/1024).toFixed(1) + "MB）。圧縮してから再度お試しください。");\n' +
'      return;\n' +
'    }\n' +
'    var ext = f.name.split(".").pop().toLowerCase();\n' +
'    if (ext !== "pdf" && ext !== "doc" && ext !== "docx") {\n' +
'      showErr("対応していない形式です。PDF または Word を選択してください。");\n' +
'      return;\n' +
'    }\n' +
'    uploading = true;\n' +
'    var btn = document.getElementById("submitBtn");\n' +
'    btn.className = "btn-submit btn-gray";\n' +
'    btn.textContent = "処理中...";\n' +
'    hideErr();\n' +
'    showProgress("サーバーに送信中...", 30, "このページを閉じないでください");\n' +
'\n' +
'    var startTime = Date.now();\n' +
'    uploadTimer = setInterval(function() {\n' +
'      var sec = Math.floor((Date.now() - startTime) / 1000);\n' +
'      if (sec >= 15 && sec < 60) showProgress("サーバーで処理中...", 60, "しばらくお待ちください（" + sec + "秒経過）");\n' +
'      else if (sec >= 60) showProgress("処理に時間がかかっています...", 70, "もう少しお待ちください（" + sec + "秒経過）");\n' +
'    }, 5000);\n' +
'\n' +
'    google.script.run\n' +
'      .withSuccessHandler(function(r) {\n' +
'        clearInterval(uploadTimer);\n' +
'        if (r && r.success) {\n' +
'          showProgress("完了!", 100, "");\n' +
'          setTimeout(function() { document.getElementById("successOverlay").className = "success-overlay active"; }, 500);\n' +
'        } else {\n' +
'          showErr(r && r.message ? r.message : "エラーが発生しました。もう一度お試しください。");\n' +
'          resetState();\n' +
'        }\n' +
'      })\n' +
'      .withFailureHandler(function(e) {\n' +
'        clearInterval(uploadTimer);\n' +
'        showErr("エラー: " + (e && e.message ? e.message : "不明なエラー") + "。もう一度お試しください。");\n' +
'        resetState();\n' +
'      })\n' +
'      .submitReportForm(document.getElementById("uploadForm"));\n' +
'  } catch(ex) {\n' +
'    showErr("エラーが発生しました: " + ex);\n' +
'    resetState();\n' +
'  }\n' +
'}\n' +
'\n' +
'function resetState() {\n' +
'  uploading = false;\n' +
'  fileReady = false;\n' +
'  var btn = document.getElementById("submitBtn");\n' +
'  btn.className = "btn-submit btn-gray";\n' +
'  btn.textContent = "アップロードして送信";\n' +
'  document.getElementById("progressBox").style.display = "none";\n' +
'}\n' +
'\n' +
'function showProgress(text, pct, note) {\n' +
'  document.getElementById("progressBox").style.display = "block";\n' +
'  document.getElementById("progressText").textContent = text;\n' +
'  document.getElementById("progressBar").style.width = pct + "%";\n' +
'  if (note) document.getElementById("progressNote").textContent = note;\n' +
'}\n' +
'\n' +
'function showErr(m) { var e = document.getElementById("errorMsg"); e.textContent = m; e.style.display = "block"; }\n' +
'function hideErr() { document.getElementById("errorMsg").style.display = "none"; }\n' +
'</script>\n' +
'</body>\n' +
'</html>';
}

/**
 * フォーム経由のレポートアップロード処理（Form Element方式）
 * google.script.run.submitReportForm(formElement) で呼ばれる
 * @param {Object} formData - { token: string, reportFile: Blob }
 * @returns {Object} { success, message }
 */
function submitReportForm(formData) {
  var steps = [];
  try {
    // Step 1: トークン検証
    var token = formData.token;
    var tokenData = validateReportToken(token);
    if (!tokenData) {
      return { success: false, message: '無効なリンクです。メール内のリンクから再度アクセスしてください。' };
    }
    steps.push('token_ok');

    // Step 2: Blob取得
    var fileBlob = formData.reportFile;
    if (!fileBlob) {
      return { success: false, message: 'ファイルが選択されていません。ページをリロードして再度お試しください。' };
    }
    steps.push('blob_ok');

    // Step 3: ファイル情報取得（defensiveに）
    var originalName = 'report.pdf';
    try { originalName = fileBlob.getName() || 'report.pdf'; } catch(e) { steps.push('getName_err'); }

    var mimeType = 'application/pdf';
    try { mimeType = fileBlob.getContentType() || 'application/pdf'; } catch(e) { steps.push('getContentType_err'); }

    // MIMEタイプ補正
    if (!mimeType || mimeType === 'application/octet-stream') {
      var ext = originalName.split('.').pop().toLowerCase();
      var mimeMap = { 'pdf': 'application/pdf', 'doc': 'application/msword', 'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
      mimeType = mimeMap[ext] || 'application/octet-stream';
    }
    steps.push('fileinfo_ok:' + originalName);

    // Step 4: Driveフォルダ取得
    var folder;
    try {
      folder = getDriveFolder('DRIVE_FOLDER_REPORT', CONFIG.REPORT.DRIVE_FOLDER_ID);
      steps.push('folder_ok');
    } catch (folderErr) {
      // フォールバック: ルートフォルダ
      folder = DriveApp.getRootFolder();
      steps.push('folder_fallback');
    }

    // Step 5: ファイル保存
    var fileName = tokenData.applicationId + '_report_' + originalName;
    var file;
    try {
      fileBlob.setName(fileName);
    } catch(e) {
      // setNameが使えない場合、newBlobで作り直す
      var bytes = fileBlob.getBytes();
      fileBlob = Utilities.newBlob(bytes, mimeType, fileName);
      steps.push('blob_recreated');
    }

    try {
      file = folder.createFile(fileBlob);
      steps.push('file_created');
    } catch (createErr) {
      console.error('createFile error:', createErr);
      return { success: false, message: 'ファイルの保存に失敗しました（step:' + steps.join(',') + ' err:' + createErr.toString() + '）' };
    }

    // Step 6: 共有設定（失敗しても続行）
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      steps.push('sharing_ok');
    } catch (shareErr) {
      console.log('setSharing skipped:', shareErr);
      steps.push('sharing_skipped');
    }

    var fileId = file.getId();
    var fileUrl = file.getUrl();
    steps.push('url_ok');

    // Step 7: レポート管理シートを更新
    try {
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
      steps.push('sheets_ok');
    } catch (sheetErr) {
      console.error('Sheet update error:', sheetErr);
      steps.push('sheets_err');
      // シート更新失敗してもファイルは保存済みなので続行
    }

    // Step 8: 相談者にレポート配信
    try {
      deliverReportToConsultee(tokenData.applicationId, fileUrl);
      steps.push('email_ok');
    } catch (emailErr) {
      console.error('Email delivery error:', emailErr);
      steps.push('email_err');
      // メール失敗してもアップロード自体は成功
    }

    // Step 9: トークンを無効化
    try {
      PropertiesService.getScriptProperties().deleteProperty('report_token_' + token);
    } catch(e) {}

    console.log('レポートアップロード完了（Form方式）: ' + tokenData.applicationId + ' steps:' + steps.join(','));
    return { success: true, message: 'アップロード完了' };

  } catch (error) {
    console.error('レポートアップロードエラー:', error, 'steps:', steps.join(','));
    return { success: false, message: 'エラーが発生しました（step:' + steps.join(',') + ' err:' + error.toString() + '）' };
  }
}

/**
 * レポートファイルのアップロード処理（旧Base64方式・後方互換用）
 * @param {Object} formData - { token, fileName, mimeType, base64 }
 * @returns {Object} { success, message }
 */
function submitReportUpload(formData) {
  try {
    var tokenData = validateReportToken(formData.token);
    if (!tokenData) {
      return { success: false, message: '無効なリンクです。メール内のリンクから再度アクセスしてください。' };
    }

    // MIMEタイプのフォールバック
    var mimeType = formData.mimeType;
    if (!mimeType || mimeType === 'application/octet-stream') {
      var ext = (formData.fileName || '').split('.').pop().toLowerCase();
      var mimeMap = { 'pdf': 'application/pdf', 'doc': 'application/msword', 'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
      mimeType = mimeMap[ext] || 'application/octet-stream';
    }

    // Base64デコードしてDriveに保存
    var decoded;
    try {
      decoded = Utilities.base64Decode(formData.base64);
    } catch (decodeErr) {
      console.error('Base64デコードエラー:', decodeErr);
      return { success: false, message: 'ファイルの処理に失敗しました。ファイルが破損していないか確認してください。' };
    }

    var blob = Utilities.newBlob(decoded, mimeType, formData.fileName);

    // サイズチェック
    var fileSizeMB = (blob.getBytes().length / 1024 / 1024).toFixed(1);
    if (blob.getBytes().length > CONFIG.REPORT.MAX_FILE_SIZE) {
      return { success: false, message: 'ファイルサイズが5MBを超えています（' + fileSizeMB + 'MB）。圧縮してから再度お試しください。' };
    }

    // Driveに保存（ScriptProperties優先、CONFIG fallback）
    var folder = getDriveFolder('DRIVE_FOLDER_REPORT', CONFIG.REPORT.DRIVE_FOLDER_ID);

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
    console.error('レポートアップロードエラー:', error, 'applicationId:', tokenData ? tokenData.applicationId : 'unknown');
    var errMsg = error.toString();
    if (errMsg.indexOf('Drive') >= 0 || errMsg.indexOf('folder') >= 0) {
      return { success: false, message: 'ファイルの保存先にアクセスできませんでした。管理者にお問い合わせください。' };
    }
    if (errMsg.indexOf('quota') >= 0 || errMsg.indexOf('limit') >= 0) {
      return { success: false, message: 'システムの処理上限に達しました。しばらく時間をおいてから再度お試しください。' };
    }
    return { success: false, message: 'エラーが発生しました。もう一度お試しください。（詳細: ' + errMsg + '）' };
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
    { name: CONFIG.SENDER_NAME, replyTo: CONFIG.REPLY_TO, cc: 'hirokazusugisugi@gmail.com' }
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

/**
 * OAuth認可トリガー（GASエディタから1回実行して認可ダイアログを表示させる）
 * DriveApp/GmailApp/SpreadsheetAppの全スコープを認可させるために、
 * 各サービスのメソッドを1回ずつ呼ぶ。
 */
function authorizeDriveAccess() {
  // DriveApp認可
  var rootFolder = DriveApp.getRootFolder();
  console.log('Drive認可OK: ルートフォルダ = ' + rootFolder.getName());

  // テストファイル作成＆削除で createFile 権限を確認
  var testBlob = Utilities.newBlob('auth test', 'text/plain', 'auth_test.txt');
  var testFile = rootFolder.createFile(testBlob);
  console.log('createFile認可OK: ' + testFile.getName());
  testFile.setTrashed(true);
  console.log('テストファイル削除済み');

  // SpreadsheetApp認可
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  console.log('Spreadsheet認可OK: ' + ss.getName());

  // GmailApp認可
  var drafts = GmailApp.getDrafts();
  console.log('Gmail認可OK: 下書き数 = ' + drafts.length);

  console.log('=== 全認可完了 ===');
  return '全認可完了。Web Appを再デプロイしてください。';
}
