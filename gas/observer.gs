/**
 * オブザーバー専用ページ
 * 相談予定一覧表示、NDAダウンロード（自動入力）、署名アップロード
 */

/**
 * オブザーバー専用ページを生成
 */
function generateObserverPage(e) {
  var schedules = getUpcomingConsultations();
  var html = getObserverPageHtml(schedules);
  return HtmlService.createHtmlOutput(html)
    .setTitle('オブザーバー専用ページ - 関西学院大学 中小企業経営診断研究会')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 予約済みの相談予定を取得
 */
function getUpcomingConsultations() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  var bookingSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!scheduleSheet || !bookingSheet) return [];

  var scheduleData = scheduleSheet.getDataRange().getValues();
  var bookingData = bookingSheet.getDataRange().getValues();
  var now = new Date();
  var results = [];

  // 予約管理シートから確定済みの案件を取得
  for (var i = 1; i < bookingData.length; i++) {
    var status = bookingData[i][COLUMNS.STATUS];
    var confirmedDate = bookingData[i][COLUMNS.CONFIRMED_DATE];
    var company = bookingData[i][COLUMNS.COMPANY];
    var staff = bookingData[i][COLUMNS.STAFF];
    var appId = bookingData[i][COLUMNS.ID];
    var name = bookingData[i][COLUMNS.NAME];

    if ((status === '確定' || status === 'NDA同意済') && confirmedDate) {
      var cDate = new Date(confirmedDate);
      // 過去7日以内〜未来の予定を表示
      var weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      if (cDate >= weekAgo) {
        results.push({
          date: Utilities.formatDate(cDate, 'Asia/Tokyo', 'yyyy年MM月dd日'),
          dateRaw: Utilities.formatDate(cDate, 'Asia/Tokyo', 'yyyy/MM/dd'),
          company: company || '（未定）',
          staff: staff || '（未定）',
          appId: appId,
          clientName: name || ''
        });
      }
    }
  }

  // 日程設定シートから予約済みスロットも取得
  for (var j = 1; j < scheduleData.length; j++) {
    var bookingStatus = scheduleData[j][SCHEDULE_COLUMNS.BOOKING_STATUS];
    if (bookingStatus === '予約済み') {
      var sDate = scheduleData[j][SCHEDULE_COLUMNS.DATE];
      var sStaff = scheduleData[j][SCHEDULE_COLUMNS.STAFF];
      var sTime = scheduleData[j][SCHEDULE_COLUMNS.TIME];

      if (sDate) {
        var schedDate = new Date(sDate);
        var weekAgo2 = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
        if (schedDate >= weekAgo2) {
          // 既にresultsにある日付は除外
          var dateStr = Utilities.formatDate(schedDate, 'Asia/Tokyo', 'yyyy年MM月dd日');
          var exists = results.some(function(r) { return r.date === dateStr; });
          if (!exists) {
            results.push({
              date: dateStr,
              dateRaw: Utilities.formatDate(schedDate, 'Asia/Tokyo', 'yyyy/MM/dd'),
              company: '（予約管理シート参照）',
              staff: sStaff || '（未定）',
              appId: '',
              clientName: ''
            });
          }
        }
      }
    }
  }

  // 日付でソート
  results.sort(function(a, b) { return a.dateRaw > b.dateRaw ? 1 : -1; });

  return results;
}

/**
 * 署名済みNDAをDriveに保存
 * @param {Object} data - { observerName, consultDate, company, staff, pdfBase64 }
 * @returns {Object} 結果
 */
function saveSignedNda(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.OBSERVER_NDA_SHEET_NAME);

    if (!sheet) {
      setupObserverNdaSheet();
      sheet = ss.getSheetByName(CONFIG.OBSERVER_NDA_SHEET_NAME);
    }

    // Base64からBlobを作成
    var pdfBytes = Utilities.base64Decode(data.pdfBase64);
    var fileName = 'NDA_' + data.company + '_' + data.observerName + '_' + data.consultDate.replace(/\//g, '') + '.pdf';
    var blob = Utilities.newBlob(pdfBytes, 'application/pdf', fileName);

    // Driveに保存
    var folder;
    if (CONFIG.OBSERVER_NDA.DRIVE_FOLDER_ID) {
      folder = DriveApp.getFolderById(CONFIG.OBSERVER_NDA.DRIVE_FOLDER_ID);
    } else {
      folder = DriveApp.getRootFolder();
    }
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    var fileId = file.getId();
    var fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

    // スプレッドシートに記録
    sheet.appendRow([
      new Date(),
      data.observerName,
      data.consultDate,
      data.company,
      data.staff,
      fileId,
      fileUrl
    ]);

    // 管理者に通知
    var subject = '【NDA提出】' + data.observerName + ' - ' + data.company;
    var body = 'オブザーバーから署名済みNDAが提出されました。\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      'オブザーバー：' + data.observerName + '\n' +
      '相談日：' + data.consultDate + '\n' +
      '相談企業：' + data.company + '\n' +
      '担当者：' + data.staff + '\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      'ファイル：' + fileUrl + '\n';

    CONFIG.ADMIN_EMAILS.forEach(function(email) {
      GmailApp.sendEmail(email, subject, body, { name: CONFIG.SENDER_NAME });
    });

    return { success: true, message: 'NDAを提出しました', fileUrl: fileUrl };
  } catch (error) {
    console.error('NDA保存エラー:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * オブザーバーNDA管理シートのセットアップ
 */
function setupObserverNdaSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.OBSERVER_NDA_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.OBSERVER_NDA_SHEET_NAME);
  }

  var headers = ['提出日時', 'オブザーバー氏名', '相談日', '相談企業名', '相談担当者', 'ファイルID', 'ダウンロードURL'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#0F2350')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  return { success: true, message: 'オブザーバーNDAシートをセットアップしました' };
}
