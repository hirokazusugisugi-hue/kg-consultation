/**
 * オブザーバー専用ページ（拡張版）
 * 相談予定一覧（企業名・相談予定可能者・オブザーバー表示）、NDA署名アップロード
 * PDF生成はサーバーサイド（Google Docs→PDF変換）で実行
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
 * 予約済みの相談予定を取得（拡張版：メンバー分類・企業名照合・オブザーバー情報付き）
 */
function getUpcomingConsultations() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var scheduleSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  var bookingSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var ndaSheet = ss.getSheetByName(CONFIG.OBSERVER_NDA_SHEET_NAME);
  var memberSheet = ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME);

  if (!scheduleSheet || !bookingSheet) return [];

  var scheduleData = scheduleSheet.getDataRange().getValues();
  var bookingData = bookingSheet.getDataRange().getValues();
  var now = new Date();
  var results = [];

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // メンバーマスタから区分情報を取得（オブザーバー判別用）
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  var memberTypes = {};
  if (memberSheet && memberSheet.getLastRow() > 1) {
    var memberData = memberSheet.getDataRange().getValues();
    for (var m = 1; m < memberData.length; m++) {
      var mName = String(memberData[m][MEMBER_COLUMNS.NAME]).trim();
      var mType = String(memberData[m][MEMBER_COLUMNS.TYPE]).trim();
      if (mName) memberTypes[mName] = mType;
    }
  }

  // メンバーリストを相談予定可能者とオブザーバーに分類
  function splitMembers(membersStr) {
    if (!membersStr) return { consultants: '', scheduledObservers: [] };
    var names = String(membersStr).split(',').map(function(n) { return n.trim(); }).filter(Boolean);
    var consultants = [];
    var scheduledObservers = [];
    names.forEach(function(name) {
      var type = memberTypes[name] || '';
      if (type.indexOf('オブザーバー') >= 0) {
        scheduledObservers.push(name);
      } else {
        consultants.push(name);
      }
    });
    return { consultants: consultants.join(', '), scheduledObservers: scheduledObservers };
  }

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // 予約管理シートから日付→企業名のマッピングを構築
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  var companyByDate = {};
  for (var b = 1; b < bookingData.length; b++) {
    var bConfirmedDate = bookingData[b][COLUMNS.CONFIRMED_DATE];
    var bDate1 = bookingData[b][COLUMNS.DATE1];
    var bCompany = bookingData[b][COLUMNS.COMPANY];
    if (!bCompany) continue;

    if (bConfirmedDate) {
      try {
        var bcd = new Date(bConfirmedDate);
        if (!isNaN(bcd.getTime())) {
          companyByDate[Utilities.formatDate(bcd, 'Asia/Tokyo', 'yyyy/MM/dd')] = bCompany;
        }
      } catch (e) {}
    }
    if (bDate1) {
      var d1Str = convertJapaneseDateToSlash(String(bDate1));
      if (d1Str && !companyByDate[d1Str]) {
        companyByDate[d1Str] = bCompany;
      }
    }
  }

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // オブザーバーNDA提出済み情報を取得
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  var ndaObservers = {};
  if (ndaSheet && ndaSheet.getLastRow() > 1) {
    var ndaData = ndaSheet.getDataRange().getValues();
    for (var n = 1; n < ndaData.length; n++) {
      var ndaDate = ndaData[n][OBSERVER_NDA_COLUMNS.CONSULT_DATE];
      if (ndaDate) {
        var ndaDateStr = String(ndaDate).trim();
        if (!ndaObservers[ndaDateStr]) ndaObservers[ndaDateStr] = [];
        ndaObservers[ndaDateStr].push(String(ndaData[n][OBSERVER_NDA_COLUMNS.OBSERVER_NAME]));
      }
    }
  }

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // 予約管理シートから確定済み・NDA同意済みの案件を取得
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  for (var i = 1; i < bookingData.length; i++) {
    var status = bookingData[i][COLUMNS.STATUS];
    var confirmedDate = bookingData[i][COLUMNS.CONFIRMED_DATE];
    var company = bookingData[i][COLUMNS.COMPANY];
    var staff = bookingData[i][COLUMNS.STAFF];
    var appId = bookingData[i][COLUMNS.ID];
    var name = bookingData[i][COLUMNS.NAME];

    if ((status === STATUS.CONFIRMED || status === STATUS.NDA_AGREED) && confirmedDate) {
      var cDate = new Date(confirmedDate);
      if (isNaN(cDate.getTime())) continue;

      var weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
      if (cDate >= weekAgo) {
        var dateFormatted = Utilities.formatDate(cDate, 'Asia/Tokyo', 'yyyy年MM月dd日');
        var dateRaw = Utilities.formatDate(cDate, 'Asia/Tokyo', 'yyyy/MM/dd');

        // 日程設定シートから同日のメンバー情報を取得
        var membersStr = '';
        for (var j = 1; j < scheduleData.length; j++) {
          var sDate = scheduleData[j][SCHEDULE_COLUMNS.DATE];
          if (sDate) {
            try {
              var sDateFormatted = Utilities.formatDate(new Date(sDate), 'Asia/Tokyo', 'yyyy/MM/dd');
              if (sDateFormatted === dateRaw) {
                membersStr = scheduleData[j][SCHEDULE_COLUMNS.MEMBERS] || '';
                break;
              }
            } catch (e) { /* skip */ }
          }
        }

        // メンバーを分類（オブザーバーを除外）
        var split = splitMembers(membersStr);
        var ndaObs = ndaObservers[dateRaw] || [];
        // 予定オブザーバー（メンバーマスタ区分）とNDA提出済みを統合（重複排除）
        var allObservers = split.scheduledObservers.slice();
        ndaObs.forEach(function(o) {
          if (allObservers.indexOf(o) < 0) allObservers.push(o);
        });

        results.push({
          date: dateFormatted,
          dateRaw: dateRaw,
          company: company || '（未定）',
          staff: staff || '（未定）',
          members: split.consultants,
          observers: allObservers,
          ndaSubmitted: ndaObs,
          appId: appId || '',
          clientName: name || ''
        });
      }
    }
  }

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // 日程設定シートから予約済みスロットも取得
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  for (var j = 1; j < scheduleData.length; j++) {
    var bookingStatus = scheduleData[j][SCHEDULE_COLUMNS.BOOKING_STATUS];
    if (bookingStatus === '予約済み') {
      var sDate = scheduleData[j][SCHEDULE_COLUMNS.DATE];
      var sStaff = scheduleData[j][SCHEDULE_COLUMNS.STAFF];
      var sMembers = scheduleData[j][SCHEDULE_COLUMNS.MEMBERS] || '';

      if (sDate) {
        try {
          var schedDate = new Date(sDate);
          if (isNaN(schedDate.getTime())) continue;

          var weekAgo2 = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
          if (schedDate >= weekAgo2) {
            var dateStr = Utilities.formatDate(schedDate, 'Asia/Tokyo', 'yyyy年MM月dd日');
            var dateRaw2 = Utilities.formatDate(schedDate, 'Asia/Tokyo', 'yyyy/MM/dd');
            var exists = results.some(function (r) { return r.date === dateStr; });
            if (!exists) {
              // 予約管理シートから企業名を照合（企業登録なしの日はスキップ）
              var matchedCompany = companyByDate[dateRaw2];
              if (!matchedCompany) continue;

              // メンバーを分類（オブザーバーを除外）
              var split2 = splitMembers(sMembers);
              var ndaObs2 = ndaObservers[dateRaw2] || [];
              var allObservers2 = split2.scheduledObservers.slice();
              ndaObs2.forEach(function(o) {
                if (allObservers2.indexOf(o) < 0) allObservers2.push(o);
              });

              results.push({
                date: dateStr,
                dateRaw: dateRaw2,
                company: matchedCompany,
                staff: sStaff || '（未定）',
                members: split2.consultants,
                observers: allObservers2,
                ndaSubmitted: ndaObs2,
                appId: '',
                clientName: ''
              });
            }
          }
        } catch (e) { /* skip */ }
      }
    }
  }

  // 日付でソート
  results.sort(function (a, b) { return a.dateRaw > b.dateRaw ? 1 : -1; });

  return results;
}

/**
 * Drive権限テスト（GASエディタから実行して再認証を促す）
 */
function testDriveAccess() {
  var root = DriveApp.getRootFolder();
  console.log('Drive access OK: ' + root.getName());
  return { success: true, message: 'Drive権限OK' };
}

/**
 * 署名済みNDAを保存（Drive保存 → 失敗時はメール添付にフォールバック）
 * @param {Object} data - { observerName, consultDate, company, staff, signatureBase64 }
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

    var docName = 'NDA_' + data.company + '_' + data.observerName + '_' + data.consultDate.replace(/\//g, '');

    // NDA内容をHTMLで生成（署名画像をdata URI埋め込み）
    var htmlContent = buildNdaHtml(data);
    var htmlBlob = Utilities.newBlob(htmlContent, 'text/html', docName + '.html');

    // Drive保存を試行（権限がない場合はフォールバック）
    var fileId = '';
    var fileUrl = '';
    try {
      var folder;
      if (CONFIG.OBSERVER_NDA.DRIVE_FOLDER_ID) {
        folder = DriveApp.getFolderById(CONFIG.OBSERVER_NDA.DRIVE_FOLDER_ID);
      } else {
        folder = DriveApp.getRootFolder();
      }
      var file = folder.createFile(htmlBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileId = file.getId();
      fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';
      // Blob再生成（createFileで消費されるため）
      htmlBlob = Utilities.newBlob(htmlContent, 'text/html', docName + '.html');
    } catch (driveError) {
      console.log('Drive保存スキップ（権限不足）: ' + driveError.toString());
      fileId = '（メール添付のみ）';
      fileUrl = '（メール添付のみ）';
    }

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

    // 管理者にメール通知（常にHTML添付付き）
    var subject = '【NDA提出】' + data.observerName + ' - ' + data.company;
    var emailBody = 'オブザーバーから署名済みNDAが提出されました。\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      'オブザーバー：' + data.observerName + '\n' +
      '相談日：' + data.consultDate + '\n' +
      '相談企業：' + data.company + '\n' +
      '相談予定可能者：' + data.staff + '\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
      (fileUrl && fileUrl.indexOf('http') === 0 ? 'ファイル：' + fileUrl + '\n' : '') +
      '署名済みNDA（HTMLファイル）を添付しています。\n' +
      'ブラウザで開いて印刷（PDF保存）できます。\n';

    CONFIG.ADMIN_EMAILS.forEach(function (email) {
      GmailApp.sendEmail(email, subject, emailBody, {
        name: CONFIG.SENDER_NAME,
        attachments: [htmlBlob]
      });
    });

    return { success: true, message: 'NDAを提出しました', fileUrl: fileUrl };
  } catch (error) {
    console.error('NDA保存エラー:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * NDA内容をHTMLで生成（署名画像はdata URI埋め込み）
 */
function buildNdaHtml(data) {
  var esc = function(s) { return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); };
  var sigDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');

  return '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<style>' +
    'body{font-family:"Hiragino Sans","Yu Gothic","Meiryo",sans-serif;font-size:11pt;margin:40px 50px;line-height:1.7;color:#222}' +
    'h1{text-align:center;font-size:16pt;margin-bottom:2px}' +
    '.sub{text-align:center;font-size:10pt;color:#666;margin-bottom:20px}' +
    '.art{font-weight:bold;margin-top:14px}' +
    'hr{margin:20px 0;border:none;border-top:1px solid #999}' +
    '.sig{margin-top:8px}' +
    '</style></head><body>' +
    '<h1>秘密保持誓約書</h1>' +
    '<p class="sub">養成課程在学生（オブザーバー）用</p>' +
    '<p>相談予定可能者：' + esc(data.staff) + ' 殿</p>' +
    '<p>私は、「経営診断研究会 無料経営相談分科会」（以下「本分科会」といいます）が実施する無料経営相談にオブザーバーとして出席するにあたり、個人の責任として、以下の事項を遵守することを誓約いたします。</p>' +
    '<p class="art">第1条（秘密情報の定義）</p>' +
    '<p>本誓約における「秘密情報」とは、本分科会の活動を通じて知り得た、相談企業の経営・財務・技術等の情報、関係者の個人情報、および活動中に作成された相談資料・録音データ等、一切の情報を指します。</p>' +
    '<p class="art">第2条（遵守事項）</p>' +
    '<p>1. 本分科会の正規メンバー以外の第三者に、秘密情報を開示・漏洩しないこと。</p>' +
    '<p>2. 経営相談およびそれに伴う学術研究・教育以外の目的で、秘密情報を使用しないこと。</p>' +
    '<p>3. SNSやブログ等のインターネット上に、相談企業が特定できる情報や活動内容を投稿しないこと。</p>' +
    '<p>4. 学術・教育目的で事例を利用する場合は、相談企業の事前同意に基づき、企業および個人が特定されないよう厳格な匿名化・統計化処理を施すこと。</p>' +
    '<p>5. 活動終了時または研究会の指示があった際は、秘密情報を含む資料・データ等を速やかに返還または廃棄すること。</p>' +
    '<p class="art">第3条（期間および損害賠償）</p>' +
    '<p>1. 本誓約の義務は、本分科会の活動終了後および養成課程修了後も存続するものとします。</p>' +
    '<p>2. 本誓約に違反し、研究会または相談企業に損害を与えた場合は、法的責任を負うとともに、研究会の処分に従います。</p>' +
    '<hr>' +
    '<p>相談日：' + esc(data.consultDate) + '</p>' +
    '<p>相談企業名：' + esc(data.company) + '</p>' +
    '<p><strong>【誓約者】</strong></p>' +
    '<p>氏名：' + esc(data.observerName) + '</p>' +
    '<p>署名日：' + sigDate + '</p>' +
    '<div class="sig"><img src="data:image/png;base64,' + data.signatureBase64 + '" width="200" height="60"></div>' +
    '</body></html>';
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

  var headers = ['提出日時', 'オブザーバー氏名', '相談日', '相談企業名', '相談予定可能者', 'ファイルID', 'ダウンロードURL'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#0F2350')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);

  return { success: true, message: 'オブザーバーNDAシートをセットアップしました' };
}
