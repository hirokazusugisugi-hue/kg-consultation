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

    // 申込概要をDriveに自動保存
    saveApplicationSummaryToDrive(data, applicationId);

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

    // メンバー情報取得（LP用）
    if (action === 'members') {
      var allMembers = getAllMembers();
      var lpMembers = allMembers
        .filter(function(m) { return m.type !== '顧問' && m.active !== false; })
        .map(function(m) {
          var isObserver = m.term === '3期' || m.term === '4期';
          return {
            name: m.name,
            term: m.term,
            cert: isObserver ? '' : m.cert,
            type: m.type,
            titles: isObserver ? '' : m.titles,
            specialties: isObserver ? '' : m.specialties,
            themes: isObserver ? '' : m.themes
          };
        });
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, members: lpMembers }))
        .setMimeType(ContentService.MimeType.JSON);
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
    // Zoom API 診断・リトライ
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // Zoom API接続診断
    if (action === 'zoom-status') {
      var zoomResult = testZoomConnection();
      var hasConfig = !!(CONFIG.ZOOM && CONFIG.ZOOM.ACCOUNT_ID);
      var props = PropertiesService.getScriptProperties();
      var hasProp = !!props.getProperty('ZOOM_ACCOUNT_ID');
      return ContentService
        .createTextOutput(JSON.stringify({
          success: zoomResult.success,
          configSet: hasConfig,
          scriptPropertySet: hasProp,
          zoomEmail: zoomResult.email || null,
          error: zoomResult.error || null
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // シートデータ読み取り（管理用）
    if (action === 'read-sheet') {
      var sheetName = e.parameter.sheet || CONFIG.SHEET_NAME;
      var readSheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(sheetName);
      if (!readSheet) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'シート "' + sheetName + '" が見つかりません' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var allData = readSheet.getDataRange().getValues();
      var rows = [];
      for (var ri = 0; ri < allData.length; ri++) {
        var rowArr = [];
        for (var ci = 0; ci < allData[ri].length; ci++) {
          var val = allData[ri][ci];
          rowArr.push((val instanceof Date)
            ? Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
            : (val !== null && val !== undefined ? val.toString() : ''));
        }
        rows.push(rowArr);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, sheet: sheetName, totalRows: rows.length, rows: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 行データ診断（管理用）
    if (action === 'row-info') {
      var infoRow = parseInt(e.parameter.row);
      if (!infoRow || infoRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var infoData = getRowData(infoRow);
      return ContentService
        .createTextOutput(JSON.stringify({
          row: infoRow,
          id: infoData.id || null,
          name: infoData.name || null,
          company: infoData.company || null,
          method: infoData.method || null,
          status: infoData.status || null,
          confirmedDate: infoData.confirmedDate ? infoData.confirmedDate.toString() : null,
          zoomUrl: infoData.zoomUrl || null,
          email: infoData.email || null,
          staff: infoData.staff || null,
          leader: infoData.leader || null
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Zoom URL再作成＋確定メール再送（管理用）
    if (action === 'retry-zoom') {
      var retryRow = parseInt(e.parameter.row);
      if (!retryRow || retryRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row パラメータが必要です（2以上の行番号）' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var retryData = getRowData(retryRow);
      if (!retryData.id) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '行 ' + retryRow + ' にデータがありません' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var retryMethod = retryData.method || '';
      var retryIsOnline = retryMethod.indexOf('オンライン') >= 0 || retryMethod.indexOf('Zoom') >= 0 || retryMethod.indexOf('zoom') >= 0;
      if (!retryIsOnline) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'オンライン相談ではありません（method: ' + retryMethod + '）' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // 確定日時チェック
      if (!retryData.confirmedDate) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '確定日時（Q列）が未設定です。スプレッドシートで確定日時を入力してからリトライしてください。', id: retryData.id }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // Zoom作成
      var retryZoomUrl = retryData.zoomUrl;
      if (!retryZoomUrl) {
        retryZoomUrl = createAndSaveZoomMeeting(retryData, retryRow);
      }
      if (!retryZoomUrl) {
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            message: 'Zoomミーティング作成に失敗しました。',
            confirmedDate: retryData.confirmedDate ? retryData.confirmedDate.toString() : null,
            convertedTime: convertToZoomDateTime(retryData.confirmedDate)
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // 確定メール再送
      sendConfirmedEmail(retryData);
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          message: 'Zoom URL作成＋確定メール再送完了',
          zoomUrl: retryZoomUrl,
          email: retryData.email,
          applicationId: retryData.id
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // セル書き込み（管理用）
    if (action === 'write-cell') {
      var wcRow = parseInt(e.parameter.row);
      var wcCol = parseInt(e.parameter.col);
      var wcVal = e.parameter.value || '';
      var wcSheet = e.parameter.sheet || CONFIG.SHEET_NAME;
      var wcFormat = e.parameter.format || '';
      if (!wcRow || !wcCol) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row, col パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var ws = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(wcSheet);
      if (!ws) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'シート "' + wcSheet + '" が見つかりません' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var cell = ws.getRange(wcRow, wcCol);
      cell.setValue(wcVal);
      // format=header の場合、隣接するヘッダーセルの書式をコピー
      if (wcFormat === 'header' && wcCol > 1) {
        var refCell = ws.getRange(wcRow, wcCol - 1);
        cell.setBackground(refCell.getBackground());
        cell.setFontColor(refCell.getFontColor());
        cell.setFontWeight(refCell.getFontWeight());
        cell.setFontSize(refCell.getFontSize());
        cell.setHorizontalAlignment(refCell.getHorizontalAlignment());
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, sheet: wcSheet, row: wcRow, col: wcCol, value: wcVal, formatted: wcFormat === 'header' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 確定通知の時限送信設定（管理用）
    if (action === 'schedule-notify') {
      var notifyRow = parseInt(e.parameter.row);
      var notifyHour = parseInt(e.parameter.hour || '8');
      if (!notifyRow || notifyRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      // 行番号をプロパティに保存
      PropertiesService.getScriptProperties().setProperty('PENDING_NOTIFY_ROW', notifyRow.toString());
      // 既存の同名トリガーを削除
      ScriptApp.getProjectTriggers().forEach(function(t) {
        if (t.getHandlerFunction() === 'sendScheduledStaffNotification') {
          ScriptApp.deleteTrigger(t);
        }
      });
      // トリガー日時を計算（days=0で当日、days=1で翌日、デフォルト0）
      var daysOffset = parseInt(e.parameter.days || '0');
      var targetDate = new Date();
      targetDate.setDate(targetDate.getDate() + daysOffset);
      targetDate.setHours(notifyHour, 0, 0, 0);
      ScriptApp.newTrigger('sendScheduledStaffNotification')
        .timeBased()
        .at(targetDate)
        .create();
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          message: '確定通知を ' + Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + ' に送信予約しました',
          row: notifyRow,
          scheduledAt: Utilities.formatDate(targetDate, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 即時確定通知送信（管理用）
    // ?action=send-notify-now&row=2  (全員に送信)
    // ?action=send-notify-now&row=2&testTo=杉山 宏和  (テスト: 指定者のみ)
    if (action === 'send-notify-now') {
      var notifyRow2 = parseInt(e.parameter.row);
      if (!notifyRow2 || notifyRow2 < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var testTo = e.parameter.testTo || '';
      var debug = [];
      try {
        var data2 = getRowData(notifyRow2);
        debug.push('id=' + data2.id + ', confirmedDate=' + data2.confirmedDate);

        var memberList2 = getScheduleMembersForDate_(data2.confirmedDate);
        debug.push('memberCount=' + (memberList2 ? memberList2.length : 0));

        // メール件名・本文を構築
        var emailResult = buildStaffNotificationEmail_(data2);
        var senderName2 = emailResult.senderName || CONFIG.SENDER_NAME;

        var sentCount = 0;
        if (testTo) {
          // テストモード: 指定者のみ
          var tm = getMemberByName(testTo);
          if (tm && tm.email) {
            debug.push('TEST sending to ' + testTo + ' <' + tm.email + '>');
            GmailApp.sendEmail(tm.email, emailResult.subject, emailResult.body, { name: senderName2 });
            sentCount++;
          } else {
            debug.push('TEST target not found: ' + testTo);
          }
        } else if (memberList2 && memberList2.length > 0) {
          memberList2.forEach(function(cm) {
            var m2 = getMemberByName(cm.name);
            if (m2 && m2.email) {
              debug.push('sending to ' + cm.name + ' <' + m2.email + '>');
              GmailApp.sendEmail(m2.email, emailResult.subject, emailResult.body, { name: senderName2 });
              sentCount++;
            }
          });
        } else {
          debug.push('fallback to admin');
          CONFIG.ADMIN_EMAILS.forEach(function(adminEmail) {
            GmailApp.sendEmail(adminEmail, emailResult.subject, emailResult.body, { name: senderName2 });
            sentCount++;
          });
        }
        debug.push('sentCount=' + sentCount);

        return ContentService
          .createTextOutput(JSON.stringify({ success: true, sentCount: sentCount, debug: debug }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        debug.push('ERROR: ' + err.message);
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message, debug: debug }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // トリガー一覧取得（管理用）
    if (action === 'list-triggers') {
      var triggers = ScriptApp.getProjectTriggers();
      var result = triggers.map(function(t) {
        return {
          handler: t.getHandlerFunction(),
          type: t.getEventType().toString(),
          source: t.getTriggerSource().toString()
        };
      });
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, count: result.length, triggers: result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // スケジュールシートの重複行を削除（管理用）
    if (action === 'dedup-schedule') {
      try {
        var targetMonth = e.parameter.month ? parseInt(e.parameter.month) : null;
        var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        var sheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
        if (!sheet) throw new Error('日程設定シートが見つかりません');

        var data = sheet.getDataRange().getValues();
        var seen = {};
        var deleteRows = [];

        for (var i = 1; i < data.length; i++) {
          var dateVal = data[i][SCHEDULE_COLUMNS.DATE];
          if (!dateVal) continue;
          var date = dateVal instanceof Date ? dateVal : new Date(dateVal);
          if (targetMonth && (date.getMonth() + 1) !== targetMonth) continue;

          var timeVal = data[i][SCHEDULE_COLUMNS.TIME];
          var timeStr;
          if (timeVal instanceof Date) {
            timeStr = Utilities.formatDate(timeVal, 'Asia/Tokyo', 'HH:mm');
          } else {
            timeStr = String(timeVal);
          }

          var key = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd') + '_' + timeStr;
          if (seen[key]) {
            deleteRows.push(i + 1);
          } else {
            seen[key] = true;
          }
        }

        // 下から削除（行番号がずれないように）
        deleteRows.reverse();
        deleteRows.forEach(function(row) {
          sheet.deleteRow(row);
        });

        return ContentService
          .createTextOutput(JSON.stringify({ success: true, deletedRows: deleteRows.length, rows: deleteRows.reverse() }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // オブザーバーNDAリマインドメール送信（管理用）
    // ?action=send-observer-nda-reminder&name=高山 佳樹
    if (action === 'send-observer-nda-reminder') {
      try {
        var obsName = e.parameter.name;
        if (!obsName) {
          return ContentService
            .createTextOutput(JSON.stringify({ success: false, message: 'name パラメータが必要です' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        var member = getMemberByName(obsName);
        if (!member || !member.email) {
          return ContentService
            .createTextOutput(JSON.stringify({ success: false, message: 'メンバーが見つからないかメール未設定: ' + obsName }))
            .setMimeType(ContentService.MimeType.JSON);
        }
        var observerPageUrl = CONFIG.CONSENT.WEB_APP_URL + '?action=observer';
        var subject = '【リマインド】秘密保持誓約書（NDA）のご提出について';
        var body = obsName + ' 様\n\n' +
          'お疲れ様です。\n' +
          '関西学院大学 中小企業経営診断研究会無料経営相談分科会です。\n\n' +
          '経営相談にオブザーバーとしてご参加いただくにあたり、\n' +
          '秘密保持誓約書（NDA）のご提出をお願いしております。\n\n' +
          'まだご提出がお済みでない場合は、以下のオブザーバー専用ページから\n' +
          'NDAへの署名・提出をお願いいたします。\n\n' +
          '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
          '■ オブザーバー専用ページ\n' +
          '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
          observerPageUrl + '\n\n' +
          '上記ページでは以下の操作が可能です：\n' +
          '・相談予定の確認\n' +
          '・秘密保持誓約書（NDA）への署名・提出\n' +
          '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
          'ご不明な点がございましたら、お気軽にお問い合わせください。\n\n' +
          '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
          CONFIG.ORG.NAME + '\n' +
          'Email: ' + CONFIG.ORG.EMAIL + '\n' +
          '━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

        GmailApp.sendEmail(member.email, subject, body, {
          name: CONFIG.SENDER_NAME,
          replyTo: CONFIG.REPLY_TO
        });

        return ContentService
          .createTextOutput(JSON.stringify({ success: true, message: 'NDAリマインドメール送信完了', to: member.email, name: obsName }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // 記事管理（CMS）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // 記事管理ページ
    if (action === 'article-admin') {
      return generateArticleAdminPage();
    }

    // 記事一覧API（JSON）
    if (action === 'articles') {
      var artResult = getArticles({
        category: e.parameter.category || null,
        limit: e.parameter.limit || '20',
        offset: e.parameter.offset || '0'
      });
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, articles: artResult.articles, total: artResult.total }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 記事詳細API（JSON）
    if (action === 'article') {
      var artId = e.parameter.id;
      if (!artId) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'id パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var artDetail = getArticleById(artId);
      if (!artDetail) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '記事が見つかりません: ' + artId }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, article: artDetail }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 記事カテゴリ一覧API
    if (action === 'article-categories') {
      var cats = getArticleCategories();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, categories: cats }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 記事追加
    if (action === 'article-add') {
      addArticle({
        title: e.parameter.title,
        category: e.parameter.category,
        tags: e.parameter.tags,
        body: e.parameter.body,
        author: e.parameter.author,
        summary: e.parameter.summary,
        status: e.parameter.status,
        publishDate: e.parameter.publishDate
      });
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=article-admin";</script>'
      );
    }

    // 記事ステータス切替
    if (action === 'article-toggle') {
      toggleArticleStatus(e.parameter.row);
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=article-admin";</script>'
      );
    }

    // 記事削除
    if (action === 'article-delete') {
      deleteArticle(e.parameter.row);
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=article-admin";</script>'
      );
    }

    // 記事管理シートセットアップ（管理用）
    if (action === 'setup-articles') {
      var artSetupResult = setupArticlesSheet();
      return ContentService
        .createTextOutput(JSON.stringify(artSetupResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 日程修復（管理用）
    if (action === 'repair-schedule') {
      var repairYear = parseInt(e.parameter.year) || 2026;
      var repairMonth = parseInt(e.parameter.month) || 4;
      var repairResult = reprocessFormResponsesForMonth(repairYear, repairMonth);
      return ContentService
        .createTextOutput(JSON.stringify(repairResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3月18日対面枠追加（管理用）
    if (action === 'add-march18') {
      addMarch18InPerson();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, message: '3月18日対面枠追加完了' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // メンバー直接設定（管理用）
    if (action === 'set-slot-members') {
      var slotResult = setSlotMembers(e.parameter.date, e.parameter.time, e.parameter.members);
      return ContentService
        .createTextOutput(JSON.stringify(slotResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3/18メンバー設定（管理用）
    if (action === 'set-march18-members') {
      var m18Result = setMarch18Members();
      return ContentService
        .createTextOutput(JSON.stringify(m18Result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 記事管理シート削除（管理用）
    if (action === 'delete-articles-sheet') {
      deleteArticlesSheet();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, message: '記事管理シート削除完了' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // 会場管理
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // 会場管理ページ
    if (action === 'venue-admin') {
      return generateVenueAdminPage();
    }

    // 会場空き状況ページ
    if (action === 'venue-status') {
      return generateVenueStatusPage(e);
    }

    // 会場追加
    if (action === 'venue-add') {
      addVenue({
        name: e.parameter.name,
        address: e.parameter.address,
        capacity: e.parameter.capacity,
        equipment: e.parameter.equipment,
        price: e.parameter.price,
        notes: e.parameter.notes,
        hours: e.parameter.hours
      });
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=venue-admin";</script>'
      );
    }

    // 会場有効/無効切替
    if (action === 'venue-toggle') {
      toggleVenueActive(e.parameter.row);
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=venue-admin";</script>'
      );
    }

    // 会場削除
    if (action === 'venue-delete') {
      deleteVenue(e.parameter.row);
      return HtmlService.createHtmlOutput(
        '<script>window.location.href="' + CONFIG.CONSENT.WEB_APP_URL + '?action=venue-admin";</script>'
      );
    }

    // 会場マスタシートセットアップ（管理用）
    if (action === 'setup-venue') {
      var venueResult = setupVenueSheet();
      return ContentService
        .createTextOutput(JSON.stringify(venueResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 会場一覧API（JSON）
    if (action === 'venues') {
      var venueList = getVenues(e.parameter.all !== 'true');
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, venues: venueList }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // リーダー履歴の業種・テーマ一括補完（管理用）
    if (action === 'backfill-leader-history') {
      try {
        var bfResult = backfillLeaderHistoryFields();
        return ContentService
          .createTextOutput(JSON.stringify(bfResult))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 相談者同意処理のデバッグ（管理用）
    if (action === 'consent-debug') {
      try {
        var appId = e.parameter.appId;
        if (!appId) throw new Error('appIdパラメータが必要です');
        var rowIndex = findRowByApplicationId(appId);
        if (!rowIndex) throw new Error('申込ID ' + appId + ' が見つかりません');
        var debugData = getRowData(rowIndex);
        var isOnline = debugData.method && (debugData.method.indexOf('オンライン') >= 0 || debugData.method.indexOf('Zoom') >= 0);
        return ContentService
          .createTextOutput(JSON.stringify({
            success: true,
            rowIndex: rowIndex,
            id: debugData.id,
            name: debugData.name,
            email: debugData.email,
            status: debugData.status,
            ndaStatus: debugData.ndaStatus,
            method: debugData.method,
            isOnline: isOnline,
            confirmedDate: debugData.confirmedDate,
            zoomUrl: debugData.zoomUrl,
            leader: debugData.leader,
            staff: debugData.staff,
            industry: debugData.industry,
            theme: debugData.theme,
            fileId: debugData.fileId
          }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 手動でrunFirstPollingを実行（管理用）
    if (action === 'run-first-polling') {
      try {
        runFirstPolling();
        return ContentService
          .createTextOutput(JSON.stringify({ success: true, message: 'runFirstPolling を実行しました' }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // リーダー選定 & レポート配信
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // レポートアップロードページ
    if (action === 'report-upload') {
      return generateReportUploadPage(e);
    }

    // リーダー履歴シートセットアップ（管理用）
    if (action === 'setup-leader-history') {
      var leaderResult = setupLeaderHistorySheet();
      return ContentService
        .createTextOutput(JSON.stringify(leaderResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // レポート管理シートセットアップ（管理用）
    if (action === 'setup-report') {
      var reportResult = setupReportSheet();
      return ContentService
        .createTextOutput(JSON.stringify(reportResult))
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
    // 相談完了確認
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // 完了確認ページ表示
    if (action === 'completion') {
      return generateCompletionConfirmPage(e);
    }

    // 完了確認トリガーセットアップ（管理用）
    if (action === 'setup-completion-trigger') {
      var completionTriggerResult = setupCompletionTrigger();
      return ContentService
        .createTextOutput(JSON.stringify(completionTriggerResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // Zoom録画管理
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // 録画チェック手動実行（管理用）
    if (action === 'check-recordings') {
      try {
        var recResult = processZoomRecordings();
        return ContentService
          .createTextOutput(JSON.stringify(recResult))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // キャンセルメール検知テスト（管理用）
    if (action === 'check-cancel-emails') {
      try {
        checkCancellationEmails();
        return ContentService
          .createTextOutput(JSON.stringify({ success: true, message: 'キャンセルメール検知を実行しました' }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (e) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: e.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // キャンセルメール検知トリガーセットアップ（管理用）
    if (action === 'setup-cancel-trigger') {
      setupCancellationEmailTrigger();
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, message: 'キャンセルメール検知トリガー（10分おき）をセットアップしました' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 録画チェックトリガーセットアップ（管理用）
    if (action === 'setup-recording-trigger') {
      var recTriggerResult = setupRecordingCheckTrigger();
      return ContentService
        .createTextOutput(JSON.stringify(recTriggerResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // YouTube設定確認（管理用）
    if (action === 'youtube-status') {
      var ytProps = PropertiesService.getScriptProperties();
      var ytEnabled = !!(CONFIG.YOUTUBE && CONFIG.YOUTUBE.ENABLED);
      var ytCfUrl = ytProps.getProperty('YOUTUBE_CF_URL') || (CONFIG.YOUTUBE && CONFIG.YOUTUBE.CLOUD_FUNCTION_URL) || '';
      var ytCfSecret = ytProps.getProperty('YOUTUBE_CF_SECRET') ? true : !!(CONFIG.YOUTUBE && CONFIG.YOUTUBE.CLOUD_FUNCTION_SECRET);
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          enabled: ytEnabled,
          cloudFunctionUrl: ytCfUrl ? '設定済み' : '未設定',
          cloudFunctionSecret: ytCfSecret ? '設定済み' : '未設定'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // YouTube設定セットアップ（管理用）
    // ?action=setup-youtube&cfUrl=CLOUD_FUNCTION_URL&cfSecret=SHARED_SECRET
    if (action === 'setup-youtube') {
      try {
        var ytSetupProps = PropertiesService.getScriptProperties();
        if (e.parameter.cfUrl) {
          ytSetupProps.setProperty('YOUTUBE_CF_URL', e.parameter.cfUrl);
        }
        if (e.parameter.cfSecret) {
          ytSetupProps.setProperty('YOUTUBE_CF_SECRET', e.parameter.cfSecret);
        }
        return ContentService
          .createTextOutput(JSON.stringify({
            success: true,
            message: 'YouTube設定を保存しました。CONFIG.YOUTUBE.ENABLED = true に変更してデプロイしてください。',
            cfUrl: e.parameter.cfUrl ? '設定済み' : 'スキップ',
            cfSecret: e.parameter.cfSecret ? '設定済み' : 'スキップ'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // 文字起こし・報告書自動作成（Phase 3）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // 文字起こし手動実行（管理用）
    if (action === 'start-transcript') {
      var trRow = parseInt(e.parameter.row);
      if (!trRow || trRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var trResult = manualStartTranscription(trRow);
      return ContentService
        .createTextOutput(JSON.stringify(trResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 報告書自動生成手動実行（管理用）
    if (action === 'start-auto-report') {
      var arRow = parseInt(e.parameter.row);
      if (!arRow || arRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var arResult = manualStartReportGeneration(arRow);
      return ContentService
        .createTextOutput(JSON.stringify(arResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // パイプライン状態一覧（管理用）
    if (action === 'transcript-pipeline') {
      var pipeResult = getTranscriptPipelineStatus();
      return ContentService
        .createTextOutput(JSON.stringify(pipeResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 文字起こし・報告書設定確認（管理用）
    if (action === 'transcript-setup') {
      var setupResult = checkTranscriptSetup();
      return ContentService
        .createTextOutput(JSON.stringify(setupResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 文字起こし・報告書 Cloud Function 設定（管理用）
    if (action === 'setup-transcript') {
      try {
        var stResult = setupTranscriptCredentials(
          e.parameter.transcriptCfUrl || '',
          e.parameter.transcriptCfSecret || '',
          e.parameter.reportCfUrl || '',
          e.parameter.reportCfSecret || ''
        );
        return ContentService
          .createTextOutput(JSON.stringify(stResult))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Notion CF 設定（管理用）
    if (action === 'setup-notion-cf') {
      try {
        var props = PropertiesService.getScriptProperties();
        if (e.parameter.cfUrl) props.setProperty('NOTION_CF_URL', e.parameter.cfUrl);
        if (e.parameter.cfSecret) props.setProperty('NOTION_CF_SECRET', e.parameter.cfSecret);
        return ContentService
          .createTextOutput(JSON.stringify({
            success: true,
            message: 'Notion CF設定を保存しました',
            cfUrl: e.parameter.cfUrl ? '設定済み' : 'スキップ',
            cfSecret: e.parameter.cfSecret ? '設定済み' : 'スキップ'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // コンサルタント評価（Phase 4 v2）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // 評価実行（予約行から）
    if (action === 'run-evaluation') {
      var evalRow = parseInt(e.parameter.row);
      if (!evalRow || evalRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var evalResult = manualRunEvaluation(evalRow);
      return ContentService
        .createTextOutput(JSON.stringify(evalResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 評価結果一覧API
    if (action === 'evaluation-results') {
      var evalResults = getEvaluationResults({
        consultant: e.parameter.consultant || '',
        status: e.parameter.status || '',
        limit: e.parameter.limit || '20',
        offset: e.parameter.offset || '0'
      });
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, results: evalResults.results, total: evalResults.total }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 評価結果詳細API
    if (action === 'evaluation-detail') {
      var evalId = e.parameter.id;
      if (!evalId) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'id パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var evalDetail = getEvaluationById(evalId);
      if (!evalDetail) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '評価結果が見つかりません: ' + evalId }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, evaluation: evalDetail }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // コンサルタント別履歴API
    if (action === 'evaluation-history') {
      var histConsultant = e.parameter.consultant;
      if (!histConsultant) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'consultant パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var histResult = getConsultantHistory(histConsultant);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, consultant: histConsultant, history: histResult }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 統計サマリーAPI
    if (action === 'evaluation-stats') {
      var evalStats = getEvaluationStats();
      return ContentService
        .createTextOutput(JSON.stringify(evalStats))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 評価シートセットアップ（管理用）
    if (action === 'setup-evaluation') {
      var evalSetupResult = setupEvaluationSheet();
      return ContentService
        .createTextOutput(JSON.stringify(evalSetupResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 評価設定確認（管理用）
    if (action === 'evaluation-setup') {
      var evalSetup = checkEvaluationSetup();
      return ContentService
        .createTextOutput(JSON.stringify(evalSetup))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 評価CF設定（管理用）
    if (action === 'setup-evaluation-cf') {
      try {
        var evalProps = PropertiesService.getScriptProperties();
        if (e.parameter.cfUrl) evalProps.setProperty('EVALUATION_CF_URL', e.parameter.cfUrl);
        if (e.parameter.cfSecret) evalProps.setProperty('EVALUATION_CF_SECRET', e.parameter.cfSecret);
        return ContentService
          .createTextOutput(JSON.stringify({
            success: true,
            message: 'コンサルタント評価CF設定を保存しました',
            cfUrl: e.parameter.cfUrl ? '設定済み' : 'スキップ',
            cfSecret: e.parameter.cfSecret ? '設定済み' : 'スキップ'
          }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
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

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // スタッフポータル（Phase 5）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // ログインリクエスト（メールアドレスでマジックリンク送信）
    if (action === 'portal-login') {
      var loginResult = requestPortalLogin(e.parameter.email || '');
      return ContentService
        .createTextOutput(JSON.stringify(loginResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // マジックリンク検証 → セッション発行 → ポータルSPAにリダイレクト
    if (action === 'portal-verify') {
      var verifyResult = verifyPortalLogin(e.parameter.token || '');
      if (verifyResult.success) {
        // セッションIDを付与してGitHub PagesポータルSPAにリダイレクト
        var portalUrl = CONFIG.PORTAL.SITE_URL + '#/verified?sessionId=' + verifyResult.sessionId;
        return HtmlService.createHtmlOutput(
          '<html><head><meta charset="UTF-8"><meta http-equiv="refresh" content="0;url=' + portalUrl + '"><script>window.location.href="' + portalUrl + '";</script></head><body style="font-family:sans-serif;padding:3rem;text-align:center"><p>ポータルに移動中...</p><p><a href="' + portalUrl + '">こちらをクリック</a></p></body></html>'
        ).setTitle('ポータルに移動中');
      } else {
        // エラーページを表示
        var retryUrl = CONFIG.PORTAL.SITE_URL + '#/login';
        return HtmlService.createHtmlOutput(
          '<html><head><meta charset="UTF-8"><title>ログインエラー</title></head><body style="font-family:sans-serif;padding:3rem;text-align:center"><h2>' + verifyResult.message + '</h2><p style="margin-top:1rem"><a href="' + retryUrl + '">再度ログイン</a></p></body></html>'
        ).setTitle('ログインエラー');
      }
    }

    // セッション検証
    if (action === 'portal-session') {
      var sessionData = requireAuth(e);
      if (!sessionData) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'セッションが無効です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, session: sessionData }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ログアウト
    if (action === 'portal-logout') {
      destroySession(e.parameter.sessionId || e.parameter.session || '');
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, message: 'ログアウトしました' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ダッシュボード
    if (action === 'portal-dashboard') {
      var dashSession = requireAuth(e);
      if (!dashSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var dashData = getPortalDashboard(dashSession);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, data: dashData }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // シフト一覧
    if (action === 'portal-shifts') {
      var shiftSession = requireAuth(e);
      if (!shiftSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var shiftData = getMyShifts(shiftSession, e.parameter.month || '');
      return ContentService
        .createTextOutput(JSON.stringify(shiftData))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 案件一覧
    if (action === 'portal-cases') {
      var caseSession = requireAuth(e);
      if (!caseSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var caseData = getMyCases(caseSession);
      return ContentService
        .createTextOutput(JSON.stringify(caseData))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // シフト参加/取消
    if (action === 'portal-shift-toggle') {
      var shiftToggleSession = requireAuth(e);
      if (!shiftToggleSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var toggleResult = toggleShiftParticipation(shiftToggleSession, e.parameter.row, e.parameter.join === 'true');
      return ContentService
        .createTextOutput(JSON.stringify(toggleResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 会場設定+確定
    if (action === 'portal-set-venue') {
      var venueSession = requireAuth(e);
      if (!venueSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (!hasRole(venueSession.role, 'member')) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '権限がありません' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var venueResult = setVenueAndConfirm(venueSession, e.parameter);
      return ContentService
        .createTextOutput(JSON.stringify(venueResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // プロフィール変更依頼
    if (action === 'portal-profile-change') {
      var pcSession = requireAuth(e);
      if (!pcSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var pcResult = requestProfileChange(pcSession, e.parameter.detail || '');
      return ContentService
        .createTextOutput(JSON.stringify(pcResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // プロフィール取得
    if (action === 'portal-profile') {
      var profSession = requireAuth(e);
      if (!profSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var profData = getPortalProfile(profSession);
      return ContentService
        .createTextOutput(JSON.stringify(profData))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ニュース投稿（リーダー/管理者のみ）
    if (action === 'portal-news-post') {
      var newsSession = requireAuth(e);
      if (!newsSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (!hasRole(newsSession.role, 'member')) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '権限がありません' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var newsResult = postPortalNews(newsSession, e.parameter);
      return ContentService
        .createTextOutput(JSON.stringify(newsResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 全メンバー一覧（管理者のみ）
    if (action === 'portal-members') {
      var memSession = requireAuth(e);
      if (!memSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (!hasRole(memSession.role, 'admin')) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '管理者権限が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var memData = getPortalMembers();
      return ContentService
        .createTextOutput(JSON.stringify(memData))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ポータル: コンサルタント評価
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    // 評価一覧（認証必須）
    if (action === 'portal-evaluations') {
      var peSession = requireAuth(e);
      if (!peSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var peResults = getEvaluationResults({
        consultant: e.parameter.consultant || '',
        status: e.parameter.status || '',
        limit: e.parameter.limit || '20',
        offset: e.parameter.offset || '0'
      });
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, results: peResults.results, total: peResults.total }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 評価詳細（認証必須）
    if (action === 'portal-evaluation-detail') {
      var pdSession = requireAuth(e);
      if (!pdSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var pdId = e.parameter.id;
      if (!pdId) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'id パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var pdDetail = getEvaluationById(pdId);
      if (!pdDetail) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '評価結果が見つかりません' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, evaluation: pdDetail }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 人間採点登録（リーダー以上）
    if (action === 'portal-evaluation-human-score') {
      var phSession = requireAuth(e);
      if (!phSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (!hasRole(phSession.role, 'member')) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'メンバー以上の権限が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var hsResult = updateHumanScores(
        e.parameter.id || '',
        parseInt(e.parameter.h1) || 0,
        parseInt(e.parameter.h2) || 0,
        parseInt(e.parameter.h3) || 0,
        phSession.name
      );
      return ContentService
        .createTextOutput(JSON.stringify(hsResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 評価実行トリガー（管理者のみ）
    if (action === 'portal-run-evaluation') {
      var preSession = requireAuth(e);
      if (!preSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      if (!hasRole(preSession.role, 'admin')) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '管理者権限が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var preRow = parseInt(e.parameter.row);
      if (!preRow || preRow < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'row パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var preResult = manualRunEvaluation(preRow);
      return ContentService
        .createTextOutput(JSON.stringify(preResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 成長データ（認証必須）
    if (action === 'portal-evaluation-growth') {
      var pgSession = requireAuth(e);
      if (!pgSession) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: '認証が必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var pgConsultant = e.parameter.consultant;
      if (!pgConsultant) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, message: 'consultant パラメータが必要です' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      var pgHistory = getConsultantHistory(pgConsultant);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, consultant: pgConsultant, history: pgHistory }))
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
          'GET ?action=members': 'メンバー情報取得（LP用）',
          'GET ?action=news': 'お知らせ取得（LP用）',
          'GET ?action=news-admin': 'お知らせ管理ページ',
          'GET ?action=survey&token=xxx': 'アンケートページ',
          'GET ?action=observer': 'オブザーバー専用ページ',
          'GET ?action=report-upload&token=xxx': 'レポートアップロードページ',
          'GET ?action=completion&token=xxx': '相談完了確認ページ',
          'GET ?action=article-admin': '記事管理ページ',
          'GET ?action=articles': '記事一覧API（JSON）',
          'GET ?action=article&id=xxx': '記事詳細API（JSON）',
          'GET ?action=setup-articles': '記事管理シートセットアップ',
          'GET ?action=venue-admin': '会場管理ページ',
          'GET ?action=venue-status': '会場空き状況ページ',
          'GET ?action=venues': '会場一覧API（JSON）',
          'GET ?action=setup-venue': '会場マスタシートセットアップ',
          'GET ?action=setup-completion-trigger': '完了確認トリガーセットアップ',
          'GET ?action=check-recordings': '録画リンク手動チェック',
          'GET ?action=setup-recording-trigger': '録画チェックトリガーセットアップ',
          'GET ?action=youtube-status': 'YouTube設定確認',
          'GET ?action=setup-youtube&cfUrl=URL&cfSecret=SECRET': 'YouTube設定セットアップ',
          'GET ?action=setup-leader-history': 'リーダー履歴シートセットアップ',
          'GET ?action=setup-report': 'レポート管理シートセットアップ',
          'GET ?action=start-transcript&row=N': '文字起こし手動実行',
          'GET ?action=start-auto-report&row=N': '報告書自動生成手動実行',
          'GET ?action=transcript-pipeline': 'パイプライン状態一覧',
          'GET ?action=transcript-setup': '文字起こし設定確認',
          'GET ?action=setup-transcript': '文字起こしCF設定',
          'GET ?action=setup-notion-cf': 'Notion CF設定',
          'GET ?action=run-evaluation&row=N': '評価実行（予約データから）',
          'GET ?action=evaluation-results': '評価結果一覧',
          'GET ?action=evaluation-detail&id=xxx': '評価結果詳細',
          'GET ?action=evaluation-history&consultant=xxx': 'コンサルタント別履歴',
          'GET ?action=evaluation-stats': '評価統計',
          'GET ?action=setup-evaluation': '評価シートセットアップ',
          'GET ?action=evaluation-setup': '評価設定確認',
          'GET ?action=setup-evaluation-cf': '評価CF設定',
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

/**
 * 申込概要をDriveに自動保存
 * @param {Object} data - フォームデータ
 * @param {string} applicationId - 申込ID
 */
function saveApplicationSummaryToDrive(data, applicationId) {
  try {
    var folder = getDriveFolder('DRIVE_FOLDER_SUMMARY', '');
    var content = '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      '申込概要\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
      '申込ID：' + applicationId + '\n' +
      '申込日時：' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n\n' +
      '■ 申込者情報\n' +
      'お名前：' + (data.name || '') + '\n' +
      '企業名：' + (data.company || '') + '\n' +
      'メール：' + (data.email || '') + '\n' +
      '電話番号：' + (data.phone || '') + '\n' +
      '役職：' + (data.position || '') + '\n' +
      '業種：' + (data.industry || '') + '\n' +
      '企業URL：' + (data.companyUrl || '') + '\n\n' +
      '■ 相談内容\n' +
      '相談テーマ：' + (data.theme || '') + '\n' +
      '相談内容：' + (data.content || '') + '\n\n' +
      '■ 希望日程\n' +
      '希望日時1：' + (data.date1 || '') + '\n' +
      '希望日時2：' + (data.date2 || '') + '\n' +
      '相談方法：' + (data.method || '') + '\n' +
      '備考：' + (data.remarks || '') + '\n';

    var fileName = applicationId + '_申込概要.txt';
    folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
    console.log('申込概要保存: ' + fileName);
  } catch (e) {
    console.error('申込概要保存エラー:', e);
  }
}

