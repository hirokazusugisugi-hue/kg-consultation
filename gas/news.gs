/**
 * お知らせ管理
 * スプレッドシートの「お知らせ」シートと連携
 * 管理用Webページ + LP向けAPI
 */

/**
 * お知らせシートのセットアップ
 */
function setupNewsSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var sheet = ss.getSheetByName(CONFIG.NEWS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NEWS_SHEET_NAME);
  }

  // ヘッダー設定
  var headers = ['日付', '内容', '表示フラグ', '作成日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 120);  // 日付
  sheet.setColumnWidth(2, 400);  // 内容
  sheet.setColumnWidth(3, 80);   // 表示フラグ
  sheet.setColumnWidth(4, 160);  // 作成日時

  // 1行目を固定
  sheet.setFrozenRows(1);

  // 初期データを投入
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var sampleData = [
    ['2026.04.01', '無料経営相談の本格運用を開始しました。', 'TRUE', now]
  ];
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  console.log('お知らせシートのセットアップが完了しました');
}

/**
 * お知らせを追加
 * @param {string} date - 日付（例: 2026.05.01）
 * @param {string} content - 内容
 */
function addNews(date, content) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.NEWS_SHEET_NAME);

  if (!sheet) {
    setupNewsSheet();
    sheet = ss.getSheetByName(CONFIG.NEWS_SHEET_NAME);
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var lastRow = sheet.getLastRow();

  sheet.getRange(lastRow + 1, 1, 1, 4).setValues([
    [date, content, 'TRUE', now]
  ]);

  console.log('お知らせ追加: ' + date + ' - ' + content);
}

/**
 * 最新のお知らせを取得（LP用API）
 * 表示フラグ = TRUE のみ、日付降順、最大10件
 * @returns {Array} お知らせデータの配列
 */
function getLatestNews() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.NEWS_SHEET_NAME);

  if (!sheet) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  var news = [];

  for (var i = 1; i < data.length; i++) {
    var visible = data[i][NEWS_COLUMNS.VISIBLE];
    if (visible === true || visible === 'TRUE' || visible === 'true') {
      news.push({
        date: String(data[i][NEWS_COLUMNS.DATE]),
        content: String(data[i][NEWS_COLUMNS.CONTENT])
      });
    }
  }

  // 日付降順でソート
  news.sort(function(a, b) {
    return b.date.localeCompare(a.date);
  });

  // 最大10件
  return news.slice(0, 10);
}

/**
 * お知らせの表示/非表示を切り替え
 * @param {number} row - シートの行番号（2以上）
 */
function toggleNewsVisibility(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.NEWS_SHEET_NAME);
  if (!sheet) return;

  var rowNum = parseInt(row);
  var current = sheet.getRange(rowNum, NEWS_COLUMNS.VISIBLE + 1).getValue();
  var newValue = (current === true || current === 'TRUE' || current === 'true') ? 'FALSE' : 'TRUE';
  sheet.getRange(rowNum, NEWS_COLUMNS.VISIBLE + 1).setValue(newValue);

  console.log('お知らせ表示切替: 行' + rowNum + ' → ' + newValue);
}

/**
 * お知らせを削除
 * @param {number} row - シートの行番号（2以上）
 */
function deleteNews(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.NEWS_SHEET_NAME);
  if (!sheet) return;

  var rowNum = parseInt(row);
  sheet.deleteRow(rowNum);

  console.log('お知らせ削除: 行' + rowNum);
}

/**
 * 管理用Webページを生成
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function generateNewsAdminPage() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.NEWS_SHEET_NAME);

  var newsItems = [];
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      newsItems.push({
        row: i + 1,
        date: String(data[i][NEWS_COLUMNS.DATE]),
        content: String(data[i][NEWS_COLUMNS.CONTENT]),
        visible: data[i][NEWS_COLUMNS.VISIBLE] === true ||
                 data[i][NEWS_COLUMNS.VISIBLE] === 'TRUE' ||
                 data[i][NEWS_COLUMNS.VISIBLE] === 'true',
        createdAt: String(data[i][NEWS_COLUMNS.CREATED_AT])
      });
    }
  }

  // 日付降順ソート
  newsItems.sort(function(a, b) {
    return b.date.localeCompare(a.date);
  });

  // 最新の表示中ニュースを取得（プレビュー用）
  var latestVisible = newsItems.find(function(item) { return item.visible; });

  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy.MM.dd');
  var baseUrl = CONFIG.CONSENT.WEB_APP_URL;

  var html = '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>お知らせ管理</title>' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif; background: #f5f5f5; color: #333; padding: 16px; }' +
    'h1 { background: #0f2350; color: #fff; padding: 16px; border-radius: 8px 8px 0 0; font-size: 18px; }' +
    '.container { max-width: 640px; margin: 0 auto; }' +
    '.section { background: #fff; border-radius: 8px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden; }' +
    '.section-title { background: #e8eaf0; padding: 12px 16px; font-weight: bold; font-size: 14px; }' +
    '.section-body { padding: 16px; }' +
    'label { display: block; font-size: 13px; font-weight: bold; margin-bottom: 4px; }' +
    'input[type="text"], input[type="date"] { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; margin-bottom: 12px; }' +
    '.btn { display: inline-block; padding: 8px 20px; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; text-decoration: none; }' +
    '.btn-primary { background: #0f2350; color: #fff; }' +
    '.btn-secondary { background: #6c757d; color: #fff; }' +
    '.btn-danger { background: #dc3545; color: #fff; font-size: 12px; padding: 4px 12px; }' +
    '.btn-sm { font-size: 12px; padding: 4px 12px; }' +
    '.news-item { border-bottom: 1px solid #eee; padding: 12px 16px; }' +
    '.news-item:last-child { border-bottom: none; }' +
    '.news-date { font-size: 12px; color: #666; margin-bottom: 4px; }' +
    '.news-content { font-size: 14px; margin-bottom: 8px; }' +
    '.news-actions { display: flex; gap: 8px; align-items: center; }' +
    '.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: bold; }' +
    '.badge-visible { background: #d4edda; color: #155724; }' +
    '.badge-hidden { background: #f8d7da; color: #721c24; }' +
    '.preview-box { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; padding: 12px; font-size: 14px; }' +
    '.preview-label { font-size: 12px; color: #666; margin-top: 8px; }' +
    '</style>' +
    '</head><body>' +
    '<div class="container">' +
    '<h1>お知らせ管理</h1>';

  // 新規投稿フォーム
  html += '<div class="section">' +
    '<div class="section-title">新規投稿</div>' +
    '<div class="section-body">' +
    '<form action="' + baseUrl + '" method="get">' +
    '<input type="hidden" name="action" value="news-add">' +
    '<label>日付</label>' +
    '<input type="text" name="date" value="' + today + '" placeholder="2026.05.01">' +
    '<label>内容</label>' +
    '<input type="text" name="content" placeholder="ニュースの内容を入力">' +
    '<button type="submit" class="btn btn-primary">投稿する</button>' +
    '</form>' +
    '</div></div>';

  // 現在のお知らせ一覧
  html += '<div class="section">' +
    '<div class="section-title">現在のお知らせ一覧</div>';

  if (newsItems.length === 0) {
    html += '<div class="section-body"><p>お知らせはありません</p></div>';
  } else {
    newsItems.forEach(function(item) {
      html += '<div class="news-item">' +
        '<div class="news-date">' + item.date + '</div>' +
        '<div class="news-content">' + escapeHtml(item.content) + '</div>' +
        '<div class="news-actions">';

      if (item.visible) {
        html += '<span class="badge badge-visible">表示中</span> ' +
          '<a href="' + baseUrl + '?action=news-toggle&row=' + item.row + '" class="btn btn-secondary btn-sm">非表示にする</a>';
      } else {
        html += '<span class="badge badge-hidden">非表示</span> ' +
          '<a href="' + baseUrl + '?action=news-toggle&row=' + item.row + '" class="btn btn-secondary btn-sm">表示にする</a> ' +
          '<a href="' + baseUrl + '?action=news-delete&row=' + item.row + '" class="btn btn-danger btn-sm" onclick="return confirm(\'本当に削除しますか？\')">削除</a>';
      }

      html += '</div></div>';
    });
  }

  html += '</div>';

  // プレビュー
  html += '<div class="section">' +
    '<div class="section-title">プレビュー</div>' +
    '<div class="section-body">' +
    '<div class="preview-box">';

  if (latestVisible) {
    html += 'News ' + latestVisible.date + ' — ' + escapeHtml(latestVisible.content);
  } else {
    html += '（表示中のお知らせはありません）';
  }

  html += '</div>' +
    '<div class="preview-label">※ LP上の表示イメージ</div>' +
    '</div></div>';

  html += '</div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('お知らせ管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLエスケープ
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
