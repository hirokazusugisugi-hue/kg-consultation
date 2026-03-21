/**
 * Podcast管理
 * スプレッドシートの「Podcast」シートと連携
 * 管理用Webページ + メディアサイト向けAPI
 */

/**
 * Podcastシートの列定義
 */
var PODCAST_COLUMNS = {
  ID: 0,            // A: エピソードID
  TITLE: 1,         // B: タイトル
  DESCRIPTION: 2,   // C: 説明
  PUBLISH_DATE: 3,  // D: 公開日
  STATUS: 4,        // E: 状態（draft/published）
  SPOTIFY_URL: 5,   // F: Spotify URL
  APPLE_URL: 6,     // G: Apple Podcast URL
  YOUTUBE_URL: 7,   // H: YouTube URL
  THUMBNAIL: 8,     // I: サムネイルURL
  RELATED_ARTICLE: 9, // J: 関連記事ID
  CREATED_AT: 10    // K: 作成日時
};

var PODCAST_SHEET_NAME = 'Podcast';

/**
 * Podcastシートのセットアップ
 */
function setupPodcastSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PODCAST_SHEET_NAME);
  }

  var headers = ['エピソードID', 'タイトル', '説明', '公開日', '状態', 'Spotify URL', 'Apple Podcast URL', 'YouTube URL', 'サムネイルURL', '関連記事ID', '作成日時'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 100);   // エピソードID
  sheet.setColumnWidth(2, 300);   // タイトル
  sheet.setColumnWidth(3, 400);   // 説明
  sheet.setColumnWidth(4, 120);   // 公開日
  sheet.setColumnWidth(5, 80);    // 状態
  sheet.setColumnWidth(6, 250);   // Spotify
  sheet.setColumnWidth(7, 250);   // Apple
  sheet.setColumnWidth(8, 250);   // YouTube
  sheet.setColumnWidth(9, 250);   // サムネイル
  sheet.setColumnWidth(10, 100);  // 関連記事ID
  sheet.setColumnWidth(11, 160);  // 作成日時

  // 状態プルダウン
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['draft', 'published'])
    .build();
  sheet.getRange(2, PODCAST_COLUMNS.STATUS + 1, 500, 1).setDataValidation(statusRule);

  // 条件付き書式
  var statusRange = sheet.getRange('E2:E500');
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('published')
    .setBackground('#d4edda')
    .setRanges([statusRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('draft')
    .setBackground('#fff3cd')
    .setRanges([statusRange])
    .build());
  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenRows(1);

  // サンプルデータ
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var sampleData = [
    ['P001', '第1回：経営課題と診断士の役割', '中小企業が直面する経営課題と、中小企業診断士がどのように支援できるかを解説します。', '2026.04.01', 'published', '', '', '', '', 'A001', now],
    ['P002', '第2回：資金繰り改善の5つのポイント', '中小企業経営者が押さえておくべき資金管理の基本と改善策をまとめました。', '2026.04.15', 'published', '', '', '', '', 'A002', now]
  ];
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  console.log('Podcastシートのセットアップが完了しました');
  return { success: true, message: 'Podcastシートのセットアップが完了しました' };
}

/**
 * エピソードID生成
 */
function generatePodcastId() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return 'P001';

  var data = sheet.getDataRange().getValues();
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][PODCAST_COLUMNS.ID]);
    var match = id.match(/^P(\d+)$/);
    if (match) {
      var num = parseInt(match[1]);
      if (num > maxNum) maxNum = num;
    }
  }
  return 'P' + String(maxNum + 1).padStart(3, '0');
}

/**
 * Podcast一覧取得（API用）
 * @param {Object} options - { limit, offset }
 * @returns {Object}
 */
function getPodcasts(options) {
  options = options || {};
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { podcasts: [], total: 0 };

  var data = sheet.getDataRange().getValues();
  var podcasts = [];

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][PODCAST_COLUMNS.STATUS]);
    if (status !== 'published') continue;

    podcasts.push({
      id: String(data[i][PODCAST_COLUMNS.ID]),
      title: String(data[i][PODCAST_COLUMNS.TITLE]),
      description: String(data[i][PODCAST_COLUMNS.DESCRIPTION]),
      publishDate: String(data[i][PODCAST_COLUMNS.PUBLISH_DATE]),
      spotifyUrl: String(data[i][PODCAST_COLUMNS.SPOTIFY_URL]),
      appleUrl: String(data[i][PODCAST_COLUMNS.APPLE_URL]),
      youtubeUrl: String(data[i][PODCAST_COLUMNS.YOUTUBE_URL]),
      thumbnail: String(data[i][PODCAST_COLUMNS.THUMBNAIL]),
      relatedArticleId: String(data[i][PODCAST_COLUMNS.RELATED_ARTICLE]),
      row: i + 1
    });
  }

  // 公開日降順
  podcasts.sort(function(a, b) {
    return b.publishDate.localeCompare(a.publishDate);
  });

  var offset = parseInt(options.offset) || 0;
  var limit = parseInt(options.limit) || 20;
  var total = podcasts.length;
  podcasts = podcasts.slice(offset, offset + limit);

  return { podcasts: podcasts, total: total };
}

/**
 * Podcast詳細取得
 * @param {string} podcastId
 * @returns {Object|null}
 */
function getPodcastById(podcastId) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][PODCAST_COLUMNS.ID]) === podcastId) {
      return {
        id: String(data[i][PODCAST_COLUMNS.ID]),
        title: String(data[i][PODCAST_COLUMNS.TITLE]),
        description: String(data[i][PODCAST_COLUMNS.DESCRIPTION]),
        publishDate: String(data[i][PODCAST_COLUMNS.PUBLISH_DATE]),
        status: String(data[i][PODCAST_COLUMNS.STATUS]),
        spotifyUrl: String(data[i][PODCAST_COLUMNS.SPOTIFY_URL]),
        appleUrl: String(data[i][PODCAST_COLUMNS.APPLE_URL]),
        youtubeUrl: String(data[i][PODCAST_COLUMNS.YOUTUBE_URL]),
        thumbnail: String(data[i][PODCAST_COLUMNS.THUMBNAIL]),
        relatedArticleId: String(data[i][PODCAST_COLUMNS.RELATED_ARTICLE]),
        row: i + 1
      };
    }
  }
  return null;
}

/**
 * Podcastを追加
 * @param {Object} params
 * @returns {Object}
 */
function addPodcast(params) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  if (!sheet) {
    setupPodcastSheet();
    sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  }

  var podcastId = generatePodcastId();
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  var newRow = [
    podcastId,
    params.title || '',
    params.description || '',
    params.publishDate || '',
    params.status || 'draft',
    params.spotifyUrl || '',
    params.appleUrl || '',
    params.youtubeUrl || '',
    params.thumbnail || '',
    params.relatedArticleId || '',
    now
  ];

  sheet.appendRow(newRow);
  console.log('Podcast追加: ' + podcastId + ' - ' + params.title);
  return { success: true, podcastId: podcastId, message: 'Podcastを追加しました: ' + params.title };
}

/**
 * Podcastを更新
 * @param {string} podcastId - 更新対象のエピソードID
 * @param {Object} params - 更新フィールド
 * @returns {Object}
 */
function updatePodcast(podcastId, params) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'Podcastシートが見つかりません' };
  }

  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][PODCAST_COLUMNS.ID]) === podcastId) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    return { success: false, message: 'エピソードが見つかりません: ' + podcastId };
  }

  // 指定されたフィールドのみ更新
  var fieldMap = {
    title: PODCAST_COLUMNS.TITLE,
    description: PODCAST_COLUMNS.DESCRIPTION,
    publishDate: PODCAST_COLUMNS.PUBLISH_DATE,
    status: PODCAST_COLUMNS.STATUS,
    spotifyUrl: PODCAST_COLUMNS.SPOTIFY_URL,
    appleUrl: PODCAST_COLUMNS.APPLE_URL,
    youtubeUrl: PODCAST_COLUMNS.YOUTUBE_URL,
    thumbnail: PODCAST_COLUMNS.THUMBNAIL,
    relatedArticleId: PODCAST_COLUMNS.RELATED_ARTICLE
  };

  for (var key in params) {
    if (params.hasOwnProperty(key) && fieldMap.hasOwnProperty(key) && params[key] !== undefined && params[key] !== null) {
      sheet.getRange(targetRow, fieldMap[key] + 1).setValue(params[key]);
    }
  }

  console.log('Podcast更新: ' + podcastId);
  return { success: true, podcastId: podcastId, message: 'Podcastを更新しました: ' + podcastId };
}

/**
 * Podcastのステータスを切り替え
 * @param {number} row
 */
function togglePodcastStatus(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  if (!sheet) return;

  var rowNum = parseInt(row);
  var current = sheet.getRange(rowNum, PODCAST_COLUMNS.STATUS + 1).getValue();
  var newStatus = current === 'published' ? 'draft' : 'published';
  sheet.getRange(rowNum, PODCAST_COLUMNS.STATUS + 1).setValue(newStatus);

  if (newStatus === 'published') {
    var pubDate = sheet.getRange(rowNum, PODCAST_COLUMNS.PUBLISH_DATE + 1).getValue();
    if (!pubDate) {
      var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy.MM.dd');
      sheet.getRange(rowNum, PODCAST_COLUMNS.PUBLISH_DATE + 1).setValue(today);
    }
  }

  console.log('Podcastステータス切替: 行' + rowNum + ' → ' + newStatus);
}

/**
 * Podcastを削除
 * @param {number} row
 */
function deletePodcast(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  if (!sheet) return;
  sheet.deleteRow(parseInt(row));
  console.log('Podcast削除: 行' + row);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 管理画面
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * Podcast管理ページを生成
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function generatePodcastAdminPage() {
  var result = getPodcasts({ limit: '100' });
  // 下書きも含めて取得
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(PODCAST_SHEET_NAME);
  var allPodcasts = [];
  if (sheet && sheet.getLastRow() >= 2) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      allPodcasts.push({
        id: String(data[i][PODCAST_COLUMNS.ID]),
        title: String(data[i][PODCAST_COLUMNS.TITLE]),
        description: String(data[i][PODCAST_COLUMNS.DESCRIPTION]),
        publishDate: String(data[i][PODCAST_COLUMNS.PUBLISH_DATE]),
        status: String(data[i][PODCAST_COLUMNS.STATUS]),
        spotifyUrl: String(data[i][PODCAST_COLUMNS.SPOTIFY_URL]),
        appleUrl: String(data[i][PODCAST_COLUMNS.APPLE_URL]),
        youtubeUrl: String(data[i][PODCAST_COLUMNS.YOUTUBE_URL]),
        row: i + 1
      });
    }
  }
  var baseUrl = CONFIG.CONSENT.WEB_APP_URL;

  var html = '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Podcast管理</title>' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif; background: #f5f5f5; color: #333; padding: 16px; }' +
    'h1 { background: #0f2350; color: #fff; padding: 16px; border-radius: 8px 8px 0 0; font-size: 18px; }' +
    '.container { max-width: 800px; margin: 0 auto; }' +
    '.section { background: #fff; border-radius: 8px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden; }' +
    '.section-title { background: #e8eaf0; padding: 12px 16px; font-weight: bold; font-size: 14px; }' +
    '.section-body { padding: 16px; }' +
    'label { display: block; font-size: 13px; font-weight: bold; margin-bottom: 4px; margin-top: 10px; }' +
    'label:first-child { margin-top: 0; }' +
    'input[type="text"], textarea { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; font-family: inherit; }' +
    'textarea { min-height: 80px; resize: vertical; }' +
    '.btn { display: inline-block; padding: 8px 20px; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; text-decoration: none; margin-top: 12px; }' +
    '.btn-primary { background: #0f2350; color: #fff; }' +
    '.btn-success { background: #28a745; color: #fff; }' +
    '.btn-warning { background: #ffc107; color: #333; }' +
    '.btn-danger { background: #dc3545; color: #fff; }' +
    '.btn-sm { font-size: 12px; padding: 4px 12px; margin-top: 0; }' +
    '.item { border-bottom: 1px solid #eee; padding: 12px 16px; }' +
    '.item:last-child { border-bottom: none; }' +
    '.item-title { font-size: 15px; font-weight: bold; margin-bottom: 4px; }' +
    '.item-meta { font-size: 12px; color: #666; display: flex; gap: 12px; margin-bottom: 8px; }' +
    '.item-actions { display: flex; gap: 8px; flex-wrap: wrap; }' +
    '.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: bold; }' +
    '.badge-published { background: #d4edda; color: #155724; }' +
    '.badge-draft { background: #fff3cd; color: #856404; }' +
    '.form-row { display: flex; gap: 12px; }' +
    '.form-row > div { flex: 1; }' +
    '</style>' +
    '</head><body>' +
    '<div class="container">' +
    '<h1>Podcast管理</h1>';

  // 新規追加フォーム
  html += '<div class="section">' +
    '<div class="section-title">新規エピソードを追加</div>' +
    '<div class="section-body">' +
    '<form action="' + baseUrl + '" method="get">' +
    '<input type="hidden" name="action" value="podcast-add">' +
    '<label>タイトル *</label>' +
    '<input type="text" name="title" required placeholder="第○回：エピソードタイトル">' +
    '<label>説明</label>' +
    '<textarea name="description" placeholder="エピソードの説明"></textarea>' +
    '<div class="form-row">' +
    '<div><label>公開日</label><input type="text" name="publishDate" placeholder="2026.04.01"></div>' +
    '<div><label>状態</label><input type="text" name="status" value="draft" placeholder="draft / published"></div>' +
    '</div>' +
    '<label>Spotify URL</label>' +
    '<input type="text" name="spotifyUrl" placeholder="https://open.spotify.com/episode/...">' +
    '<label>Apple Podcast URL</label>' +
    '<input type="text" name="appleUrl" placeholder="https://podcasts.apple.com/...">' +
    '<label>YouTube URL</label>' +
    '<input type="text" name="youtubeUrl" placeholder="https://www.youtube.com/watch?v=...">' +
    '<div class="form-row">' +
    '<div><label>サムネイルURL</label><input type="text" name="thumbnail" placeholder="https://..."></div>' +
    '<div><label>関連記事ID</label><input type="text" name="relatedArticleId" placeholder="A001"></div>' +
    '</div>' +
    '<button type="submit" class="btn btn-primary">追加する</button>' +
    '</form>' +
    '</div></div>';

  // エピソード一覧
  html += '<div class="section">' +
    '<div class="section-title">エピソード一覧 (' + allPodcasts.length + '件)</div>';

  if (allPodcasts.length === 0) {
    html += '<div class="section-body"><p>エピソードがありません</p></div>';
  } else {
    allPodcasts.forEach(function(p) {
      var badgeClass = p.status === 'published' ? 'badge-published' : 'badge-draft';
      var badgeLabel = p.status === 'published' ? '公開中' : '下書き';

      html += '<div class="item">' +
        '<div class="item-title">' + escapeHtml(p.title) + '</div>' +
        '<div class="item-meta">' +
        '<span class="badge ' + badgeClass + '">' + badgeLabel + '</span>' +
        '<span>' + (p.publishDate || '未設定') + '</span>' +
        '<span>' + escapeHtml(p.id) + '</span>' +
        '</div>' +
        '<div class="item-actions">';

      if (p.status === 'published') {
        html += '<a href="' + baseUrl + '?action=podcast-toggle&row=' + p.row + '" class="btn btn-warning btn-sm">非公開にする</a>';
      } else {
        html += '<a href="' + baseUrl + '?action=podcast-toggle&row=' + p.row + '" class="btn btn-success btn-sm">公開する</a>';
      }
      html += ' <a href="' + baseUrl + '?action=podcast-delete&row=' + p.row + '" class="btn btn-danger btn-sm" onclick="return confirm(\'本当に削除しますか？\')">削除</a>';
      html += '</div></div>';
    });
  }

  html += '</div></div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('Podcast管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
