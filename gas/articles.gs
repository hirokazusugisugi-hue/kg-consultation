/**
 * 記事管理（簡易CMS）
 * スプレッドシートの「記事管理」シートと連携
 * 管理用Webページ + PRサイト向けAPI
 */

/**
 * 記事管理シートの列定義
 */
var ARTICLE_COLUMNS = {
  ID: 0,            // A: 記事ID
  TITLE: 1,         // B: タイトル
  CATEGORY: 2,      // C: カテゴリ
  TAGS: 3,          // D: タグ（カンマ区切り）
  BODY: 4,          // E: 本文（Markdown）
  AUTHOR: 5,        // F: 著者
  THUMBNAIL: 6,     // G: サムネイルURL
  STATUS: 7,        // H: 状態（draft/published）
  PUBLISH_DATE: 8,  // I: 公開日
  CREATED_AT: 9,    // J: 作成日
  AUDIO_FILE_ID: 10,// K: 音声ファイルID（Podcast連携用）
  SUMMARY: 11       // L: 要約
};

var ARTICLE_SHEET_NAME = '記事管理';

var ARTICLE_CATEGORIES = [
  '経営戦略',
  '財務・会計',
  'マーケティング',
  '組織・人材',
  'DX・IT',
  '事例紹介',
  'コラム',
  'その他'
];

/**
 * 記事管理シートのセットアップ
 */
function setupArticlesSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var sheet = ss.getSheetByName(ARTICLE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ARTICLE_SHEET_NAME);
  }

  var headers = ['記事ID', 'タイトル', 'カテゴリ', 'タグ', '本文(Markdown)', '著者', 'サムネイルURL', '状態', '公開日', '作成日', '音声ファイルID', '要約'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 80);    // 記事ID
  sheet.setColumnWidth(2, 250);   // タイトル
  sheet.setColumnWidth(3, 100);   // カテゴリ
  sheet.setColumnWidth(4, 150);   // タグ
  sheet.setColumnWidth(5, 400);   // 本文
  sheet.setColumnWidth(6, 100);   // 著者
  sheet.setColumnWidth(7, 200);   // サムネイル
  sheet.setColumnWidth(8, 80);    // 状態
  sheet.setColumnWidth(9, 120);   // 公開日
  sheet.setColumnWidth(10, 120);  // 作成日
  sheet.setColumnWidth(11, 200);  // 音声ファイルID
  sheet.setColumnWidth(12, 300);  // 要約

  // 状態プルダウン
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['draft', 'published'])
    .build();
  sheet.getRange(2, ARTICLE_COLUMNS.STATUS + 1, 500, 1).setDataValidation(statusRule);

  // カテゴリプルダウン
  var catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ARTICLE_CATEGORIES)
    .build();
  sheet.getRange(2, ARTICLE_COLUMNS.CATEGORY + 1, 500, 1).setDataValidation(catRule);

  // 条件付き書式
  var statusRange = sheet.getRange('H2:H500');
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

  // サンプル記事
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  var sampleData = [
    ['A001', '中小企業の経営課題と診断士の役割', '経営戦略', '経営戦略,中小企業診断士',
     '## はじめに\n\n中小企業を取り巻く経営環境は年々変化しています。\n\n## 主な経営課題\n\n1. **人材不足** — 採用難と離職率の課題\n2. **デジタル化** — DX推進の遅れ\n3. **事業承継** — 後継者問題\n\n## 診断士の役割\n\n中小企業診断士は、これらの課題に対し、客観的な分析と実践的な提言を行います。\n\n> 「経営者に寄り添い、共に解決策を見出す」\n\nそれが私たちの使命です。',
     '中小企業経営相談研究会', '', 'published', '2026.04.01', now, '', '中小企業が直面する経営課題と、中小企業診断士がどのように支援できるかを解説します。'],
    ['A002', '資金繰り改善のための5つのポイント', '財務・会計', '資金繰り,財務,キャッシュフロー',
     '## 資金繰りの重要性\n\n企業の存続に最も直結するのが資金繰りです。\n\n## 改善ポイント\n\n### 1. 売掛金回収の迅速化\n回収サイトの短縮、早期回収インセンティブの導入\n\n### 2. 在庫管理の適正化\n過剰在庫の削減、発注ロットの見直し\n\n### 3. 支払条件の交渉\n仕入先との条件交渉、支払いサイトの調整\n\n### 4. 融資・補助金の活用\n政府系金融機関の低利融資、各種補助金の活用\n\n### 5. 月次資金繰り表の作成\n3ヶ月先までの資金予測を可視化',
     '中小企業経営相談研究会', '', 'published', '2026.04.01', now, '', '中小企業経営者が押さえておくべき資金管理の基本と改善策をまとめました。']
  ];
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  console.log('記事管理シートのセットアップが完了しました');
  return { success: true, message: '記事管理シートのセットアップが完了しました' };
}

/**
 * 記事ID生成
 */
function generateArticleId() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ARTICLE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return 'A001';

  var data = sheet.getDataRange().getValues();
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][ARTICLE_COLUMNS.ID]);
    var match = id.match(/^A(\d+)$/);
    if (match) {
      var num = parseInt(match[1]);
      if (num > maxNum) maxNum = num;
    }
  }
  return 'A' + String(maxNum + 1).padStart(3, '0');
}

/**
 * 記事一覧取得（API用）
 * @param {Object} options - { category, limit, offset, status }
 * @returns {Array}
 */
function getArticles(options) {
  options = options || {};
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ARTICLE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var articles = [];

  for (var i = 1; i < data.length; i++) {
    var status = String(data[i][ARTICLE_COLUMNS.STATUS]);
    // デフォルトは published のみ
    if (!options.includeAll && status !== 'published') continue;

    var category = String(data[i][ARTICLE_COLUMNS.CATEGORY]);
    if (options.category && category !== options.category) continue;

    articles.push({
      id: String(data[i][ARTICLE_COLUMNS.ID]),
      title: String(data[i][ARTICLE_COLUMNS.TITLE]),
      category: category,
      tags: String(data[i][ARTICLE_COLUMNS.TAGS]),
      author: String(data[i][ARTICLE_COLUMNS.AUTHOR]),
      thumbnail: String(data[i][ARTICLE_COLUMNS.THUMBNAIL]),
      status: status,
      publishDate: String(data[i][ARTICLE_COLUMNS.PUBLISH_DATE]),
      summary: String(data[i][ARTICLE_COLUMNS.SUMMARY]),
      row: i + 1
    });
  }

  // 公開日降順
  articles.sort(function(a, b) {
    return b.publishDate.localeCompare(a.publishDate);
  });

  // ページネーション
  var offset = parseInt(options.offset) || 0;
  var limit = parseInt(options.limit) || 20;
  var total = articles.length;
  articles = articles.slice(offset, offset + limit);

  return { articles: articles, total: total };
}

/**
 * 記事詳細取得（本文含む）
 * @param {string} articleId - 記事ID
 * @returns {Object|null}
 */
function getArticleById(articleId) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ARTICLE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][ARTICLE_COLUMNS.ID]) === articleId) {
      return {
        id: String(data[i][ARTICLE_COLUMNS.ID]),
        title: String(data[i][ARTICLE_COLUMNS.TITLE]),
        category: String(data[i][ARTICLE_COLUMNS.CATEGORY]),
        tags: String(data[i][ARTICLE_COLUMNS.TAGS]),
        body: String(data[i][ARTICLE_COLUMNS.BODY]),
        author: String(data[i][ARTICLE_COLUMNS.AUTHOR]),
        thumbnail: String(data[i][ARTICLE_COLUMNS.THUMBNAIL]),
        status: String(data[i][ARTICLE_COLUMNS.STATUS]),
        publishDate: String(data[i][ARTICLE_COLUMNS.PUBLISH_DATE]),
        createdAt: String(data[i][ARTICLE_COLUMNS.CREATED_AT]),
        summary: String(data[i][ARTICLE_COLUMNS.SUMMARY]),
        row: i + 1
      };
    }
  }
  return null;
}

/**
 * 記事を追加
 * @param {Object} params
 * @returns {Object}
 */
function addArticle(params) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ARTICLE_SHEET_NAME);
  if (!sheet) {
    setupArticlesSheet();
    sheet = ss.getSheetByName(ARTICLE_SHEET_NAME);
  }

  var articleId = generateArticleId();
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  var newRow = [
    articleId,
    params.title || '',
    params.category || 'コラム',
    params.tags || '',
    params.body || '',
    params.author || '',
    params.thumbnail || '',
    params.status || 'draft',
    params.publishDate || '',
    now,
    params.audioFileId || '',
    params.summary || ''
  ];

  sheet.appendRow(newRow);
  console.log('記事追加: ' + articleId + ' - ' + params.title);
  return { success: true, articleId: articleId, message: '記事を追加しました: ' + params.title };
}

/**
 * 記事のステータスを切り替え（draft ↔ published）
 * @param {number} row
 */
function toggleArticleStatus(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ARTICLE_SHEET_NAME);
  if (!sheet) return;

  var rowNum = parseInt(row);
  var current = sheet.getRange(rowNum, ARTICLE_COLUMNS.STATUS + 1).getValue();
  var newStatus = current === 'published' ? 'draft' : 'published';
  sheet.getRange(rowNum, ARTICLE_COLUMNS.STATUS + 1).setValue(newStatus);

  // published に切り替え時、公開日が空なら今日を設定
  if (newStatus === 'published') {
    var pubDate = sheet.getRange(rowNum, ARTICLE_COLUMNS.PUBLISH_DATE + 1).getValue();
    if (!pubDate) {
      var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy.MM.dd');
      sheet.getRange(rowNum, ARTICLE_COLUMNS.PUBLISH_DATE + 1).setValue(today);
    }
  }

  console.log('記事ステータス切替: 行' + rowNum + ' → ' + newStatus);
}

/**
 * 記事を削除
 * @param {number} row
 */
function deleteArticle(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(ARTICLE_SHEET_NAME);
  if (!sheet) return;
  sheet.deleteRow(parseInt(row));
  console.log('記事削除: 行' + row);
}

/**
 * カテゴリ一覧を取得（記事が存在するカテゴリのみ）
 * @returns {Array}
 */
function getArticleCategories() {
  var result = getArticles({});
  var cats = {};
  result.articles.forEach(function(a) {
    if (a.category) cats[a.category] = (cats[a.category] || 0) + 1;
  });
  return Object.keys(cats).map(function(c) {
    return { name: c, count: cats[c] };
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 管理画面
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 記事管理ページを生成
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function generateArticleAdminPage() {
  var result = getArticles({ includeAll: true });
  var articles = result.articles;
  var baseUrl = CONFIG.CONSENT.WEB_APP_URL;

  var html = '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>記事管理</title>' +
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
    'input[type="text"], textarea, select { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; font-family: inherit; }' +
    'textarea { min-height: 120px; resize: vertical; }' +
    '.btn { display: inline-block; padding: 8px 20px; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; text-decoration: none; margin-top: 12px; }' +
    '.btn-primary { background: #0f2350; color: #fff; }' +
    '.btn-secondary { background: #6c757d; color: #fff; }' +
    '.btn-success { background: #28a745; color: #fff; }' +
    '.btn-warning { background: #ffc107; color: #333; }' +
    '.btn-danger { background: #dc3545; color: #fff; }' +
    '.btn-sm { font-size: 12px; padding: 4px 12px; margin-top: 0; }' +
    '.article-item { border-bottom: 1px solid #eee; padding: 12px 16px; }' +
    '.article-item:last-child { border-bottom: none; }' +
    '.article-title { font-size: 15px; font-weight: bold; margin-bottom: 4px; }' +
    '.article-meta { font-size: 12px; color: #666; display: flex; gap: 12px; margin-bottom: 8px; }' +
    '.article-actions { display: flex; gap: 8px; flex-wrap: wrap; }' +
    '.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: bold; }' +
    '.badge-published { background: #d4edda; color: #155724; }' +
    '.badge-draft { background: #fff3cd; color: #856404; }' +
    '.form-row { display: flex; gap: 12px; }' +
    '.form-row > div { flex: 1; }' +
    '</style>' +
    '</head><body>' +
    '<div class="container">' +
    '<h1>記事管理</h1>';

  // 新規投稿フォーム
  html += '<div class="section">' +
    '<div class="section-title">新規記事を追加</div>' +
    '<div class="section-body">' +
    '<form action="' + baseUrl + '" method="get">' +
    '<input type="hidden" name="action" value="article-add">' +
    '<label>タイトル *</label>' +
    '<input type="text" name="title" required placeholder="記事タイトル">' +
    '<div class="form-row">' +
    '<div><label>カテゴリ</label><select name="category">';
  ARTICLE_CATEGORIES.forEach(function(c) {
    html += '<option value="' + c + '">' + c + '</option>';
  });
  html += '</select></div>' +
    '<div><label>著者</label><input type="text" name="author" placeholder="著者名"></div>' +
    '</div>' +
    '<label>タグ（カンマ区切り）</label>' +
    '<input type="text" name="tags" placeholder="経営戦略, 財務, DX">' +
    '<label>要約</label>' +
    '<input type="text" name="summary" placeholder="記事の要約（一覧表示用）">' +
    '<label>本文（Markdown）</label>' +
    '<textarea name="body" placeholder="## 見出し&#10;&#10;本文テキスト&#10;&#10;- リスト項目"></textarea>' +
    '<div class="form-row">' +
    '<div><label>状態</label><select name="status"><option value="draft">下書き</option><option value="published">公開</option></select></div>' +
    '<div><label>公開日</label><input type="text" name="publishDate" placeholder="2026.04.01"></div>' +
    '</div>' +
    '<button type="submit" class="btn btn-primary">追加する</button>' +
    '</form>' +
    '</div></div>';

  // 記事一覧
  html += '<div class="section">' +
    '<div class="section-title">記事一覧 (' + articles.length + '件)</div>';

  if (articles.length === 0) {
    html += '<div class="section-body"><p>記事がありません</p></div>';
  } else {
    articles.forEach(function(a) {
      var badgeClass = a.status === 'published' ? 'badge-published' : 'badge-draft';
      var badgeLabel = a.status === 'published' ? '公開中' : '下書き';

      html += '<div class="article-item">' +
        '<div class="article-title">' + escapeHtml(a.title) + '</div>' +
        '<div class="article-meta">' +
        '<span class="badge ' + badgeClass + '">' + badgeLabel + '</span>' +
        '<span>' + escapeHtml(a.category) + '</span>' +
        '<span>' + (a.publishDate || '未設定') + '</span>' +
        '<span>' + escapeHtml(a.author) + '</span>' +
        '</div>' +
        '<div class="article-actions">';

      if (a.status === 'published') {
        html += '<a href="' + baseUrl + '?action=article-toggle&row=' + a.row + '" class="btn btn-warning btn-sm">非公開にする</a>';
      } else {
        html += '<a href="' + baseUrl + '?action=article-toggle&row=' + a.row + '" class="btn btn-success btn-sm">公開する</a>';
      }
      html += ' <a href="' + baseUrl + '?action=article-delete&row=' + a.row + '" class="btn btn-danger btn-sm" onclick="return confirm(\'本当に削除しますか？\')">削除</a>';
      html += '</div></div>';
    });
  }

  html += '</div></div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('記事管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
