/**
 * 会場管理機能
 * 会場マスタのCRUD、空き状況表示、予約連動
 * 管理用Webページ + ステータス表示API
 */

/**
 * 会場マスタシートのセットアップ
 */
function setupVenueSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  var sheet = ss.getSheetByName(CONFIG.VENUE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.VENUE_SHEET_NAME);
  }

  // ヘッダー設定
  var headers = ['会場ID', '名称', '住所', '収容人数', '設備', '料金', '備考', '利用可能時間', '有効フラグ'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 80);    // 会場ID
  sheet.setColumnWidth(2, 180);   // 名称
  sheet.setColumnWidth(3, 250);   // 住所
  sheet.setColumnWidth(4, 80);    // 収容人数
  sheet.setColumnWidth(5, 200);   // 設備
  sheet.setColumnWidth(6, 120);   // 料金
  sheet.setColumnWidth(7, 200);   // 備考
  sheet.setColumnWidth(8, 150);   // 利用可能時間
  sheet.setColumnWidth(9, 80);    // 有効フラグ

  // 有効フラグのプルダウン
  var activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TRUE', 'FALSE'])
    .build();
  sheet.getRange(2, VENUE_COLUMNS.ACTIVE + 1, 500, 1).setDataValidation(activeRule);

  // 1行目を固定
  sheet.setFrozenRows(1);

  // 初期データ投入（既存LOCATION_OPTIONS_DEFAULTから）
  var initialData = [
    ['V001', 'アプローズタワー', '大阪市北区茶屋町19-19', 20, 'プロジェクター, Wi-Fi, ホワイトボード', '無料（大学提携）', '', '9:00-21:00', 'TRUE'],
    ['V002', 'スミセスペース', '大阪市北区梅田1丁目', 15, 'プロジェクター, Wi-Fi', '有料', '', '9:00-21:00', 'TRUE'],
    ['V003', 'ナレッジサロン', '大阪市北区大深町3-1 グランフロント大阪', 10, 'Wi-Fi, 電源', '会員制', '', '7:00-24:00', 'TRUE']
  ];
  sheet.getRange(2, 1, initialData.length, initialData[0].length).setValues(initialData);

  // 条件付き書式
  var activeRange = sheet.getRange('I2:I500');
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('TRUE')
    .setBackground('#d4edda')
    .setRanges([activeRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('FALSE')
    .setBackground('#f8d7da')
    .setRanges([activeRange])
    .build());
  sheet.setConditionalFormatRules(rules);

  console.log('会場マスタシートのセットアップが完了しました');
  return { success: true, message: '会場マスタシートのセットアップが完了しました' };
}

/**
 * 会場ID生成
 * @returns {string} V001形式のID
 */
function generateVenueId() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.VENUE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return 'V001';

  var data = sheet.getDataRange().getValues();
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][VENUE_COLUMNS.ID]);
    var match = id.match(/^V(\d+)$/);
    if (match) {
      var num = parseInt(match[1]);
      if (num > maxNum) maxNum = num;
    }
  }
  return 'V' + String(maxNum + 1).padStart(3, '0');
}

/**
 * 全会場を取得
 * @param {boolean} activeOnly - trueの場合、有効な会場のみ
 * @returns {Array} 会場データの配列
 */
function getVenues(activeOnly) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.VENUE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var venues = [];
  for (var i = 1; i < data.length; i++) {
    var active = data[i][VENUE_COLUMNS.ACTIVE];
    var isActive = active === true || active === 'TRUE' || active === 'true';
    if (activeOnly && !isActive) continue;

    venues.push({
      row: i + 1,
      id: String(data[i][VENUE_COLUMNS.ID]),
      name: String(data[i][VENUE_COLUMNS.NAME]),
      address: String(data[i][VENUE_COLUMNS.ADDRESS]),
      capacity: data[i][VENUE_COLUMNS.CAPACITY],
      equipment: String(data[i][VENUE_COLUMNS.EQUIPMENT]),
      price: String(data[i][VENUE_COLUMNS.PRICE]),
      notes: String(data[i][VENUE_COLUMNS.NOTES]),
      hours: String(data[i][VENUE_COLUMNS.HOURS]),
      active: isActive
    });
  }
  return venues;
}

/**
 * 会場IDで会場を取得
 * @param {string} venueId - 会場ID
 * @returns {Object|null}
 */
function getVenueById(venueId) {
  var venues = getVenues(false);
  for (var i = 0; i < venues.length; i++) {
    if (venues[i].id === venueId) return venues[i];
  }
  return null;
}

/**
 * 会場名で会場情報を取得
 * @param {string} venueName - 会場名
 * @returns {Object|null}
 */
function getVenueByName(venueName) {
  var venues = getVenues(false);
  for (var i = 0; i < venues.length; i++) {
    if (venues[i].name === venueName) return venues[i];
  }
  return null;
}

/**
 * 会場を追加
 * @param {Object} params - 会場データ
 * @returns {Object} 結果
 */
function addVenue(params) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.VENUE_SHEET_NAME);
  if (!sheet) {
    setupVenueSheet();
    sheet = ss.getSheetByName(CONFIG.VENUE_SHEET_NAME);
  }

  var venueId = generateVenueId();
  var newRow = [
    venueId,
    params.name || '',
    params.address || '',
    params.capacity || '',
    params.equipment || '',
    params.price || '',
    params.notes || '',
    params.hours || '',
    'TRUE'
  ];

  sheet.appendRow(newRow);
  console.log('会場追加: ' + venueId + ' - ' + params.name);
  return { success: true, venueId: venueId, message: '会場を追加しました: ' + params.name };
}

/**
 * 会場を更新
 * @param {number} row - 行番号
 * @param {Object} params - 更新データ
 * @returns {Object} 結果
 */
function updateVenue(row, params) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.VENUE_SHEET_NAME);
  if (!sheet) return { success: false, message: '会場マスタシートが見つかりません' };

  var rowNum = parseInt(row);
  if (rowNum < 2) return { success: false, message: '無効な行番号です' };

  if (params.name !== undefined) sheet.getRange(rowNum, VENUE_COLUMNS.NAME + 1).setValue(params.name);
  if (params.address !== undefined) sheet.getRange(rowNum, VENUE_COLUMNS.ADDRESS + 1).setValue(params.address);
  if (params.capacity !== undefined) sheet.getRange(rowNum, VENUE_COLUMNS.CAPACITY + 1).setValue(params.capacity);
  if (params.equipment !== undefined) sheet.getRange(rowNum, VENUE_COLUMNS.EQUIPMENT + 1).setValue(params.equipment);
  if (params.price !== undefined) sheet.getRange(rowNum, VENUE_COLUMNS.PRICE + 1).setValue(params.price);
  if (params.notes !== undefined) sheet.getRange(rowNum, VENUE_COLUMNS.NOTES + 1).setValue(params.notes);
  if (params.hours !== undefined) sheet.getRange(rowNum, VENUE_COLUMNS.HOURS + 1).setValue(params.hours);

  console.log('会場更新: 行' + rowNum);
  return { success: true, message: '会場を更新しました' };
}

/**
 * 会場の有効/無効を切り替え
 * @param {number} row - 行番号
 */
function toggleVenueActive(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.VENUE_SHEET_NAME);
  if (!sheet) return;

  var rowNum = parseInt(row);
  var current = sheet.getRange(rowNum, VENUE_COLUMNS.ACTIVE + 1).getValue();
  var newValue = (current === true || current === 'TRUE' || current === 'true') ? 'FALSE' : 'TRUE';
  sheet.getRange(rowNum, VENUE_COLUMNS.ACTIVE + 1).setValue(newValue);

  console.log('会場有効切替: 行' + rowNum + ' → ' + newValue);
}

/**
 * 会場を削除
 * @param {number} row - 行番号
 */
function deleteVenue(row) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.VENUE_SHEET_NAME);
  if (!sheet) return;

  var rowNum = parseInt(row);
  sheet.deleteRow(rowNum);
  console.log('会場削除: 行' + rowNum);
}

/**
 * 会場別の予約状況を取得
 * 日程設定シートの場所(LOCATION)列と予約状況を突き合わせ
 * @param {string} venueFilter - 会場名フィルタ（省略時は全会場）
 * @param {number} targetMonth - 月フィルタ（省略時は当月+翌月）
 * @returns {Array} 予約状況データ
 */
function getVenueBookingStatus(venueFilter, targetMonth) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var schedSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  var bookSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!schedSheet) return [];

  var schedData = schedSheet.getDataRange().getValues();
  var now = new Date();
  var results = [];

  for (var i = 1; i < schedData.length; i++) {
    var dateVal = schedData[i][SCHEDULE_COLUMNS.DATE];
    if (!dateVal) continue;

    var d = dateVal instanceof Date ? dateVal : new Date(dateVal);

    // 月フィルタ
    if (targetMonth) {
      if ((d.getMonth() + 1) !== targetMonth) continue;
    } else {
      // デフォルト: 今月と来月のみ
      var monthDiff = (d.getFullYear() - now.getFullYear()) * 12 + (d.getMonth() - now.getMonth());
      if (monthDiff < 0 || monthDiff > 1) continue;
    }

    var dateStr = dateVal instanceof Date
      ? Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/MM/dd')
      : String(dateVal);

    var timeVal = schedData[i][SCHEDULE_COLUMNS.TIME];
    var timeStr = timeVal instanceof Date
      ? Utilities.formatDate(timeVal, 'Asia/Tokyo', 'HH:mm')
      : String(timeVal);

    var method = schedData[i][SCHEDULE_COLUMNS.METHOD] || '';
    var bookingStatus = schedData[i][SCHEDULE_COLUMNS.BOOKING_STATUS] || '空き';
    var available = schedData[i][SCHEDULE_COLUMNS.AVAILABLE] || '';
    var bookable = schedData[i][SCHEDULE_COLUMNS.BOOKABLE] || '';

    // 対面 or 両方の枠のみ会場が関係する
    var needsVenue = (method === '対面' || method === '両方');

    results.push({
      date: dateStr,
      time: timeStr,
      method: method,
      bookingStatus: bookingStatus,
      available: available,
      bookable: bookable,
      needsVenue: needsVenue,
      dayOfWeek: ['日', '月', '火', '水', '木', '金', '土'][d.getDay()]
    });
  }

  // 日付+時間順にソート
  results.sort(function(a, b) {
    return (a.date + a.time).localeCompare(b.date + b.time);
  });

  return results;
}

/**
 * 日程設定シートにVENUE列を追加（拡張用: 将来の会場×日程マッピング）
 * 現時点では予約管理シートのN列(場所)で管理
 */

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 管理画面
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 会場管理ページを生成
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function generateVenueAdminPage() {
  var venues = getVenues(false);
  var baseUrl = CONFIG.CONSENT.WEB_APP_URL;

  var html = '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>会場管理</title>' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif; background: #f5f5f5; color: #333; padding: 16px; }' +
    'h1 { background: #0f2350; color: #fff; padding: 16px; border-radius: 8px 8px 0 0; font-size: 18px; }' +
    '.container { max-width: 800px; margin: 0 auto; }' +
    '.section { background: #fff; border-radius: 8px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden; }' +
    '.section-title { background: #e8eaf0; padding: 12px 16px; font-weight: bold; font-size: 14px; }' +
    '.section-body { padding: 16px; }' +
    'label { display: block; font-size: 13px; font-weight: bold; margin-bottom: 4px; margin-top: 8px; }' +
    'label:first-child { margin-top: 0; }' +
    'input[type="text"], input[type="number"] { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }' +
    '.form-row { display: flex; gap: 12px; }' +
    '.form-row > div { flex: 1; }' +
    '.btn { display: inline-block; padding: 8px 20px; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; text-decoration: none; margin-top: 12px; }' +
    '.btn-primary { background: #0f2350; color: #fff; }' +
    '.btn-secondary { background: #6c757d; color: #fff; }' +
    '.btn-warning { background: #ffc107; color: #333; }' +
    '.btn-danger { background: #dc3545; color: #fff; }' +
    '.btn-sm { font-size: 12px; padding: 4px 12px; margin-top: 0; }' +
    '.venue-card { border: 1px solid #eee; border-radius: 8px; padding: 16px; margin-bottom: 12px; }' +
    '.venue-card.inactive { opacity: 0.5; background: #f8f8f8; }' +
    '.venue-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }' +
    '.venue-name { font-size: 16px; font-weight: bold; }' +
    '.venue-id { font-size: 12px; color: #999; margin-left: 8px; }' +
    '.venue-detail { font-size: 13px; color: #555; margin-bottom: 4px; }' +
    '.venue-detail span { color: #999; margin-right: 4px; }' +
    '.venue-actions { display: flex; gap: 8px; margin-top: 8px; flex-wrap: wrap; }' +
    '.badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: bold; }' +
    '.badge-active { background: #d4edda; color: #155724; }' +
    '.badge-inactive { background: #f8d7da; color: #721c24; }' +
    '.nav-links { margin-bottom: 16px; display: flex; gap: 8px; }' +
    '.nav-links a { font-size: 13px; }' +
    '</style>' +
    '</head><body>' +
    '<div class="container">' +
    '<h1>会場管理</h1>';

  // ナビゲーション
  html += '<div class="section"><div class="section-body">' +
    '<div class="nav-links">' +
    '<a href="' + baseUrl + '?action=venue-admin" class="btn btn-primary btn-sm">会場管理</a>' +
    '<a href="' + baseUrl + '?action=venue-status" class="btn btn-secondary btn-sm">空き状況</a>' +
    '</div></div></div>';

  // 新規追加フォーム
  html += '<div class="section">' +
    '<div class="section-title">会場を追加</div>' +
    '<div class="section-body">' +
    '<form action="' + baseUrl + '" method="get">' +
    '<input type="hidden" name="action" value="venue-add">' +
    '<label>名称 *</label>' +
    '<input type="text" name="name" required placeholder="例: アプローズタワー">' +
    '<label>住所</label>' +
    '<input type="text" name="address" placeholder="例: 大阪市北区茶屋町19-19">' +
    '<div class="form-row">' +
    '<div><label>収容人数</label><input type="number" name="capacity" placeholder="例: 20"></div>' +
    '<div><label>料金</label><input type="text" name="price" placeholder="例: 無料"></div>' +
    '</div>' +
    '<label>設備</label>' +
    '<input type="text" name="equipment" placeholder="例: プロジェクター, Wi-Fi, ホワイトボード">' +
    '<label>利用可能時間</label>' +
    '<input type="text" name="hours" placeholder="例: 9:00-21:00">' +
    '<label>備考</label>' +
    '<input type="text" name="notes" placeholder="備考">' +
    '<button type="submit" class="btn btn-primary">追加する</button>' +
    '</form>' +
    '</div></div>';

  // 会場一覧
  html += '<div class="section">' +
    '<div class="section-title">登録済み会場 (' + venues.length + '件)</div>';

  if (venues.length === 0) {
    html += '<div class="section-body"><p>会場が登録されていません</p></div>';
  } else {
    html += '<div class="section-body">';
    venues.forEach(function(v) {
      var cardClass = v.active ? 'venue-card' : 'venue-card inactive';
      html += '<div class="' + cardClass + '">' +
        '<div class="venue-header">' +
        '<div><span class="venue-name">' + escapeHtml(v.name) + '</span><span class="venue-id">' + v.id + '</span></div>';

      if (v.active) {
        html += '<span class="badge badge-active">有効</span>';
      } else {
        html += '<span class="badge badge-inactive">無効</span>';
      }
      html += '</div>';

      if (v.address) html += '<div class="venue-detail"><span>住所:</span>' + escapeHtml(v.address) + '</div>';
      if (v.capacity) html += '<div class="venue-detail"><span>収容:</span>' + v.capacity + '名</div>';
      if (v.equipment) html += '<div class="venue-detail"><span>設備:</span>' + escapeHtml(v.equipment) + '</div>';
      if (v.price) html += '<div class="venue-detail"><span>料金:</span>' + escapeHtml(v.price) + '</div>';
      if (v.hours) html += '<div class="venue-detail"><span>時間:</span>' + escapeHtml(v.hours) + '</div>';
      if (v.notes) html += '<div class="venue-detail"><span>備考:</span>' + escapeHtml(v.notes) + '</div>';

      html += '<div class="venue-actions">';
      if (v.active) {
        html += '<a href="' + baseUrl + '?action=venue-toggle&row=' + v.row + '" class="btn btn-warning btn-sm">無効にする</a>';
      } else {
        html += '<a href="' + baseUrl + '?action=venue-toggle&row=' + v.row + '" class="btn btn-secondary btn-sm">有効にする</a>' +
          ' <a href="' + baseUrl + '?action=venue-delete&row=' + v.row + '" class="btn btn-danger btn-sm" onclick="return confirm(\'本当に削除しますか？\')">削除</a>';
      }
      html += '</div></div>';
    });
    html += '</div>';
  }

  html += '</div></div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('会場管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 空き状況・予約ダッシュボードページを生成
 * @param {Object} e - リクエストパラメータ
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function generateVenueStatusPage(e) {
  var targetMonth = e && e.parameter && e.parameter.month ? parseInt(e.parameter.month) : null;
  var bookingData = getVenueBookingStatus(null, targetMonth);
  var venues = getVenues(true);
  var baseUrl = CONFIG.CONSENT.WEB_APP_URL;

  var now = new Date();
  var currentMonth = now.getMonth() + 1;
  var currentYear = now.getFullYear();

  // 月ごとにグループ化
  var byMonth = {};
  bookingData.forEach(function(item) {
    var parts = item.date.split('/');
    var monthKey = parts[0] + '/' + parts[1];
    if (!byMonth[monthKey]) byMonth[monthKey] = [];
    byMonth[monthKey].push(item);
  });

  var html = '<!DOCTYPE html>' +
    '<html><head>' +
    '<meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>会場 空き状況</title>' +
    '<style>' +
    '* { box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Hiragino Sans", "Hiragino Kaku Gothic ProN", "Noto Sans JP", sans-serif; background: #f5f5f5; color: #333; padding: 16px; }' +
    'h1 { background: #0f2350; color: #fff; padding: 16px; border-radius: 8px 8px 0 0; font-size: 18px; }' +
    '.container { max-width: 800px; margin: 0 auto; }' +
    '.section { background: #fff; border-radius: 8px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); overflow: hidden; }' +
    '.section-title { background: #e8eaf0; padding: 12px 16px; font-weight: bold; font-size: 14px; }' +
    '.section-body { padding: 16px; }' +
    '.btn { display: inline-block; padding: 8px 20px; border: none; border-radius: 4px; font-size: 14px; cursor: pointer; text-decoration: none; }' +
    '.btn-primary { background: #0f2350; color: #fff; }' +
    '.btn-secondary { background: #6c757d; color: #fff; }' +
    '.btn-sm { font-size: 12px; padding: 4px 12px; }' +
    '.nav-links { margin-bottom: 16px; display: flex; gap: 8px; }' +
    '.month-filter { display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap; }' +
    '.month-filter a { padding: 4px 12px; border-radius: 16px; font-size: 13px; text-decoration: none; border: 1px solid #ccc; color: #333; }' +
    '.month-filter a.active { background: #0f2350; color: #fff; border-color: #0f2350; }' +
    'table { width: 100%; border-collapse: collapse; font-size: 13px; }' +
    'th { background: #f0f0f0; padding: 8px 12px; text-align: left; border-bottom: 2px solid #ddd; }' +
    'td { padding: 8px 12px; border-bottom: 1px solid #eee; }' +
    '.status-available { color: #28a745; font-weight: bold; }' +
    '.status-booked { color: #dc3545; font-weight: bold; }' +
    '.status-unavailable { color: #999; }' +
    '.method-badge { display: inline-block; padding: 1px 6px; border-radius: 4px; font-size: 11px; }' +
    '.method-both { background: #e0f0ff; color: #0066cc; }' +
    '.method-online { background: #fff3e0; color: #e65100; }' +
    '.method-visit { background: #e8f5e9; color: #2e7d32; }' +
    '.venue-legend { display: flex; gap: 16px; margin-bottom: 12px; flex-wrap: wrap; font-size: 12px; }' +
    '.venue-legend-item { display: flex; align-items: center; gap: 4px; }' +
    '.venue-legend-dot { width: 12px; height: 12px; border-radius: 50%; display: inline-block; }' +
    '</style>' +
    '</head><body>' +
    '<div class="container">' +
    '<h1>空き状況</h1>';

  // ナビゲーション
  html += '<div class="section"><div class="section-body">' +
    '<div class="nav-links">' +
    '<a href="' + baseUrl + '?action=venue-admin" class="btn btn-secondary btn-sm">会場管理</a>' +
    '<a href="' + baseUrl + '?action=venue-status" class="btn btn-primary btn-sm">空き状況</a>' +
    '</div></div></div>';

  // 月フィルタ
  html += '<div class="section"><div class="section-body">' +
    '<div class="month-filter">';
  html += '<a href="' + baseUrl + '?action=venue-status"' +
    (!targetMonth ? ' class="active"' : '') + '>今月+来月</a>';
  for (var m = 0; m < 3; m++) {
    var filterMonth = ((currentMonth - 1 + m) % 12) + 1;
    var filterYear = currentYear + Math.floor((currentMonth - 1 + m) / 12);
    html += '<a href="' + baseUrl + '?action=venue-status&month=' + filterMonth + '"' +
      (targetMonth === filterMonth ? ' class="active"' : '') + '>' +
      filterYear + '年' + filterMonth + '月</a>';
  }
  html += '</div></div></div>';

  // 会場一覧サマリー
  html += '<div class="section">' +
    '<div class="section-title">登録会場</div>' +
    '<div class="section-body">' +
    '<div class="venue-legend">';
  venues.forEach(function(v) {
    html += '<div class="venue-legend-item">' +
      '<span>' + escapeHtml(v.name) + '</span>' +
      (v.capacity ? '<span style="color:#999">(' + v.capacity + '名)</span>' : '') +
      '</div>';
  });
  html += '</div></div></div>';

  // 日程テーブル
  var monthKeys = Object.keys(byMonth).sort();
  if (monthKeys.length === 0) {
    html += '<div class="section"><div class="section-body"><p>表示する日程がありません</p></div></div>';
  } else {
    monthKeys.forEach(function(monthKey) {
      var items = byMonth[monthKey];
      html += '<div class="section">' +
        '<div class="section-title">' + monthKey.replace('/', '年') + '月</div>' +
        '<div class="section-body" style="overflow-x:auto;">' +
        '<table>' +
        '<tr><th>日付</th><th>曜日</th><th>時間</th><th>方法</th><th>予約可能</th><th>状況</th></tr>';

      items.forEach(function(item) {
        var statusClass = '';
        var statusText = '';
        if (item.bookingStatus === '予約済み') {
          statusClass = 'status-booked';
          statusText = '予約済み';
        } else if (item.available === '○' && item.bookable === '○') {
          statusClass = 'status-available';
          statusText = '空き';
        } else {
          statusClass = 'status-unavailable';
          statusText = item.available !== '○' ? '対応不可' : '人員不足';
        }

        var methodClass = 'method-both';
        if (item.method === 'オンライン') methodClass = 'method-online';
        else if (item.method === '対面') methodClass = 'method-visit';

        html += '<tr>' +
          '<td>' + item.date.split('/').slice(1).join('/') + '</td>' +
          '<td>' + item.dayOfWeek + '</td>' +
          '<td>' + item.time + '</td>' +
          '<td><span class="method-badge ' + methodClass + '">' + item.method + '</span></td>' +
          '<td>' + (item.bookable === '○' ? '<span style="color:#28a745">○</span>' : '<span style="color:#999">×</span>') + '</td>' +
          '<td><span class="' + statusClass + '">' + statusText + '</span></td>' +
          '</tr>';
      });

      html += '</table></div></div>';
    });
  }

  html += '</div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('空き状況')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 会場情報をフォーマットして返す（メール本文用）
 * @param {string} venueName - 会場名
 * @returns {string} フォーマット済み会場情報
 */
function formatVenueInfoForEmail(venueName) {
  if (!venueName || venueName === 'その他') return venueName || '';

  var venue = getVenueByName(venueName);
  if (!venue) return venueName;

  var info = venue.name;
  if (venue.address) info += '\n  住所: ' + venue.address;
  if (venue.hours) info += '\n  利用時間: ' + venue.hours;
  return info;
}
