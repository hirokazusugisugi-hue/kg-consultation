/**
 * AI診断管理システム（Phase 4）
 *
 * 6軸評価による経営診断スコアリングと診断レポート管理。
 * Cloud Function（ai_diagnosis）と連携し、文字起こしテキストまたは
 * テキスト入力からAI診断を実行する。
 *
 * 事前設定:
 *   ScriptProperties に以下を設定:
 *   - DIAGNOSIS_CF_URL: ai_diagnosis Cloud Function URL
 *   - DIAGNOSIS_CF_SECRET: 共有シークレット
 */

/**
 * AI診断結果シートの列定義
 */
var DIAGNOSIS_COLUMNS = {
  DIAGNOSIS_ID: 0,    // A: 診断ID
  APP_ID: 1,          // B: 申込ID
  COMPANY: 2,         // C: 企業名
  INDUSTRY: 3,        // D: 業種
  DIAGNOSIS_DATE: 4,  // E: 診断日
  TOTAL_SCORE: 5,     // F: 総合スコア
  STRATEGY: 6,        // G: 経営戦略
  FINANCE: 7,         // H: 財務
  ORGANIZATION: 8,    // I: 組織・人材
  MARKETING: 9,       // J: マーケティング
  OPERATIONS: 10,     // K: 業務プロセス
  DIGITAL: 11,        // L: IT・DX
  CHALLENGES: 12,     // M: 主要課題（JSON）
  RECOMMENDATIONS: 13,// N: 提言サマリ
  REPORT_FILE_ID: 14, // O: レポートファイルID
  SOURCE: 15,         // P: ソース（文字起こし/直接入力）
  TRANSCRIPT_ID: 16   // Q: 文字起こしID
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// シートセットアップ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * AI診断結果シートのセットアップ
 */
function setupDiagnosisSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = 'AI診断結果';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = [
    '診断ID', '申込ID', '企業名', '業種', '診断日',
    '総合スコア', '経営戦略', '財務', '組織・人材',
    'マーケティング', '業務プロセス', 'IT・DX',
    '主要課題', '提言サマリ', 'レポートファイルID',
    'ソース', '文字起こしID'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅
  sheet.setColumnWidth(1, 120);   // 診断ID
  sheet.setColumnWidth(2, 120);   // 申込ID
  sheet.setColumnWidth(3, 150);   // 企業名
  sheet.setColumnWidth(4, 100);   // 業種
  sheet.setColumnWidth(5, 120);   // 診断日
  sheet.setColumnWidth(6, 80);    // 総合スコア
  sheet.setColumnWidth(7, 80);    // 経営戦略
  sheet.setColumnWidth(8, 80);    // 財務
  sheet.setColumnWidth(9, 80);    // 組織・人材
  sheet.setColumnWidth(10, 100);  // マーケティング
  sheet.setColumnWidth(11, 100);  // 業務プロセス
  sheet.setColumnWidth(12, 80);   // IT・DX
  sheet.setColumnWidth(13, 300);  // 主要課題
  sheet.setColumnWidth(14, 300);  // 提言サマリ
  sheet.setColumnWidth(15, 200);  // レポートファイルID
  sheet.setColumnWidth(16, 100);  // ソース
  sheet.setColumnWidth(17, 200);  // 文字起こしID

  sheet.setFrozenRows(1);

  // 条件付き書式: 総合スコアで色分け
  var scoreRange = sheet.getRange(2, DIAGNOSIS_COLUMNS.TOTAL_SCORE + 1, 1000, 1);

  // 80-100: 緑
  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(80)
    .setBackground('#d4edda')
    .setRanges([scoreRange])
    .build();

  // 60-79: 黄
  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(60, 79)
    .setBackground('#fff3cd')
    .setRanges([scoreRange])
    .build();

  // 0-59: 赤
  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(60)
    .setBackground('#f8d7da')
    .setRanges([scoreRange])
    .build();

  sheet.setConditionalFormatRules([rule1, rule2, rule3]);

  console.log('AI診断結果シートのセットアップが完了しました');
  return { success: true, message: 'AI診断結果シートをセットアップしました' };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 診断実行
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 診断IDの生成
 */
function generateDiagnosisId() {
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('AI診断結果');
  if (!sheet) return 'D' + dateStr + '-001';

  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var id = data[i][DIAGNOSIS_COLUMNS.DIAGNOSIS_ID];
    if (id && id.toString().indexOf(dateStr) >= 0) count++;
  }
  return 'D' + dateStr + '-' + String(count + 1).padStart(3, '0');
}

/**
 * Cloud Function経由でAI診断を実行
 * @param {string} text - 診断対象テキスト
 * @param {Object} info - { applicationId, company, industry, theme }
 * @param {string} source - ソース種別 ('transcript' or 'direct')
 * @param {string} transcriptId - 文字起こしファイルID（オプション）
 * @returns {Object} 診断結果
 */
function runAiDiagnosis(text, info, source, transcriptId) {
  var props = PropertiesService.getScriptProperties();
  var cfUrl = props.getProperty('DIAGNOSIS_CF_URL') || (CONFIG.DIAGNOSIS && CONFIG.DIAGNOSIS.CLOUD_FUNCTION_URL) || '';
  var cfSecret = props.getProperty('DIAGNOSIS_CF_SECRET') || (CONFIG.DIAGNOSIS && CONFIG.DIAGNOSIS.CLOUD_FUNCTION_SECRET) || '';

  if (!cfUrl) {
    return { success: false, message: 'AI診断Cloud Function URLが未設定です' };
  }

  var payload = {
    secret: cfSecret,
    text: text,
    company: info.company || '',
    industry: info.industry || '',
    theme: info.theme || '',
    application_id: info.applicationId || ''
  };

  console.log('AI診断リクエスト送信: ' + (info.applicationId || 'direct') + ' -> ' + cfUrl);

  try {
    var response = UrlFetchApp.fetch(cfUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 120
    });

    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      console.error('AI診断エラー (' + code + '): ' + body);
      return { success: false, message: 'AI診断に失敗しました: ' + code };
    }

    var result = JSON.parse(body);
    if (!result.success || !result.diagnosis) {
      return { success: false, message: '診断結果の取得に失敗しました' };
    }

    var diagnosis = result.diagnosis;

    // シートに記録
    var diagnosisId = saveDiagnosisResult(diagnosis, info, source, transcriptId);

    return {
      success: true,
      diagnosisId: diagnosisId,
      diagnosis: diagnosis,
      tokenUsage: result.token_usage || {}
    };

  } catch (e) {
    console.error('AI診断実行エラー:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * 診断結果をシートに保存
 * @returns {string} 診断ID
 */
function saveDiagnosisResult(diagnosis, info, source, transcriptId) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('AI診断結果');
  if (!sheet) {
    setupDiagnosisSheet();
    sheet = ss.getSheetByName('AI診断結果');
  }

  var diagnosisId = generateDiagnosisId();
  var scores = diagnosis.scores || {};

  // 課題をJSON文字列で保存
  var challengesStr = '';
  try {
    challengesStr = JSON.stringify(diagnosis.challenges || []);
  } catch (e) { challengesStr = ''; }

  // 提言サマリ
  var recsStr = '';
  if (diagnosis.recommendations && diagnosis.recommendations.length > 0) {
    recsStr = diagnosis.recommendations.map(function(r) {
      return '【' + (r.timeline || '') + '】' + (r.title || '') + ': ' + (r.description || '');
    }).join('\n');
  }

  sheet.appendRow([
    diagnosisId,
    info.applicationId || '',
    info.company || '',
    info.industry || '',
    new Date(),
    diagnosis.total_score || 0,
    scores.strategy || 3,
    scores.finance || 3,
    scores.organization || 3,
    scores.marketing || 3,
    scores.operations || 3,
    scores.digital || 3,
    challengesStr,
    recsStr,
    '',  // レポートファイルID（後で更新）
    source || 'direct',
    transcriptId || ''
  ]);

  console.log('AI診断結果保存: ' + diagnosisId + ', スコア=' + (diagnosis.total_score || 0));
  return diagnosisId;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 診断結果の取得
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 診断結果一覧を取得
 * @param {Object} options - { company, limit, offset }
 * @returns {Object} { results, total }
 */
function getDiagnosisResults(options) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('AI診断結果');
  if (!sheet || sheet.getLastRow() <= 1) return { results: [], total: 0 };

  var data = sheet.getDataRange().getValues();
  var results = [];

  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var company = row[DIAGNOSIS_COLUMNS.COMPANY] || '';

    // 企業名フィルタ
    if (options.company && company.indexOf(options.company) < 0) continue;

    results.push({
      diagnosisId: row[DIAGNOSIS_COLUMNS.DIAGNOSIS_ID],
      applicationId: row[DIAGNOSIS_COLUMNS.APP_ID],
      company: company,
      industry: row[DIAGNOSIS_COLUMNS.INDUSTRY],
      date: row[DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE] instanceof Date
        ? Utilities.formatDate(row[DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE], 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(row[DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE]),
      totalScore: row[DIAGNOSIS_COLUMNS.TOTAL_SCORE],
      scores: {
        strategy: row[DIAGNOSIS_COLUMNS.STRATEGY],
        finance: row[DIAGNOSIS_COLUMNS.FINANCE],
        organization: row[DIAGNOSIS_COLUMNS.ORGANIZATION],
        marketing: row[DIAGNOSIS_COLUMNS.MARKETING],
        operations: row[DIAGNOSIS_COLUMNS.OPERATIONS],
        digital: row[DIAGNOSIS_COLUMNS.DIGITAL]
      },
      source: row[DIAGNOSIS_COLUMNS.SOURCE]
    });
  }

  var total = results.length;
  var limit = parseInt(options.limit) || 20;
  var offset = parseInt(options.offset) || 0;
  results = results.slice(offset, offset + limit);

  return { results: results, total: total };
}

/**
 * 診断結果詳細を取得（診断IDで検索）
 * @param {string} diagnosisId - 診断ID
 * @returns {Object|null} 診断結果詳細
 */
function getDiagnosisById(diagnosisId) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('AI診断結果');
  if (!sheet || sheet.getLastRow() <= 1) return null;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][DIAGNOSIS_COLUMNS.DIAGNOSIS_ID] === diagnosisId) {
      var row = data[i];
      var challenges = [];
      try { challenges = JSON.parse(row[DIAGNOSIS_COLUMNS.CHALLENGES] || '[]'); } catch (e) {}

      return {
        diagnosisId: row[DIAGNOSIS_COLUMNS.DIAGNOSIS_ID],
        applicationId: row[DIAGNOSIS_COLUMNS.APP_ID],
        company: row[DIAGNOSIS_COLUMNS.COMPANY],
        industry: row[DIAGNOSIS_COLUMNS.INDUSTRY],
        date: row[DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE] instanceof Date
          ? Utilities.formatDate(row[DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE], 'Asia/Tokyo', 'yyyy/MM/dd')
          : String(row[DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE]),
        totalScore: row[DIAGNOSIS_COLUMNS.TOTAL_SCORE],
        scores: {
          strategy: row[DIAGNOSIS_COLUMNS.STRATEGY],
          finance: row[DIAGNOSIS_COLUMNS.FINANCE],
          organization: row[DIAGNOSIS_COLUMNS.ORGANIZATION],
          marketing: row[DIAGNOSIS_COLUMNS.MARKETING],
          operations: row[DIAGNOSIS_COLUMNS.OPERATIONS],
          digital: row[DIAGNOSIS_COLUMNS.DIGITAL]
        },
        challenges: challenges,
        recommendations: row[DIAGNOSIS_COLUMNS.RECOMMENDATIONS],
        reportFileId: row[DIAGNOSIS_COLUMNS.REPORT_FILE_ID],
        source: row[DIAGNOSIS_COLUMNS.SOURCE],
        transcriptId: row[DIAGNOSIS_COLUMNS.TRANSCRIPT_ID]
      };
    }
  }
  return null;
}

/**
 * 企業別の診断履歴を取得（時系列比較用）
 * @param {string} company - 企業名
 * @returns {Array} 診断結果配列（日付昇順）
 */
function getDiagnosisHistory(company) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('AI診断結果');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var history = [];

  for (var i = 1; i < data.length; i++) {
    var rowCompany = data[i][DIAGNOSIS_COLUMNS.COMPANY] || '';
    if (rowCompany !== company) continue;

    history.push({
      diagnosisId: data[i][DIAGNOSIS_COLUMNS.DIAGNOSIS_ID],
      date: data[i][DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE] instanceof Date
        ? Utilities.formatDate(data[i][DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE], 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(data[i][DIAGNOSIS_COLUMNS.DIAGNOSIS_DATE]),
      totalScore: data[i][DIAGNOSIS_COLUMNS.TOTAL_SCORE],
      scores: {
        strategy: data[i][DIAGNOSIS_COLUMNS.STRATEGY],
        finance: data[i][DIAGNOSIS_COLUMNS.FINANCE],
        organization: data[i][DIAGNOSIS_COLUMNS.ORGANIZATION],
        marketing: data[i][DIAGNOSIS_COLUMNS.MARKETING],
        operations: data[i][DIAGNOSIS_COLUMNS.OPERATIONS],
        digital: data[i][DIAGNOSIS_COLUMNS.DIGITAL]
      }
    });
  }

  return history;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 手動操作 & 管理用
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 指定行の予約データからAI診断を手動実行
 * 文字起こし済みの場合は文字起こしを使用
 * @param {number} rowIndex - 行番号（1-based）
 * @returns {Object} 結果
 */
function manualRunDiagnosis(rowIndex) {
  try {
    var rowData = getRowData(rowIndex);
    if (!rowData.id) {
      return { success: false, message: '行 ' + rowIndex + ' にデータがありません' };
    }

    var text = '';
    var source = 'direct';
    var transcriptId = '';

    // 文字起こしファイルがあればそれを使用
    var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    var tfId = sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_FILE_ID + 1).getValue();

    if (tfId) {
      text = getTranscriptText(tfId);
      source = 'transcript';
      transcriptId = tfId;
    }

    if (!text) {
      // 文字起こしがなければ相談内容テキストを使用
      text = '企業名: ' + (rowData.company || '') + '\n' +
        '業種: ' + (rowData.industry || '') + '\n' +
        'テーマ: ' + (rowData.theme || '') + '\n' +
        '相談内容: ' + (rowData.content || '') + '\n';

      if (!rowData.content) {
        return { success: false, message: '文字起こしも相談内容もありません。診断対象テキストが必要です。' };
      }
      source = 'direct';
    }

    var result = runAiDiagnosis(text, {
      applicationId: rowData.id,
      company: rowData.company,
      industry: rowData.industry,
      theme: rowData.theme
    }, source, transcriptId);

    return result;

  } catch (e) {
    console.error('手動AI診断エラー:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * テキスト直接入力からAI診断を実行（API用）
 * @param {Object} params - { text, company, industry, theme }
 * @returns {Object} 結果
 */
function runDirectDiagnosis(params) {
  if (!params.text) {
    return { success: false, message: '診断対象テキストが必要です' };
  }

  return runAiDiagnosis(params.text, {
    applicationId: '',
    company: params.company || '',
    industry: params.industry || '',
    theme: params.theme || ''
  }, 'direct', '');
}

/**
 * AI診断の統計サマリーを取得
 * @returns {Object} 統計データ
 */
function getDiagnosisStats() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('AI診断結果');
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, totalDiagnoses: 0, averageScore: 0, axisAverages: {} };
  }

  var data = sheet.getDataRange().getValues();
  var count = 0;
  var totalScore = 0;
  var axisTotals = { strategy: 0, finance: 0, organization: 0, marketing: 0, operations: 0, digital: 0 };
  var industryScores = {};

  for (var i = 1; i < data.length; i++) {
    count++;
    totalScore += data[i][DIAGNOSIS_COLUMNS.TOTAL_SCORE] || 0;
    axisTotals.strategy += data[i][DIAGNOSIS_COLUMNS.STRATEGY] || 0;
    axisTotals.finance += data[i][DIAGNOSIS_COLUMNS.FINANCE] || 0;
    axisTotals.organization += data[i][DIAGNOSIS_COLUMNS.ORGANIZATION] || 0;
    axisTotals.marketing += data[i][DIAGNOSIS_COLUMNS.MARKETING] || 0;
    axisTotals.operations += data[i][DIAGNOSIS_COLUMNS.OPERATIONS] || 0;
    axisTotals.digital += data[i][DIAGNOSIS_COLUMNS.DIGITAL] || 0;

    var industry = data[i][DIAGNOSIS_COLUMNS.INDUSTRY] || '不明';
    if (!industryScores[industry]) industryScores[industry] = { count: 0, total: 0 };
    industryScores[industry].count++;
    industryScores[industry].total += data[i][DIAGNOSIS_COLUMNS.TOTAL_SCORE] || 0;
  }

  var axisAverages = {};
  for (var key in axisTotals) {
    axisAverages[key] = count > 0 ? Math.round(axisTotals[key] / count * 10) / 10 : 0;
  }

  var industryAvg = {};
  for (var ind in industryScores) {
    industryAvg[ind] = {
      count: industryScores[ind].count,
      averageScore: Math.round(industryScores[ind].total / industryScores[ind].count)
    };
  }

  return {
    success: true,
    totalDiagnoses: count,
    averageScore: count > 0 ? Math.round(totalScore / count) : 0,
    axisAverages: axisAverages,
    industryBreakdown: industryAvg
  };
}

/**
 * AI診断設定確認
 */
function checkDiagnosisSetup() {
  var props = PropertiesService.getScriptProperties();
  return {
    success: true,
    enabled: !!(CONFIG.DIAGNOSIS && CONFIG.DIAGNOSIS.ENABLED),
    cfUrl: !!(props.getProperty('DIAGNOSIS_CF_URL') || (CONFIG.DIAGNOSIS && CONFIG.DIAGNOSIS.CLOUD_FUNCTION_URL)),
    cfSecret: !!(props.getProperty('DIAGNOSIS_CF_SECRET') || (CONFIG.DIAGNOSIS && CONFIG.DIAGNOSIS.CLOUD_FUNCTION_SECRET))
  };
}
