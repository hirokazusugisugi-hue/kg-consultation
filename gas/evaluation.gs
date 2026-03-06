/**
 * コンサルタント評価管理システム（Phase 4 v2）
 *
 * ICMCI CMC・Schein理論・SERVQUAL・MITIに基づく22項目+3項目の学術的評価フレームワーク。
 * Cloud Function（consultation_evaluation）と連携し、文字起こしテキストから
 * 6カテゴリ×6回のClaude API呼出による精密評価を実行する。
 *
 * スコア構成:
 *   AI評価（22項目）: 0-90点
 *   人間評価（3項目）: 0-10点
 *   総合スコア: 0-100点
 *
 * 事前設定:
 *   ScriptProperties に以下を設定:
 *   - EVALUATION_CF_URL: consultation_evaluation Cloud Function URL
 *   - EVALUATION_CF_SECRET: 共有シークレット
 */

/**
 * コンサルタント評価シートの列定義（A-Z = 26列）
 */
var EVAL_COLUMNS = {
  EVAL_ID: 0,             // A: 評価ID (E20260305-001)
  APP_ID: 1,              // B: 申込ID
  CONSULTANT_NAME: 2,     // C: コンサルタント名
  COMPANY_NAME: 3,        // D: 相談企業名
  EVAL_DATE: 4,           // E: 評価日
  TOTAL_SCORE: 5,         // F: 総合スコア (0-100)
  AI_TOTAL: 6,            // G: AI合計 (0-90)
  HUMAN_TOTAL: 7,         // H: 人間評価合計 (0-10)
  C1_PROBLEM: 8,          // I: C1 問題把握
  C2_SOLUTION: 9,         // J: C2 解決策
  C3_COMMUNICATION: 10,   // K: C3 コミュニケーション
  C4_TIME: 11,            // L: C4 時間管理
  C5_LOGIC: 12,           // M: C5 論理的展開
  C6_ETHICS: 13,          // N: C6 倫理・自律性
  H1_ATMOSPHERE: 14,      // O: H1 雰囲気 (0-4)
  H2_ATTITUDE: 15,        // P: H2 態度変化 (0-3)
  H3_APPEARANCE: 16,      // Q: H3 身だしなみ (0-3)
  ITEM_SCORES_JSON: 17,   // R: 22項目個別スコアJSON
  EVIDENCE_JSON: 18,      // S: エビデンスJSON
  NG_WORDS_JSON: 19,      // T: NG語句検出JSON
  TRANSCRIPT_FILE_ID: 20, // U: 文字起こしファイルID
  REPORT_FILE_ID: 21,     // V: 評価レポートファイルID
  STATUS: 22,             // W: ステータス
  HUMAN_EVALUATOR: 23,    // X: 人間評価者名
  CREATED_AT: 24,         // Y: 作成日時
  UPDATED_AT: 25          // Z: 更新日時
};

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// シートセットアップ
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * コンサルタント評価シートのセットアップ
 */
function setupEvaluationSheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.EVALUATION.SHEET_NAME || 'コンサルタント評価';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = [
    '評価ID', '申込ID', 'コンサルタント名', '相談企業名', '評価日',
    '総合スコア', 'AI合計', '人間評価合計',
    'C1:問題把握', 'C2:解決策', 'C3:コミュニケーション', 'C4:時間管理', 'C5:論理的展開', 'C6:倫理・自律性',
    'H1:雰囲気', 'H2:態度変化', 'H3:身だしなみ',
    '22項目スコアJSON', 'エビデンスJSON', 'NG語句JSON',
    '文字起こしID', 'レポートID', 'ステータス', '人間評価者',
    '作成日時', '更新日時'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅設定
  sheet.setColumnWidth(1, 140);   // 評価ID
  sheet.setColumnWidth(2, 120);   // 申込ID
  sheet.setColumnWidth(3, 120);   // コンサルタント名
  sheet.setColumnWidth(4, 150);   // 相談企業名
  sheet.setColumnWidth(5, 100);   // 評価日
  sheet.setColumnWidth(6, 80);    // 総合スコア
  sheet.setColumnWidth(7, 80);    // AI合計
  sheet.setColumnWidth(8, 80);    // 人間評価合計
  var i;
  for (i = 9; i <= 14; i++) sheet.setColumnWidth(i, 80);  // C1-C6
  for (i = 15; i <= 17; i++) sheet.setColumnWidth(i, 80); // H1-H3
  for (i = 18; i <= 20; i++) sheet.setColumnWidth(i, 200); // JSON列
  sheet.setColumnWidth(21, 200);  // 文字起こしID
  sheet.setColumnWidth(22, 200);  // レポートID
  sheet.setColumnWidth(23, 100);  // ステータス
  sheet.setColumnWidth(24, 100);  // 人間評価者
  sheet.setColumnWidth(25, 150);  // 作成日時
  sheet.setColumnWidth(26, 150);  // 更新日時

  sheet.setFrozenRows(1);

  // 条件付き書式: 総合スコアで色分け
  var scoreRange = sheet.getRange(2, EVAL_COLUMNS.TOTAL_SCORE + 1, 1000, 1);

  var rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(80)
    .setBackground('#d4edda')
    .setRanges([scoreRange])
    .build();

  var rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(60, 79)
    .setBackground('#fff3cd')
    .setRanges([scoreRange])
    .build();

  var rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(60)
    .setBackground('#f8d7da')
    .setRanges([scoreRange])
    .build();

  // ステータス条件付き書式
  var statusRange = sheet.getRange(2, EVAL_COLUMNS.STATUS + 1, 1000, 1);

  var rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(EVALUATION_STATUS.COMPLETE)
    .setBackground('#d4edda')
    .setRanges([statusRange])
    .build();

  var rule5 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(EVALUATION_STATUS.ERROR)
    .setBackground('#f8d7da')
    .setRanges([statusRange])
    .build();

  sheet.setConditionalFormatRules([rule1, rule2, rule3, rule4, rule5]);

  console.log('コンサルタント評価シートのセットアップが完了しました');
  return { success: true, message: 'コンサルタント評価シートをセットアップしました' };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ID生成
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 評価ID生成 (E{yyyyMMdd}-{NNN}形式)
 */
function generateEvaluationId() {
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.EVALUATION.SHEET_NAME || 'コンサルタント評価';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 'E' + dateStr + '-001';

  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var id = data[i][EVAL_COLUMNS.EVAL_ID];
    if (id && id.toString().indexOf(dateStr) >= 0) count++;
  }
  return 'E' + dateStr + '-' + String(count + 1).padStart(3, '0');
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 評価実行
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * メイン評価処理: 文字起こし取得 → Cloud Function呼出 → 結果保存
 * @param {string} transcriptFileId - DriveファイルID
 * @param {number} rowIndex - 予約管理シート行番号
 * @returns {Object} 結果
 */
function runConsultationEvaluation(transcriptFileId, rowIndex) {
  var props = PropertiesService.getScriptProperties();
  var cfUrl = props.getProperty('EVALUATION_CF_URL') || (CONFIG.EVALUATION && CONFIG.EVALUATION.CLOUD_FUNCTION_URL) || '';
  var cfSecret = props.getProperty('EVALUATION_CF_SECRET') || (CONFIG.EVALUATION && CONFIG.EVALUATION.CLOUD_FUNCTION_SECRET) || '';

  if (!cfUrl) {
    return { success: false, message: 'コンサルタント評価Cloud Function URLが未設定です' };
  }

  // 文字起こしテキスト取得
  var transcript = getTranscriptText(transcriptFileId);
  if (!transcript) {
    return { success: false, message: '文字起こしテキストの読み込みに失敗しました' };
  }

  // 予約データ取得
  var rowData = getRowData(rowIndex);
  if (!rowData.id) {
    return { success: false, message: '行 ' + rowIndex + ' にデータがありません' };
  }

  // 評価IDを事前生成
  var evalId = generateEvaluationId();

  // ステータス更新: AI評価中
  var evalSheet = getOrCreateEvaluationSheet_();
  var newRowNum = evalSheet.getLastRow() + 1;
  evalSheet.getRange(newRowNum, EVAL_COLUMNS.EVAL_ID + 1).setValue(evalId);
  evalSheet.getRange(newRowNum, EVAL_COLUMNS.APP_ID + 1).setValue(rowData.id || '');
  evalSheet.getRange(newRowNum, EVAL_COLUMNS.CONSULTANT_NAME + 1).setValue(rowData.leader || '');
  evalSheet.getRange(newRowNum, EVAL_COLUMNS.COMPANY_NAME + 1).setValue(rowData.company || '');
  evalSheet.getRange(newRowNum, EVAL_COLUMNS.STATUS + 1).setValue(EVALUATION_STATUS.AI_RUNNING);
  evalSheet.getRange(newRowNum, EVAL_COLUMNS.CREATED_AT + 1).setValue(new Date());
  evalSheet.getRange(newRowNum, EVAL_COLUMNS.TRANSCRIPT_FILE_ID + 1).setValue(transcriptFileId);

  var payload = {
    secret: cfSecret,
    transcript: transcript,
    evaluation_id: evalId,
    application_id: rowData.id || '',
    consultant_name: rowData.leader || '',
    company_name: rowData.company || '',
    industry: rowData.industry || '',
    theme: rowData.theme || ''
  };

  console.log('コンサルタント評価リクエスト送信: ' + evalId + ' -> ' + cfUrl);

  try {
    var response = UrlFetchApp.fetch(cfUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
      timeout: 600
    });

    var code = response.getResponseCode();
    var body = response.getContentText();

    if (code !== 200) {
      console.error('コンサルタント評価エラー (' + code + '): ' + body);
      evalSheet.getRange(newRowNum, EVAL_COLUMNS.STATUS + 1).setValue(EVALUATION_STATUS.ERROR);
      evalSheet.getRange(newRowNum, EVAL_COLUMNS.UPDATED_AT + 1).setValue(new Date());
      return { success: false, message: '評価に失敗しました: ' + code };
    }

    var result = JSON.parse(body);
    if (!result.success) {
      evalSheet.getRange(newRowNum, EVAL_COLUMNS.STATUS + 1).setValue(EVALUATION_STATUS.ERROR);
      evalSheet.getRange(newRowNum, EVAL_COLUMNS.UPDATED_AT + 1).setValue(new Date());
      return { success: false, message: '評価結果の取得に失敗しました' };
    }

    // 結果保存
    saveEvaluationResult(result, newRowNum, evalSheet, transcriptFileId);

    return {
      success: true,
      evaluationId: evalId,
      totalScore: result.ai_total || 0,
      message: '評価が完了しました（AI評価: ' + (result.ai_total || 0) + '/90点）'
    };

  } catch (e) {
    console.error('コンサルタント評価実行エラー:', e);
    evalSheet.getRange(newRowNum, EVAL_COLUMNS.STATUS + 1).setValue(EVALUATION_STATUS.ERROR);
    evalSheet.getRange(newRowNum, EVAL_COLUMNS.UPDATED_AT + 1).setValue(new Date());
    return { success: false, error: e.toString() };
  }
}

/**
 * CF応答をシートに保存
 */
function saveEvaluationResult(cfResult, rowNum, sheet, transcriptFileId) {
  var now = new Date();
  var aiTotal = cfResult.ai_total || 0;

  sheet.getRange(rowNum, EVAL_COLUMNS.EVAL_DATE + 1).setValue(now);
  sheet.getRange(rowNum, EVAL_COLUMNS.AI_TOTAL + 1).setValue(aiTotal);
  sheet.getRange(rowNum, EVAL_COLUMNS.TOTAL_SCORE + 1).setValue(aiTotal); // 人間採点前はAI分のみ
  sheet.getRange(rowNum, EVAL_COLUMNS.HUMAN_TOTAL + 1).setValue(0);

  // カテゴリ別小計
  var categories = cfResult.categories || {};
  sheet.getRange(rowNum, EVAL_COLUMNS.C1_PROBLEM + 1).setValue(categories.c1 || 0);
  sheet.getRange(rowNum, EVAL_COLUMNS.C2_SOLUTION + 1).setValue(categories.c2 || 0);
  sheet.getRange(rowNum, EVAL_COLUMNS.C3_COMMUNICATION + 1).setValue(categories.c3 || 0);
  sheet.getRange(rowNum, EVAL_COLUMNS.C4_TIME + 1).setValue(categories.c4 || 0);
  sheet.getRange(rowNum, EVAL_COLUMNS.C5_LOGIC + 1).setValue(categories.c5 || 0);
  sheet.getRange(rowNum, EVAL_COLUMNS.C6_ETHICS + 1).setValue(categories.c6 || 0);

  // 人間評価（初期値）
  sheet.getRange(rowNum, EVAL_COLUMNS.H1_ATMOSPHERE + 1).setValue(0);
  sheet.getRange(rowNum, EVAL_COLUMNS.H2_ATTITUDE + 1).setValue(0);
  sheet.getRange(rowNum, EVAL_COLUMNS.H3_APPEARANCE + 1).setValue(0);

  // JSON列
  try {
    sheet.getRange(rowNum, EVAL_COLUMNS.ITEM_SCORES_JSON + 1).setValue(JSON.stringify(cfResult.item_scores || {}));
  } catch (e) { sheet.getRange(rowNum, EVAL_COLUMNS.ITEM_SCORES_JSON + 1).setValue('{}'); }

  try {
    sheet.getRange(rowNum, EVAL_COLUMNS.EVIDENCE_JSON + 1).setValue(JSON.stringify(cfResult.evidence || {}));
  } catch (e) { sheet.getRange(rowNum, EVAL_COLUMNS.EVIDENCE_JSON + 1).setValue('{}'); }

  try {
    sheet.getRange(rowNum, EVAL_COLUMNS.NG_WORDS_JSON + 1).setValue(JSON.stringify(cfResult.ng_words || []));
  } catch (e) { sheet.getRange(rowNum, EVAL_COLUMNS.NG_WORDS_JSON + 1).setValue('[]'); }

  sheet.getRange(rowNum, EVAL_COLUMNS.TRANSCRIPT_FILE_ID + 1).setValue(transcriptFileId || '');
  sheet.getRange(rowNum, EVAL_COLUMNS.STATUS + 1).setValue(EVALUATION_STATUS.AI_COMPLETE);
  sheet.getRange(rowNum, EVAL_COLUMNS.UPDATED_AT + 1).setValue(now);

  console.log('評価結果保存完了: row=' + rowNum + ', AI合計=' + aiTotal);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// データ取得
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 評価結果一覧取得（ページネーション・フィルタ対応）
 * @param {Object} options - { consultant, status, limit, offset }
 * @returns {Object} { results, total }
 */
function getEvaluationResults(options) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.EVALUATION.SHEET_NAME || 'コンサルタント評価';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return { results: [], total: 0 };

  var data = sheet.getDataRange().getValues();
  var results = [];

  for (var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var consultant = row[EVAL_COLUMNS.CONSULTANT_NAME] || '';
    var status = row[EVAL_COLUMNS.STATUS] || '';

    // フィルタ
    if (options.consultant && consultant.indexOf(options.consultant) < 0) continue;
    if (options.status && status !== options.status) continue;

    results.push({
      evalId: row[EVAL_COLUMNS.EVAL_ID],
      appId: row[EVAL_COLUMNS.APP_ID],
      consultantName: consultant,
      companyName: row[EVAL_COLUMNS.COMPANY_NAME],
      evalDate: row[EVAL_COLUMNS.EVAL_DATE] instanceof Date
        ? Utilities.formatDate(row[EVAL_COLUMNS.EVAL_DATE], 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(row[EVAL_COLUMNS.EVAL_DATE] || ''),
      totalScore: row[EVAL_COLUMNS.TOTAL_SCORE] || 0,
      aiTotal: row[EVAL_COLUMNS.AI_TOTAL] || 0,
      humanTotal: row[EVAL_COLUMNS.HUMAN_TOTAL] || 0,
      categories: {
        c1: row[EVAL_COLUMNS.C1_PROBLEM] || 0,
        c2: row[EVAL_COLUMNS.C2_SOLUTION] || 0,
        c3: row[EVAL_COLUMNS.C3_COMMUNICATION] || 0,
        c4: row[EVAL_COLUMNS.C4_TIME] || 0,
        c5: row[EVAL_COLUMNS.C5_LOGIC] || 0,
        c6: row[EVAL_COLUMNS.C6_ETHICS] || 0
      },
      status: status
    });
  }

  var total = results.length;
  var limit = parseInt(options.limit) || 20;
  var offset = parseInt(options.offset) || 0;
  results = results.slice(offset, offset + limit);

  return { results: results, total: total };
}

/**
 * 評価詳細取得（JSON展開付き）
 * @param {string} evalId - 評価ID
 * @returns {Object|null}
 */
function getEvaluationById(evalId) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.EVALUATION.SHEET_NAME || 'コンサルタント評価';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return null;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][EVAL_COLUMNS.EVAL_ID] === evalId) {
      var row = data[i];
      var itemScores = {};
      try { itemScores = JSON.parse(row[EVAL_COLUMNS.ITEM_SCORES_JSON] || '{}'); } catch (e) {}
      var evidence = {};
      try { evidence = JSON.parse(row[EVAL_COLUMNS.EVIDENCE_JSON] || '{}'); } catch (e) {}
      var ngWords = [];
      try { ngWords = JSON.parse(row[EVAL_COLUMNS.NG_WORDS_JSON] || '[]'); } catch (e) {}

      return {
        evalId: row[EVAL_COLUMNS.EVAL_ID],
        appId: row[EVAL_COLUMNS.APP_ID],
        consultantName: row[EVAL_COLUMNS.CONSULTANT_NAME],
        companyName: row[EVAL_COLUMNS.COMPANY_NAME],
        evalDate: row[EVAL_COLUMNS.EVAL_DATE] instanceof Date
          ? Utilities.formatDate(row[EVAL_COLUMNS.EVAL_DATE], 'Asia/Tokyo', 'yyyy/MM/dd')
          : String(row[EVAL_COLUMNS.EVAL_DATE] || ''),
        totalScore: row[EVAL_COLUMNS.TOTAL_SCORE] || 0,
        aiTotal: row[EVAL_COLUMNS.AI_TOTAL] || 0,
        humanTotal: row[EVAL_COLUMNS.HUMAN_TOTAL] || 0,
        categories: {
          c1: row[EVAL_COLUMNS.C1_PROBLEM] || 0,
          c2: row[EVAL_COLUMNS.C2_SOLUTION] || 0,
          c3: row[EVAL_COLUMNS.C3_COMMUNICATION] || 0,
          c4: row[EVAL_COLUMNS.C4_TIME] || 0,
          c5: row[EVAL_COLUMNS.C5_LOGIC] || 0,
          c6: row[EVAL_COLUMNS.C6_ETHICS] || 0
        },
        humanScores: {
          h1: row[EVAL_COLUMNS.H1_ATMOSPHERE] || 0,
          h2: row[EVAL_COLUMNS.H2_ATTITUDE] || 0,
          h3: row[EVAL_COLUMNS.H3_APPEARANCE] || 0
        },
        itemScores: itemScores,
        evidence: evidence,
        ngWords: ngWords,
        transcriptFileId: row[EVAL_COLUMNS.TRANSCRIPT_FILE_ID] || '',
        reportFileId: row[EVAL_COLUMNS.REPORT_FILE_ID] || '',
        status: row[EVAL_COLUMNS.STATUS] || '',
        humanEvaluator: row[EVAL_COLUMNS.HUMAN_EVALUATOR] || '',
        createdAt: row[EVAL_COLUMNS.CREATED_AT] instanceof Date
          ? Utilities.formatDate(row[EVAL_COLUMNS.CREATED_AT], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
          : '',
        updatedAt: row[EVAL_COLUMNS.UPDATED_AT] instanceof Date
          ? Utilities.formatDate(row[EVAL_COLUMNS.UPDATED_AT], 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
          : ''
      };
    }
  }
  return null;
}

/**
 * コンサルタント別履歴取得（成長追跡用）
 * @param {string} consultantName
 * @returns {Array}
 */
function getConsultantHistory(consultantName) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.EVALUATION.SHEET_NAME || 'コンサルタント評価';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return [];

  var data = sheet.getDataRange().getValues();
  var history = [];

  for (var i = 1; i < data.length; i++) {
    var name = data[i][EVAL_COLUMNS.CONSULTANT_NAME] || '';
    if (name !== consultantName) continue;
    var status = data[i][EVAL_COLUMNS.STATUS] || '';
    if (status === EVALUATION_STATUS.ERROR || status === EVALUATION_STATUS.AI_RUNNING) continue;

    history.push({
      evalId: data[i][EVAL_COLUMNS.EVAL_ID],
      evalDate: data[i][EVAL_COLUMNS.EVAL_DATE] instanceof Date
        ? Utilities.formatDate(data[i][EVAL_COLUMNS.EVAL_DATE], 'Asia/Tokyo', 'yyyy/MM/dd')
        : String(data[i][EVAL_COLUMNS.EVAL_DATE] || ''),
      companyName: data[i][EVAL_COLUMNS.COMPANY_NAME] || '',
      totalScore: data[i][EVAL_COLUMNS.TOTAL_SCORE] || 0,
      aiTotal: data[i][EVAL_COLUMNS.AI_TOTAL] || 0,
      humanTotal: data[i][EVAL_COLUMNS.HUMAN_TOTAL] || 0,
      categories: {
        c1: data[i][EVAL_COLUMNS.C1_PROBLEM] || 0,
        c2: data[i][EVAL_COLUMNS.C2_SOLUTION] || 0,
        c3: data[i][EVAL_COLUMNS.C3_COMMUNICATION] || 0,
        c4: data[i][EVAL_COLUMNS.C4_TIME] || 0,
        c5: data[i][EVAL_COLUMNS.C5_LOGIC] || 0,
        c6: data[i][EVAL_COLUMNS.C6_ETHICS] || 0
      },
      status: status
    });
  }

  return history;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 人間採点
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 人間採点登録 + 総合スコア再計算
 * @param {string} evalId - 評価ID
 * @param {number} h1 - 雰囲気 (0-4)
 * @param {number} h2 - 態度変化 (0-3)
 * @param {number} h3 - 身だしなみ (0-3)
 * @param {string} evaluatorName - 評価者名
 * @returns {Object}
 */
function updateHumanScores(evalId, h1, h2, h3, evaluatorName) {
  if (!evalId) return { success: false, message: '評価IDが必要です' };

  // バリデーション
  h1 = Math.max(0, Math.min(4, parseInt(h1) || 0));
  h2 = Math.max(0, Math.min(3, parseInt(h2) || 0));
  h3 = Math.max(0, Math.min(3, parseInt(h3) || 0));

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.EVALUATION.SHEET_NAME || 'コンサルタント評価';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: '評価シートが見つかりません' };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][EVAL_COLUMNS.EVAL_ID] === evalId) {
      var rowNum = i + 1;
      var status = data[i][EVAL_COLUMNS.STATUS] || '';

      // AI完了済みのみ人間採点可能
      if (status !== EVALUATION_STATUS.AI_COMPLETE && status !== EVALUATION_STATUS.HUMAN_PENDING && status !== EVALUATION_STATUS.COMPLETE) {
        return { success: false, message: 'AI評価が完了していません（ステータス: ' + status + '）' };
      }

      var humanTotal = h1 + h2 + h3;
      var aiTotal = data[i][EVAL_COLUMNS.AI_TOTAL] || 0;
      var totalScore = aiTotal + humanTotal;

      sheet.getRange(rowNum, EVAL_COLUMNS.H1_ATMOSPHERE + 1).setValue(h1);
      sheet.getRange(rowNum, EVAL_COLUMNS.H2_ATTITUDE + 1).setValue(h2);
      sheet.getRange(rowNum, EVAL_COLUMNS.H3_APPEARANCE + 1).setValue(h3);
      sheet.getRange(rowNum, EVAL_COLUMNS.HUMAN_TOTAL + 1).setValue(humanTotal);
      sheet.getRange(rowNum, EVAL_COLUMNS.TOTAL_SCORE + 1).setValue(totalScore);
      sheet.getRange(rowNum, EVAL_COLUMNS.STATUS + 1).setValue(EVALUATION_STATUS.COMPLETE);
      sheet.getRange(rowNum, EVAL_COLUMNS.HUMAN_EVALUATOR + 1).setValue(evaluatorName || '');
      sheet.getRange(rowNum, EVAL_COLUMNS.UPDATED_AT + 1).setValue(new Date());

      console.log('人間採点登録: ' + evalId + ', H1=' + h1 + ' H2=' + h2 + ' H3=' + h3 + ' 合計=' + totalScore);

      return {
        success: true,
        evalId: evalId,
        humanTotal: humanTotal,
        totalScore: totalScore,
        message: '人間採点を登録しました（総合スコア: ' + totalScore + '/100点）'
      };
    }
  }

  return { success: false, message: '評価ID ' + evalId + ' が見つかりません' };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 統計
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 統計サマリー
 */
function getEvaluationStats() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.EVALUATION.SHEET_NAME || 'コンサルタント評価';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { success: true, totalEvaluations: 0, averageScore: 0, categoryAverages: {} };
  }

  var data = sheet.getDataRange().getValues();
  var count = 0;
  var totalScore = 0;
  var catTotals = { c1: 0, c2: 0, c3: 0, c4: 0, c5: 0, c6: 0 };
  var consultantScores = {};

  for (var i = 1; i < data.length; i++) {
    var status = data[i][EVAL_COLUMNS.STATUS] || '';
    if (status === EVALUATION_STATUS.ERROR || status === EVALUATION_STATUS.AI_RUNNING || status === EVALUATION_STATUS.NONE) continue;

    count++;
    totalScore += data[i][EVAL_COLUMNS.TOTAL_SCORE] || 0;
    catTotals.c1 += data[i][EVAL_COLUMNS.C1_PROBLEM] || 0;
    catTotals.c2 += data[i][EVAL_COLUMNS.C2_SOLUTION] || 0;
    catTotals.c3 += data[i][EVAL_COLUMNS.C3_COMMUNICATION] || 0;
    catTotals.c4 += data[i][EVAL_COLUMNS.C4_TIME] || 0;
    catTotals.c5 += data[i][EVAL_COLUMNS.C5_LOGIC] || 0;
    catTotals.c6 += data[i][EVAL_COLUMNS.C6_ETHICS] || 0;

    var consultant = data[i][EVAL_COLUMNS.CONSULTANT_NAME] || '不明';
    if (!consultantScores[consultant]) consultantScores[consultant] = { count: 0, total: 0 };
    consultantScores[consultant].count++;
    consultantScores[consultant].total += data[i][EVAL_COLUMNS.TOTAL_SCORE] || 0;
  }

  var categoryAverages = {};
  for (var key in catTotals) {
    categoryAverages[key] = count > 0 ? Math.round(catTotals[key] / count * 10) / 10 : 0;
  }

  var consultantAvg = {};
  for (var name in consultantScores) {
    consultantAvg[name] = {
      count: consultantScores[name].count,
      averageScore: Math.round(consultantScores[name].total / consultantScores[name].count)
    };
  }

  return {
    success: true,
    totalEvaluations: count,
    averageScore: count > 0 ? Math.round(totalScore / count) : 0,
    categoryAverages: categoryAverages,
    consultantBreakdown: consultantAvg
  };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 手動操作
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 手動で評価を実行（予約管理シートの行を指定）
 * @param {number} rowIndex - 行番号
 * @returns {Object}
 */
function manualRunEvaluation(rowIndex) {
  try {
    var rowData = getRowData(rowIndex);
    if (!rowData.id) {
      return { success: false, message: '行 ' + rowIndex + ' にデータがありません' };
    }

    // 文字起こしファイルIDを取得
    var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_NAME);
    var transcriptFileId = sheet.getRange(rowIndex, COLUMNS.TRANSCRIPT_FILE_ID + 1).getValue();

    if (!transcriptFileId) {
      return { success: false, message: '文字起こしファイルが未設定です（AE列）。評価には文字起こしが必要です。' };
    }

    return runConsultationEvaluation(transcriptFileId, rowIndex);

  } catch (e) {
    console.error('手動評価実行エラー:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * 評価設定確認
 */
function checkEvaluationSetup() {
  var props = PropertiesService.getScriptProperties();
  return {
    success: true,
    enabled: !!(CONFIG.EVALUATION && CONFIG.EVALUATION.ENABLED),
    cfUrl: !!(props.getProperty('EVALUATION_CF_URL') || (CONFIG.EVALUATION && CONFIG.EVALUATION.CLOUD_FUNCTION_URL)),
    cfSecret: !!(props.getProperty('EVALUATION_CF_SECRET') || (CONFIG.EVALUATION && CONFIG.EVALUATION.CLOUD_FUNCTION_SECRET)),
    sheetName: (CONFIG.EVALUATION && CONFIG.EVALUATION.SHEET_NAME) || 'コンサルタント評価'
  };
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// ヘルパー
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 評価シートを取得（なければ作成）
 */
function getOrCreateEvaluationSheet_() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheetName = CONFIG.EVALUATION.SHEET_NAME || 'コンサルタント評価';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    setupEvaluationSheet();
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}
