/**
 * リーダー選定システム
 * 相談完了後、当日参加メンバーから最適なリーダーを自動選定
 */

/**
 * メインの選定関数
 * @param {Object} data - 予約管理シートの行データ（getRowData形式）
 * @param {string} memberNames - カンマ区切りの参加メンバー名
 * @returns {Object} { leaderName, score, reason } or null
 */
function selectLeader(data, memberNames) {
  var candidates = getLeaderCandidates(memberNames);

  if (candidates.length === 0) {
    console.log('リーダー候補なし（1期/2期の診断士が参加していません）');
    return null;
  }

  // 候補が1名のみの場合
  if (candidates.length === 1) {
    var solo = candidates[0];
    var soloScore = calculateExpertiseScore(solo, data.industry, data.theme, data.content);
    return {
      leaderName: solo.name,
      score: soloScore,
      reason: '候補1名のため自動選定'
    };
  }

  // 複数候補: スコアリング
  var scored = candidates.map(function(member) {
    var expertiseScore = calculateExpertiseScore(member, data.industry, data.theme, data.content);
    var fairnessWeight = calculateFairnessWeight(member.name);
    var totalScore = expertiseScore + fairnessWeight;

    return {
      member: member,
      expertiseScore: expertiseScore,
      fairnessWeight: fairnessWeight,
      totalScore: totalScore
    };
  });

  // スコア降順でソート
  scored.sort(function(a, b) { return b.totalScore - a.totalScore; });

  // 最高スコアの候補を取得
  var maxScore = scored[0].totalScore;
  var topCandidates = scored.filter(function(s) { return s.totalScore === maxScore; });

  // 同点の場合はランダム選択
  var selected;
  if (topCandidates.length > 1) {
    var randomIndex = Math.floor(Math.random() * topCandidates.length);
    selected = topCandidates[randomIndex];
  } else {
    selected = topCandidates[0];
  }

  // 選定理由の構築
  var reasons = [];
  if (selected.expertiseScore > 0) {
    reasons.push('専門性スコア:' + selected.expertiseScore);
  }
  if (selected.fairnessWeight < 0) {
    reasons.push('履歴ペナルティ:' + selected.fairnessWeight);
  }
  if (topCandidates.length > 1) {
    reasons.push('同点' + topCandidates.length + '名からランダム選出');
  }
  var reason = reasons.length > 0 ? reasons.join(', ') : 'スコアリングによる選定';

  return {
    leaderName: selected.member.name,
    score: selected.totalScore,
    reason: reason
  };
}

/**
 * 専門性スコア計算
 * @param {Object} member - メンバー情報（specialties, themes含む）
 * @param {string} industry - 相談者の業種
 * @param {string} theme - 相談テーマ
 * @param {string} content - 相談内容
 * @returns {number} 専門性スコア
 */
function calculateExpertiseScore(member, industry, theme, content) {
  var score = 0;

  // 業種マッチ (+3)
  if (industry && member.specialties) {
    var specialties = member.specialties.split(',').map(function(s) { return s.trim(); });
    for (var i = 0; i < specialties.length; i++) {
      if (specialties[i] && industry.indexOf(specialties[i]) >= 0) {
        score += 3;
        break;
      }
    }
  }

  // テーママッチ (+3)
  if (theme && member.themes) {
    var themes = member.themes.split(',').map(function(s) { return s.trim(); });
    for (var j = 0; j < themes.length; j++) {
      if (themes[j] && theme.indexOf(themes[j]) >= 0) {
        score += 3;
        break;
      }
    }
  }

  // 相談内容キーワードマッチ (+1)
  if (content && member.themes) {
    var themeKeywords = member.themes.split(',').map(function(s) { return s.trim(); });
    for (var k = 0; k < themeKeywords.length; k++) {
      if (themeKeywords[k] && content.indexOf(themeKeywords[k]) >= 0) {
        score += 1;
        break;
      }
    }
  }

  return score;
}

/**
 * 公平性ペナルティ計算（直近5件のリーダー履歴を参照）
 * @param {string} memberName - メンバー名
 * @returns {number} ペナルティ（負の値）
 */
function calculateFairnessWeight(memberName) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.LEADER_HISTORY_SHEET_NAME);

  if (!sheet || sheet.getLastRow() <= 1) {
    return 0;
  }

  var data = sheet.getDataRange().getValues();
  var recentCount = 0;
  var lookback = Math.min(data.length - 1, 5);

  // 直近5件を逆順で確認
  for (var i = data.length - 1; i >= Math.max(1, data.length - lookback); i--) {
    if (data[i][LEADER_HISTORY_COLUMNS.LEADER] === memberName) {
      recentCount++;
    }
  }

  return recentCount * -2;
}

/**
 * リーダー履歴シートに記録
 * @param {Object} data - 予約管理シートの行データ
 * @param {string} leaderName - 選定されたリーダー名
 * @param {string} memberNames - 参加メンバー名（カンマ区切り）
 * @param {number} score - マッチスコア
 * @param {string} reason - 選定理由
 */
function recordLeaderAssignment(data, leaderName, memberNames, score, reason) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.LEADER_HISTORY_SHEET_NAME);

  if (!sheet) {
    setupLeaderHistorySheet();
    sheet = ss.getSheetByName(CONFIG.LEADER_HISTORY_SHEET_NAME);
  }

  var row = [
    new Date(),        // A: 日付
    data.id,           // B: 申込ID
    data.company,      // C: 相談企業
    data.industry,     // D: 業種
    data.theme,        // E: テーマ
    leaderName,        // F: リーダー
    memberNames,       // G: 参加メンバー
    score,             // H: マッチスコア
    reason             // I: 選定理由
  ];

  sheet.appendRow(row);
  console.log('リーダー履歴に記録: ' + leaderName + ' (スコア: ' + score + ')');
}

/**
 * ステータス完了時の自動リーダー選定処理
 * @param {number} rowIndex - 予約管理シートの行番号（1-based）
 */
function autoSelectLeaderOnComplete(rowIndex) {
  var data = getRowData(rowIndex);

  // 既にリーダーが設定されている場合はスキップ
  if (data.leader) {
    console.log('リーダー既設定のためスキップ: ' + data.leader);
    return;
  }

  // 日程設定シートから参加メンバーを取得
  var memberNames = getParticipatingMembers(data.confirmedDate);
  if (!memberNames) {
    console.log('参加メンバーが取得できないためリーダー選定スキップ');
    return;
  }

  // リーダー選定
  var result = selectLeader(data, memberNames);
  if (!result) {
    console.log('リーダー候補なし');
    return;
  }

  // 予約管理シートのY列に記録
  var sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.SHEET_NAME);
  sheet.getRange(rowIndex, COLUMNS.LEADER + 1).setValue(result.leaderName);

  // リーダー履歴に記録
  recordLeaderAssignment(data, result.leaderName, memberNames, result.score, result.reason);

  console.log('リーダー自動選定完了: ' + result.leaderName + ' (行: ' + rowIndex + ')');
}

/**
 * 確定日時から日程設定シートの参加メンバーを取得
 * @param {string} confirmedDate - 確定日時
 * @returns {string|null} カンマ区切りのメンバー名
 */
function getParticipatingMembers(confirmedDate) {
  if (!confirmedDate) return null;

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var schedSheet = ss.getSheetByName(CONFIG.SCHEDULE_SHEET_NAME);
  if (!schedSheet) return null;

  var parsed = parseConfirmedDateTime(confirmedDate);
  if (!parsed || !parsed.date) return null;

  var data = schedSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var sDate = data[i][SCHEDULE_COLUMNS.DATE];
    var sDateStr = sDate instanceof Date
      ? Utilities.formatDate(sDate, 'Asia/Tokyo', 'yyyy/MM/dd')
      : String(sDate);

    if (sDateStr === parsed.date) {
      var members = data[i][SCHEDULE_COLUMNS.MEMBERS];
      if (members) {
        return members.toString();
      }
    }
  }

  return null;
}

/**
 * リーダー履歴シートのセットアップ
 */
function setupLeaderHistorySheet() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.LEADER_HISTORY_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.LEADER_HISTORY_SHEET_NAME);
  }

  var headers = ['日付', '申込ID', '相談企業', '業種', 'テーマ', 'リーダー', '参加メンバー', 'マッチスコア', '選定理由'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 120);  // 日付
  sheet.setColumnWidth(2, 120);  // 申込ID
  sheet.setColumnWidth(3, 150);  // 相談企業
  sheet.setColumnWidth(4, 100);  // 業種
  sheet.setColumnWidth(5, 120);  // テーマ
  sheet.setColumnWidth(6, 100);  // リーダー
  sheet.setColumnWidth(7, 250);  // 参加メンバー
  sheet.setColumnWidth(8, 100);  // マッチスコア
  sheet.setColumnWidth(9, 250);  // 選定理由

  sheet.setFrozenRows(1);

  console.log('リーダー履歴シートのセットアップが完了しました');
  return { success: true, message: 'リーダー履歴シートをセットアップしました' };
}
