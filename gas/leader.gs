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
 * @param {string} status - ステータス（'予定' or '完了'）
 */
function recordLeaderAssignment(data, leaderName, memberNames, score, reason, status) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.LEADER_HISTORY_SHEET_NAME);

  if (!sheet) {
    setupLeaderHistorySheet();
    sheet = ss.getSheetByName(CONFIG.LEADER_HISTORY_SHEET_NAME);
  }

  var row = [
    status || '予定',        // A: ステータス
    data.confirmedDate || '',// B: 相談日時
    data.id,                // C: 申込ID
    data.company,           // D: 相談企業
    leaderName,             // E: リーダー
    data.industry,          // F: 業種
    data.theme,             // G: テーマ
    memberNames,            // H: 参加メンバー
    score,                  // I: マッチスコア
    reason,                 // J: 選定理由
    new Date()              // K: 登録日
  ];

  sheet.appendRow(row);
  console.log('リーダー履歴に記録: ' + leaderName + ' (' + (status || '予定') + ')');
}

/**
 * リーダー履歴シートの既存行のステータスを更新
 * @param {string} appId - 申込ID
 * @param {string} newStatus - 新しいステータス
 * @param {string} [newLeader] - リーダー変更がある場合
 */
function updateLeaderHistoryStatus(appId, newStatus, newLeader) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.LEADER_HISTORY_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return false;

  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][LEADER_HISTORY_COLUMNS.APP_ID] === appId) {
      sheet.getRange(i + 1, LEADER_HISTORY_COLUMNS.STATUS + 1).setValue(newStatus);
      if (newLeader) {
        sheet.getRange(i + 1, LEADER_HISTORY_COLUMNS.LEADER + 1).setValue(newLeader);
      }
      console.log('リーダー履歴ステータス更新: ' + appId + ' → ' + newStatus);
      return true;
    }
  }
  return false;
}

/**
 * ステータス完了時の自動リーダー選定処理
 * @param {number} rowIndex - 予約管理シートの行番号（1-based）
 */
function autoSelectLeaderOnComplete(rowIndex) {
  var data = getRowData(rowIndex);

  // 既にリーダーが設定されている場合 → 履歴を「完了」に更新のみ
  if (data.leader) {
    var updated = updateLeaderHistoryStatus(data.id, '完了');
    if (!updated) {
      // 履歴に「予定」行がない場合（手動設定など）は新規追加
      var memberNames = getParticipatingMembers(data.confirmedDate) || '';
      recordLeaderAssignment(data, data.leader, memberNames, 0, '手動設定', '完了');
    }
    console.log('リーダー既設定 → 履歴を完了に更新: ' + data.leader);
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

  // 履歴の「予定」行を「完了」に更新、なければ新規追加
  var updated = updateLeaderHistoryStatus(data.id, '完了', result.leaderName);
  if (!updated) {
    recordLeaderAssignment(data, result.leaderName, memberNames, result.score, result.reason, '完了');
  }

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

  var headers = ['ステータス', '相談日時', '申込ID', '相談企業', 'リーダー', '業種', 'テーマ', '参加メンバー', 'マッチスコア', '選定理由', '登録日'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 80);   // ステータス
  sheet.setColumnWidth(2, 150);  // 相談日時
  sheet.setColumnWidth(3, 120);  // 申込ID
  sheet.setColumnWidth(4, 150);  // 相談企業
  sheet.setColumnWidth(5, 100);  // リーダー
  sheet.setColumnWidth(6, 100);  // 業種
  sheet.setColumnWidth(7, 120);  // テーマ
  sheet.setColumnWidth(8, 250);  // 参加メンバー
  sheet.setColumnWidth(9, 100);  // マッチスコア
  sheet.setColumnWidth(10, 250); // 選定理由
  sheet.setColumnWidth(11, 120); // 登録日

  // ステータスの条件付き書式（A列）
  var statusRange = sheet.getRange('A2:A500');
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('予定')
    .setBackground('#fff3cd')
    .setRanges([statusRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('完了')
    .setBackground('#d4edda')
    .setRanges([statusRange])
    .build());
  sheet.setConditionalFormatRules(rules);

  sheet.setFrozenRows(1);

  console.log('リーダー履歴シートのセットアップが完了しました');
  return { success: true, message: 'リーダー履歴シートをセットアップしました' };
}
