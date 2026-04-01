/**
 * メンバーマスタ管理
 * メンバー情報の参照、配置点数の計算
 */

/**
 * メンバーマスタから全メンバーを取得
 * @returns {Array<Object>} メンバー情報の配列
 */
function getAllMembers() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME);

  if (!sheet) {
    console.log('メンバーマスタシートが見つかりません');
    return [];
  }

  const data = sheet.getDataRange().getValues();
  const members = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[MEMBER_COLUMNS.NAME]) continue;

    members.push({
      name: (row[MEMBER_COLUMNS.NAME] || '').toString().trim(),
      term: row[MEMBER_COLUMNS.TERM],
      cert: row[MEMBER_COLUMNS.CERT],
      type: row[MEMBER_COLUMNS.TYPE],
      email: row[MEMBER_COLUMNS.EMAIL],
      phone: row[MEMBER_COLUMNS.PHONE],
      notes: row[MEMBER_COLUMNS.NOTES],
      specialties: row[MEMBER_COLUMNS.SPECIALTIES] ? row[MEMBER_COLUMNS.SPECIALTIES].toString() : '',
      themes: row[MEMBER_COLUMNS.THEMES] ? row[MEMBER_COLUMNS.THEMES].toString() : '',
      active: row[MEMBER_COLUMNS.ACTIVE] !== false,
      titles: row[MEMBER_COLUMNS.TITLES] ? row[MEMBER_COLUMNS.TITLES].toString() : ''
    });
  }

  return members;
}

/**
 * 日程関連に参加するメンバーのみ取得（顧問を除外、システム参加=FALSEを除外）
 * @returns {Array<Object>} スケジュール対象メンバーの配列
 */
function getScheduleMembers() {
  const members = getAllMembers();
  return members.filter(m => m.type !== '顧問' && m.active !== false);
}

/**
 * 名前からメンバー情報を取得
 * @param {string} name - メンバー名
 * @returns {Object|null} メンバー情報
 */
function getMemberByName(name) {
  if (!name) return null;
  const trimmed = name.toString().trim();
  const members = getAllMembers();
  return members.find(m => m.name === trimmed) || null;
}

/**
 * メールアドレスからメンバー情報を取得
 * @param {string} email - メールアドレス
 * @returns {Object|null} メンバー情報
 */
function getMemberByEmail(email) {
  if (!email) return null;
  const members = getAllMembers();
  return members.find(m => m.email && m.email.toLowerCase() === email.toLowerCase()) || null;
}

/**
 * 複数メンバー名からメンバー情報を一括取得
 * @param {string} memberNames - カンマ区切りのメンバー名
 * @returns {Array<Object>} メンバー情報の配列
 */
function getMembersByNames(memberNames) {
  if (!memberNames) return [];

  const names = memberNames.split(',').map(n => n.trim()).filter(n => n);
  const allMembers = getAllMembers();

  return names.map(name => {
    return allMembers.find(m => m.name === name) || { name: name, term: '', cert: '', type: '', email: '', phone: '', notes: '', specialties: '', themes: '', active: true, titles: '' };
  });
}

/**
 * メンバーの配置点数を計算
 * 1期・2期の診断士 = 2点/人
 * 3期・4期オブザーバー = 1点/人
 * @param {string} memberNames - カンマ区切りのメンバー名
 * @returns {number} 配置点数
 */
function calculateStaffScore(memberNames) {
  if (!memberNames) return 0;

  const members = getMembersByNames(memberNames);
  let score = 0;

  members.forEach(member => {
    const term = member.term ? member.term.toString() : '';

    if (term === '1期' || term === '2期') {
      // 1期・2期の診断士 = 2点
      score += 2;
    } else if (term === '3期' || term === '4期') {
      // 3期・4期オブザーバー = 1点
      score += 1;
    }
  });

  return score;
}

/**
 * 診断士（1期・2期）の人数を返す
 * @param {string} memberNames - カンマ区切りのメンバー名
 * @returns {number} 診断士の人数
 */
function countDiagnosticians(memberNames) {
  if (!memberNames) return 0;
  const members = getMembersByNames(memberNames);
  return members.filter(function(m) {
    var term = m.term ? m.term.toString() : '';
    return term === '1期' || term === '2期';
  }).length;
}

/**
 * 予約可能判定
 * - 診断士（1期・2期）が1名以上 かつ 配置点数 >= 4 → ○
 * - 診断士（1期・2期）が1名以上 かつ 配置点数 >= 2 かつ 特別対応フラグ = TRUE → ○
 * - それ以外 → ×
 * @param {number} score - 配置点数
 * @param {boolean} specialFlag - 特別対応フラグ（プロコン・指導教員同席）
 * @param {number} diagCount - 診断士（1期・2期）の人数
 * @returns {string} '○' or '×'
 */
function getBookableStatus(score, specialFlag, diagCount) {
  if (typeof diagCount === 'undefined') diagCount = 0;
  if (diagCount < 1) return '×';
  if (score >= 4) {
    return '○';
  }
  if (score >= 2 && specialFlag === true) {
    return '○';
  }
  return '×';
}

/**
 * 担当者のメールアドレスを取得
 * @param {string} staffName - 担当者名
 * @returns {string|null} メールアドレス
 */
function getStaffEmail(staffName) {
  const member = getMemberByName(staffName);
  if (member && member.email) {
    return member.email;
  }
  return null;
}

/**
 * メンバーマスタシートのセットアップ
 */
function setupMemberSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let sheet = ss.getSheetByName(CONFIG.MEMBER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.MEMBER_SHEET_NAME);
  }

  // ヘッダー設定
  const headers = ['氏名', '期', '資格', '区分', 'メール', '電話番号', '', '備考', '得意業種', '得意テーマ', 'システム参加', '肩書き'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // ヘッダー行の書式
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f2350');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');

  // 列幅調整
  sheet.setColumnWidth(1, 100);  // 氏名
  sheet.setColumnWidth(2, 60);   // 期
  sheet.setColumnWidth(3, 80);   // 資格
  sheet.setColumnWidth(4, 80);   // 区分
  sheet.setColumnWidth(5, 250);  // メール
  sheet.setColumnWidth(6, 120);  // 電話番号
  sheet.setColumnWidth(7, 200);  // （未使用）
  sheet.setColumnWidth(8, 200);  // 備考
  sheet.setColumnWidth(9, 200);  // 得意業種
  sheet.setColumnWidth(10, 200); // 得意テーマ
  sheet.setColumnWidth(11, 100); // システム参加
  sheet.setColumnWidth(12, 200); // 肩書き

  // 1行目を固定
  sheet.setFrozenRows(1);

  // サンプルデータを投入
  const sampleData = [
    ['杉山 宏和', '1期', '中小企業診断士,MBA,FP1級/CFP,販売士1級,G検定/E資格', '正会員', 'hirokazusugisugi@gmail.com', '', '', '', '製造業,小売業', '経営戦略,マーケティング', true, ''],
    ['川崎 真規', '1期', '中小企業診断士,MBA', '正会員', 'stevenm.kawasaki@gmail.com', '', '', 'TA', 'IT,サービス業', '財務,IT活用', true, '関西学院大学 経営戦略科 TA,大阪商工会議所 アドバイザー'],
    ['原 真人', '1期', '中小企業診断士,MBA', '正会員', 'm.hara.2006@gmail.com', '', '', '', '', '財務,経理,経営企画', true, ''],
    ['小椋 孝博', '2期', '中小企業診断士,MBA', '正会員', 'takahiro09.03.21@gmail.com', '', '', '', '製造業', '物流,営業,DX', true, ''],
    ['秋月 仁志', '2期', '中小企業診断士,MBA,ITストラテジスト', '正会員', 'akizukihitoshi@gmail.com', '', '', '', '', '経営企画,DX,生成AI,業務改革', true, ''],
    ['谷村 真里', '0期', '', '顧問', 'mari_tanimura@k-mba.com', '', '', '', '', '', true, ''],
    ['野田 慎士', '3期', '', 'オブザーバー', 'jimi320320320@gmail.com', '', '', '', '', '', true, ''],
    ['高山 佳樹', '3期', '', 'オブザーバー', 'gaoshanjiashu@gmail.com', '', '', 'Zoom参加のみ', '', '', true, ''],
    ['高乘 麻美', '4期', '', 'オブザーバー', 'asami.koujou@gmail.com', '', '', '', '', '', true, ''],
    ['村本 将之', '4期', '', 'オブザーバー', 'katsu.mura3.teru@gmail.com', '', '', '', '', '', true, ''],
    ['織田 美智子', '4期', '', 'オブザーバー', 'amdt.ked@gmail.com', '', '', '', '', '', true, '']
  ];

  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);

  console.log('メンバーマスタシートのセットアップが完了しました');
}

/**
 * 参加メンバーから1期/2期の診断士（リーダー候補）を抽出
 * @param {string} memberNames - カンマ区切りのメンバー名
 * @returns {Array<Object>} リーダー候補メンバーの配列
 */
function getLeaderCandidates(memberNames) {
  if (!memberNames) return [];

  const members = getMembersByNames(memberNames);
  return members.filter(function(m) {
    var term = m.term ? m.term.toString() : '';
    return term === '1期' || term === '2期';
  });
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 担当者自動選定
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 相談テーマと得意テーマのマッチング判定
 * @param {string} memberThemes - メンバーの得意テーマ（カンマ区切り）
 * @param {string} consultTheme - 相談テーマ
 * @returns {boolean} マッチするかどうか
 */
function matchTheme(memberThemes, consultTheme) {
  if (!memberThemes || !consultTheme) return false;
  var themes = memberThemes.split(',').map(function(t) { return t.trim().toLowerCase(); }).filter(Boolean);
  var target = consultTheme.toString().trim().toLowerCase();
  if (!target) return false;
  for (var i = 0; i < themes.length; i++) {
    if (target.indexOf(themes[i]) >= 0 || themes[i].indexOf(target) >= 0) {
      return true;
    }
  }
  return false;
}

/**
 * 配列をランダムにシャッフル（Fisher-Yates）
 * @param {Array} arr - シャッフルする配列
 * @returns {Array} シャッフルされた新しい配列
 */
function shuffleArray(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = a[i];
    a[i] = a[j];
    a[j] = tmp;
  }
  return a;
}

/**
 * メンバーリストの配置点数を計算
 * @param {Array<Object>} members - メンバー情報の配列
 * @returns {number} 配置点数
 */
function calculateMemberScore(members) {
  var score = 0;
  members.forEach(function(m) {
    var term = m.term ? m.term.toString() : '';
    if (term === '1期' || term === '2期') {
      score += 2;
    } else if (term === '3期' || term === '4期') {
      score += 1;
    }
  });
  return score;
}

/**
 * メンバーリスト中の診断士（1期/2期）の人数
 * @param {Array<Object>} members - メンバー情報の配列
 * @returns {number}
 */
function countShindanshi(members) {
  return members.filter(function(m) {
    var term = m.term ? m.term.toString() : '';
    return term === '1期' || term === '2期';
  }).length;
}

/**
 * 確定時に全アクティブメンバーから担当者を自動選定
 *
 * ロジック:
 * 1. 全アクティブメンバー（顧問除外）を候補プールとする
 * 2. 診断士（1期/2期）からテーママッチ優先で1名以上選出
 * 3. 残り枠をランダムで充填し合計SELECT名
 * 4. SELECT名から FINAL名を確定、残りを予備
 *
 * @param {string} consultTheme - 相談テーマ
 * @param {boolean} specialFlag - 特別対応フラグ
 * @returns {Object} { success, confirmed: [name], reserve: [name], score, message }
 */
function selectStaffMembers(consultTheme, specialFlag) {
  var limits = CONFIG.STAFF_LIMITS;
  var selectCount = limits.SELECT;   // 4
  var finalCount = limits.FINAL;     // 3
  var minSpecial = limits.MIN_SPECIAL; // 2
  var reqShindanshi = limits.REQUIRE_SHINDANSHI; // 1

  // ── 候補プールの構築 ──
  var allMembers = getAllMembers();
  var pool = allMembers.filter(function(m) {
    return m.active !== false && m.type !== '顧問';
  });

  var shindanshiPool = pool.filter(function(m) {
    var term = m.term ? m.term.toString() : '';
    return term === '1期' || term === '2期';
  });
  var obPool = pool.filter(function(m) {
    var term = m.term ? m.term.toString() : '';
    return term !== '1期' && term !== '2期';
  });

  // 診断士が最低人数に満たない場合
  if (shindanshiPool.length < reqShindanshi) {
    return { success: false, confirmed: [], reserve: [], score: 0,
      message: '診断士（1期/2期）が' + reqShindanshi + '名以上必要ですが、有効な診断士が' + shindanshiPool.length + '名しかいません。' };
  }

  // ── ステップ2: 診断士の選定（テーママッチ優先） ──
  var matchedShindanshi = [];
  var unmatchedShindanshi = [];
  shindanshiPool.forEach(function(m) {
    if (matchTheme(m.themes, consultTheme)) {
      matchedShindanshi.push(m);
    } else {
      unmatchedShindanshi.push(m);
    }
  });

  matchedShindanshi = shuffleArray(matchedShindanshi);
  unmatchedShindanshi = shuffleArray(unmatchedShindanshi);

  // 診断士選出: テーマ一致を優先し最低1名
  var selected = [];
  if (matchedShindanshi.length > 0) {
    selected.push(matchedShindanshi[0]);
  } else {
    selected.push(unmatchedShindanshi[0]);
    unmatchedShindanshi = unmatchedShindanshi.slice(1);
  }

  // ── ステップ3: 残り枠の充填 ──
  // 残り候補（選出済みを除く）をシャッフル
  var selectedNames = {};
  selected.forEach(function(m) { selectedNames[m.name] = true; });

  var remaining = [];
  // 未選出の診断士（テーマ一致）
  matchedShindanshi.forEach(function(m) {
    if (!selectedNames[m.name]) remaining.push(m);
  });
  // 未選出の診断士（テーマ不一致）
  unmatchedShindanshi.forEach(function(m) {
    if (!selectedNames[m.name]) remaining.push(m);
  });
  // OB（シャッフル済み）
  var shuffledOb = shuffleArray(obPool);
  remaining = remaining.concat(shuffledOb);

  // selectCount名まで追加
  for (var i = 0; i < remaining.length && selected.length < selectCount; i++) {
    selected.push(remaining[i]);
  }

  // 候補が足りない場合はそのまま進む

  // ── 点数チェック＋調整 ──
  var score = calculateMemberScore(selected);
  var minScore = specialFlag ? 2 : 4;

  // 点数不足時: OBを診断士に入れ替え
  if (score < minScore) {
    var unselectedShindanshi = [];
    var allSelectedNames = {};
    selected.forEach(function(m) { allSelectedNames[m.name] = true; });
    shindanshiPool.forEach(function(m) {
      if (!allSelectedNames[m.name]) unselectedShindanshi.push(m);
    });

    // OBを後ろから入れ替え
    for (var s = selected.length - 1; s >= 0 && score < minScore && unselectedShindanshi.length > 0; s--) {
      var t = selected[s].term ? selected[s].term.toString() : '';
      if (t !== '1期' && t !== '2期') {
        var replacement = unselectedShindanshi.shift();
        score = score - 1 + 2; // OB(1pt) → 診断士(2pt)
        selected[s] = replacement;
      }
    }
  }

  // ── ステップ4: 確定 + 予備の振り分け ──
  var confirmed = [];
  var reserve = [];
  var minFinal = specialFlag ? minSpecial : finalCount;

  if (selected.length <= minFinal) {
    // 人数が確定人数以下ならそのまま全員確定
    confirmed = selected;
  } else {
    // 4名から1名を予備に（OB優先、条件を崩さない範囲で）
    var bestReserveIdx = -1;
    for (var r = selected.length - 1; r >= 0; r--) {
      var without = selected.filter(function(_, idx) { return idx !== r; });
      var wScore = calculateMemberScore(without);
      var wShindanshi = countShindanshi(without);
      if (wShindanshi >= reqShindanshi && wScore >= minScore && without.length >= minFinal) {
        var rTerm = selected[r].term ? selected[r].term.toString() : '';
        // OBを優先的に予備にする
        if (rTerm !== '1期' && rTerm !== '2期') {
          bestReserveIdx = r;
          break;
        }
        // 診断士でも条件を満たすなら候補に（OBが見つからなかった場合のフォールバック）
        if (bestReserveIdx < 0) bestReserveIdx = r;
      }
    }

    if (bestReserveIdx >= 0) {
      reserve.push(selected[bestReserveIdx]);
      confirmed = selected.filter(function(_, idx) { return idx !== bestReserveIdx; });
    } else {
      // 誰も予備にできない場合は全員確定
      confirmed = selected;
    }
  }

  // ── ステップ5: 最終バリデーション ──
  var finalScore = calculateMemberScore(confirmed);
  var finalShindanshi = countShindanshi(confirmed);

  if (finalShindanshi < reqShindanshi) {
    return { success: false, confirmed: [], reserve: [], score: finalScore,
      message: '確定メンバーに診断士（1期/2期）が含まれていません。メンバーマスタを確認してください。' };
  }
  if (finalScore < minScore) {
    return { success: false, confirmed: [], reserve: [], score: finalScore,
      message: '配置点数が' + minScore + '点以上必要ですが、' + finalScore + '点です。' };
  }

  var confirmedNames = confirmed.map(function(m) { return m.name; });
  var reserveNames = reserve.map(function(m) { return m.name; });

  console.log('担当者自動選定完了: 確定=' + confirmedNames.join(',') + ' 予備=' + reserveNames.join(',') + ' スコア=' + finalScore);

  return {
    success: true,
    confirmed: confirmedNames,
    reserve: reserveNames,
    score: finalScore,
    message: '担当者選定完了（確定' + confirmedNames.length + '名、予備' + reserveNames.length + '名、' + finalScore + '点）'
  };
}
