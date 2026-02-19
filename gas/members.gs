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
      name: row[MEMBER_COLUMNS.NAME],
      term: row[MEMBER_COLUMNS.TERM],
      cert: row[MEMBER_COLUMNS.CERT],
      type: row[MEMBER_COLUMNS.TYPE],
      email: row[MEMBER_COLUMNS.EMAIL],
      phone: row[MEMBER_COLUMNS.PHONE],
      lineId: row[MEMBER_COLUMNS.LINE_ID],
      notes: row[MEMBER_COLUMNS.NOTES],
      specialties: row[MEMBER_COLUMNS.SPECIALTIES] ? row[MEMBER_COLUMNS.SPECIALTIES].toString() : '',
      themes: row[MEMBER_COLUMNS.THEMES] ? row[MEMBER_COLUMNS.THEMES].toString() : ''
    });
  }

  return members;
}

/**
 * 日程関連に参加するメンバーのみ取得（顧問を除外）
 * @returns {Array<Object>} スケジュール対象メンバーの配列
 */
function getScheduleMembers() {
  const members = getAllMembers();
  return members.filter(m => m.type !== '顧問');
}

/**
 * 名前からメンバー情報を取得
 * @param {string} name - メンバー名
 * @returns {Object|null} メンバー情報
 */
function getMemberByName(name) {
  const members = getAllMembers();
  return members.find(m => m.name === name) || null;
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
    return allMembers.find(m => m.name === name) || { name: name, term: '', cert: '', type: '', email: '', phone: '', lineId: '', notes: '', specialties: '', themes: '' };
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
 * 予約可能判定
 * - 配置点数 >= 4（診断士2名以上）→ ○
 * - 配置点数 >= 2 かつ 特別対応フラグ = TRUE → ○
 * - それ以外 → ×
 * @param {number} score - 配置点数
 * @param {boolean} specialFlag - 特別対応フラグ（プロコン・指導教員同席）
 * @returns {string} '○' or '×'
 */
function getBookableStatus(score, specialFlag) {
  if (score >= 4) {
    return '○';
  }
  if (score >= 2 && specialFlag === true) {
    return '○';
  }
  return '×';
}

/**
 * 担当者のLINE IDを取得
 * @param {string} staffName - 担当者名
 * @returns {string|null} LINE ID
 */
function getStaffLineId(staffName) {
  const member = getMemberByName(staffName);
  if (member && member.lineId) {
    return member.lineId;
  }
  return null;
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
  const headers = ['氏名', '期', '資格', '区分', 'メール', '電話番号', 'LINE ID', '備考', '得意業種', '得意テーマ'];
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
  sheet.setColumnWidth(7, 200);  // LINE ID
  sheet.setColumnWidth(8, 200);  // 備考
  sheet.setColumnWidth(9, 200);  // 得意業種
  sheet.setColumnWidth(10, 200); // 得意テーマ

  // 1行目を固定
  sheet.setFrozenRows(1);

  // サンプルデータを投入
  const sampleData = [
    ['杉山 宏和', '1期', '診断士', '正会員', 'hirokazusugisugi@gmail.com', '', '', '', '製造業,小売業', '経営戦略,マーケティング'],
    ['川崎 真規', '1期', '診断士', '正会員', 'stevenm.kawasaki@gmail.com', '', '', 'TA', 'IT,サービス業', '財務,IT活用'],
    ['原 真人', '1期', '診断士', '正会員', 'm.hara.2006@gmail.com', '', '', '', '製造業,建設業', '生産管理,品質管理'],
    ['小椋 孝博', '2期', '診断士', '正会員', 'takahiro09.03.21@gmail.com', '', '', '', '小売業,飲食業', '人事,組織'],
    ['秋月 仁志', '2期', '診断士', '正会員', 'akizukihitoshi@gmail.com', '', '', '', 'サービス業,IT', 'マーケティング,新規事業'],
    ['谷村 真里', '0期', '', '顧問', 'mari_tanimura@k-mba.com', '', '', '', '', ''],
    ['野田 慎士', '3期', '', 'オブザーバー', 'jimi320320320@gmail.com', '', '', '', '', ''],
    ['高乘 麻美', '4期', '', 'オブザーバー', 'asami.koujou@gmail.com', '', '', '', '', ''],
    ['村本 将之', '4期', '', 'オブザーバー', 'kastu.mura3.teru@gmail.com', '', '', '', '', ''],
    ['織田 美智子', '4期', '', 'オブザーバー', 'amdt.ked@gmail.com', '', '', '', '', '']
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
