/**
 * 担当者自動選定のテスト関数
 * GASエディタで手動実行してログを確認する
 */
function testSelectStaffMembers() {
  console.log('━━━━━━━━━━━━━━━━━━━━━━');
  console.log('担当者自動選定テスト開始');
  console.log('━━━━━━━━━━━━━━━━━━━━━━');

  // テスト1: テーマ=マーケティング（杉山がマッチ）
  console.log('\n■ テスト1: テーマ=マーケティング');
  var r1 = selectStaffMembers('マーケティング', false);
  console.log('結果: ' + JSON.stringify(r1));
  console.log('確定: ' + r1.confirmed.join(', '));
  console.log('予備: ' + r1.reserve.join(', '));
  console.log('スコア: ' + r1.score);
  assert_(r1.success, 'テスト1: success=true');
  assert_(r1.confirmed.length === 3, 'テスト1: 確定3名 (実際: ' + r1.confirmed.length + ')');
  assert_(r1.reserve.length === 1, 'テスト1: 予備1名 (実際: ' + r1.reserve.length + ')');
  assert_(r1.score >= 4, 'テスト1: スコア>=4 (実際: ' + r1.score + ')');

  // テスト2: テーマ=物流（小椋がマッチ）
  console.log('\n■ テスト2: テーマ=物流');
  var r2 = selectStaffMembers('物流', false);
  console.log('確定: ' + r2.confirmed.join(', '));
  console.log('予備: ' + r2.reserve.join(', '));
  assert_(r2.success, 'テスト2: success=true');
  assert_(r2.confirmed.length === 3, 'テスト2: 確定3名');
  assert_(r2.score >= 4, 'テスト2: スコア>=4');

  // テスト3: テーマ=人事労務（一致者なし→診断士ランダム）
  console.log('\n■ テスト3: テーマ=人事労務（マッチなし）');
  var r3 = selectStaffMembers('人事労務', false);
  console.log('確定: ' + r3.confirmed.join(', '));
  console.log('予備: ' + r3.reserve.join(', '));
  assert_(r3.success, 'テスト3: success=true');
  assert_(r3.confirmed.length === 3, 'テスト3: 確定3名');

  // テスト4: 特別対応（2名以上でOK）
  console.log('\n■ テスト4: 特別対応');
  var r4 = selectStaffMembers('経営戦略', true);
  console.log('確定: ' + r4.confirmed.join(', '));
  console.log('予備: ' + r4.reserve.join(', '));
  assert_(r4.success, 'テスト4: success=true');
  assert_(r4.score >= 2, 'テスト4: スコア>=2');

  // テスト5: 診断士が確定メンバーに含まれるか
  console.log('\n■ テスト5: 確定メンバーに診断士がいるか確認');
  var allM = getAllMembers();
  var shindanshiNames = allM.filter(function(m) {
    var t = m.term ? m.term.toString() : '';
    return (t === '1期' || t === '2期') && m.active !== false && m.type !== '顧問';
  }).map(function(m) { return m.name; });

  var hasShindanshi = r1.confirmed.some(function(name) {
    return shindanshiNames.indexOf(name) >= 0;
  });
  assert_(hasShindanshi, 'テスト5: 確定メンバーに診断士が含まれている');

  // テスト6: matchTheme関数の単体テスト
  console.log('\n■ テスト6: matchTheme単体テスト');
  assert_(matchTheme('経営戦略,マーケティング', 'マーケティング') === true, 'matchTheme: 完全一致');
  assert_(matchTheme('経営企画,DX,生成AI', 'DX推進') === true, 'matchTheme: 部分一致(DX⊂DX推進)');
  assert_(matchTheme('財務,IT活用', '経営戦略') === false, 'matchTheme: 不一致');
  assert_(matchTheme('', 'マーケティング') === false, 'matchTheme: 空テーマ');
  assert_(matchTheme('経営戦略', '') === false, 'matchTheme: 空相談テーマ');

  // テスト7: ランダム性の確認（10回実行して全同一でないか）
  console.log('\n■ テスト7: ランダム性確認（10回実行）');
  var results = [];
  for (var i = 0; i < 10; i++) {
    var rr = selectStaffMembers('経営戦略', false);
    results.push(rr.confirmed.join(','));
  }
  var uniqueResults = results.filter(function(v, i, self) { return self.indexOf(v) === i; });
  console.log('ユニーク結果数: ' + uniqueResults.length + '/10');
  // 全10回が同一はあり得るがメンバー数が十分なら通常は複数パターンになるはず

  console.log('\n━━━━━━━━━━━━━━━━━━━━━━');
  console.log('全テスト完了');
  console.log('━━━━━━━━━━━━━━━━━━━━━━');
}

/**
 * テスト用アサーション
 */
function assert_(condition, message) {
  if (condition) {
    console.log('  ✓ PASS: ' + message);
  } else {
    console.error('  ✗ FAIL: ' + message);
  }
}

/**
 * テーママッチング単体テスト
 */
function testMatchTheme() {
  console.log('matchTheme テスト');
  var cases = [
    { member: '経営戦略,マーケティング', consult: 'マーケティング', expect: true },
    { member: '財務,IT活用', consult: '経営戦略', expect: false },
    { member: '経営企画,DX,生成AI', consult: 'DX推進', expect: true },
    { member: '経営企画,DX,生成AI,業務改革', consult: '生成AI活用', expect: true },
    { member: '物流,営業,DX', consult: '物流改善', expect: true },
    { member: '', consult: 'テスト', expect: false },
    { member: 'テスト', consult: '', expect: false },
  ];
  cases.forEach(function(c, i) {
    var result = matchTheme(c.member, c.consult);
    var status = result === c.expect ? '✓' : '✗';
    console.log('  ' + status + ' Case' + (i+1) + ': matchTheme("' + c.member + '", "' + c.consult + '") = ' + result + ' (期待: ' + c.expect + ')');
  });
}
