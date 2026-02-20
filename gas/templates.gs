/**
 * メールテンプレート（拡張版）
 * 同意書ページHTML、リマインド3日前/前日、担当者向けテンプレート、当日受付用を追加
 */

/**
 * 受付確認メール（申込者向け）- 同意書確認リンク付き
 */
function getConfirmationEmailBody(data, consentUrl) {
  const consentSection = consentUrl
    ? `\n【重要: 相談同意書のご確認】
以下のリンクから相談同意書の内容をご確認の上、
同意のお手続きをお願いいたします。

同意書確認ページ: ${consentUrl}\n`
    : '';

  return `${data.name} 様

この度は、関西学院大学 中小企業経営診断研究会の
無料経営相談にお申し込みいただき、誠にありがとうございます。

以下の内容で受け付けいたしました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ お申込内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
お名前：${data.name}
貴社名：${data.company}
ご連絡先：${data.email}
ご希望日時：${data.date1}${data.date2 ? '\n第二希望：' + data.date2 : ''}
相談方法：${data.method}
相談テーマ：${data.theme}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${consentSection}
【次のステップ】
1. 上記リンクから相談同意書をご確認・ご同意ください
${data.method === 'オンライン' || data.method === 'オンライン（Zoom）'
? '2. 同意完了後、予約確定メールをお送りいたします'
: '2. 同意完了後、会場を確保いたします\n3. 会場確保後、3日以内に予約確定メールをお送りいたします\n\n※会場の都合により、ご希望の日程での開催が難しい場合は\n  別日程へのご変更をお願いする場合がございます。\n  あらかじめご了承ください。'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご注意事項
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
・本メールは自動送信です
・同意書へのご同意をもって予約受付となります
・日程の変更・キャンセルは上記メールアドレスまでご連絡ください
・オンライン相談の場合、相談内容の品質向上および
  記録のため、原則としてZoom録画を実施いたします

ご不明な点がございましたら、お気軽にお問い合わせください。

【ご相談の流れ・よくある質問】
相談者マニュアル: https://hirokazusugisugi-hue.github.io/kg-consultation/docs/manual_consultee.html

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
URL: ${CONFIG.ORG.URL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}

/**
 * 当日受付用確認メール（申込者向け）
 */
function getWalkInConfirmationEmailBody(data, consentUrl) {
  return `${data.name} 様

本日は関西学院大学 中小企業経営診断研究会の
無料経営相談にお越しいただき、誠にありがとうございます。

当日受付として登録いたしました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 受付内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
お名前：${data.name}
貴社名：${data.company}
相談テーマ：${data.theme}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【重要: 相談同意書のご確認】
相談にあたり、以下のリンクから相談同意書の内容を
ご確認の上、本日中に同意のお手続きをお願いいたします。

同意書確認ページ: ${consentUrl}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}

/**
 * 担当者通知メール（企業URL追加）
 */
function getAdminNotificationBody(data) {
  const companyUrlInfo = data.companyUrl
    ? `\n【企業URL】\n${data.companyUrl}\n※事前リサーチにご活用ください（AIツール活用を推奨）\n`
    : '';

  const walkInInfo = data.walkInFlag === 'TRUE' || data.walkInFlag === true
    ? '\n【当日受付】\nこの申込は当日受付です。\n'
    : '';

  return `新規の相談申込がありました。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 申込内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
受付日時：${Utilities.formatDate(data.timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')}
${walkInInfo}
【申込者情報】
お名前：${data.name}
貴社名：${data.company}
役職：${data.position}
業種：${data.industry}
メール：${data.email}
電話：${data.phone}
${companyUrlInfo}
【相談内容】
テーマ：${data.theme}
詳細：
${data.content || '（記載なし）'}

【希望日時】
第一希望：${data.date1}
第二希望：${data.date2 || '（なし）'}

【相談方法】
${data.method}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

スプレッドシートで担当者をアサインしてください。
日程を調整し、ステータスを「確定」に変更すると確定メールが自動送信されます。`;
}

/**
 * 予約確定メール（申込者向け）- 企業URL・事前リサーチ案内付き
 */
function getConfirmedEmailBody(data) {
  let locationInfo = '';

  if (data.method === 'オンライン' || data.method === 'オンライン（Zoom）') {
    locationInfo = data.zoomUrl
      ? `【オンライン相談】
Zoom URL：${data.zoomUrl}

※開始時刻の5分前を目安にご参加ください
※接続に不具合がある場合は、こちらからご連絡を差し上げることがあります
※相談内容の品質向上・記録のため、原則としてZoom録画を実施いたします`
      : `【オンライン相談】
Zoomのアドレスについては、前日までに改めてメールでお送りいたします。

※開始時刻の5分前を目安にご参加ください
※接続に不具合がある場合は、こちらからご連絡を差し上げることがあります
※相談内容の品質向上・記録のため、原則としてZoom録画を実施いたします`;
  } else {
    const loc = data.location || '（後日ご案内）';
    locationInfo = `【対面相談】
場所：${loc}

※受付にて「中小企業経営診断研究会の相談予約」とお伝えください`;
  }

  return `${data.name} 様

無料経営相談のご予約が確定しましたのでお知らせいたします。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご予約内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
日時：${data.confirmedDate}
相談方法：${data.method}
担当：${data.staff}

${locationInfo}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【当日の流れ】
1. 現状のヒアリング（15分程度）
   - お話を伺います

2. 課題の整理・ディスカッション（30〜45分）
   - 課題を整理し、解決の方向性を一緒に考えます

3. 今後のアクション整理（15分程度）
   - 次のステップを明確にします

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご準備いただくもの
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
・関連資料（決算書、事業計画書等）があればお持ちください
  ※必須ではありません

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ キャンセル・変更について
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ご都合が悪くなった場合は、できるだけ早めにご連絡ください。
連絡先：${CONFIG.ORG.EMAIL}

ご不明な点がございましたら、お気軽にお問い合わせください。
当日お会いできることを楽しみにしております。

【ご相談の流れ・よくある質問】
相談者マニュアル: https://hirokazusugisugi-hue.github.io/kg-consultation/docs/manual_consultee.html

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
URL: ${CONFIG.ORG.URL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}

/**
 * リマインドメール（3日前・準備案内）
 */
function getReminderEmail3DaysBefore(data) {
  return `${data.name} 様

3日後にご相談のご予約をいただいております。
ご準備のご案内をさせていただきます。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご予約内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
日時：${data.confirmedDate}
相談方法：${data.method}
担当：${data.staff}
${data.method === 'オンライン' || data.method === 'オンライン（Zoom）' ? 'Zoom URL：' + (data.zoomUrl || '（確定メールをご確認ください）') : ''}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【ご準備のお願い】
・関連資料（決算書、事業計画書等）があればご準備ください
・ご相談されたい内容を整理しておいていただけると、
  より充実した相談時間となります

ご不明な点がございましたら、お気軽にお問い合わせください。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}

/**
 * リマインドメール（前日・最終確認）
 */
function getReminderEmailDayBefore(data) {
  return `${data.name} 様

明日のご相談について、最終確認のご連絡です。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご予約内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
日時：${data.confirmedDate}
相談方法：${data.method}
担当：${data.staff}
${data.method === 'オンライン' || data.method === 'オンライン（Zoom）' ? 'Zoom URL：' + (data.zoomUrl || '（確定メールをご確認ください）') : ''}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【最終確認事項】
・日時・方法に変更はございませんか？
・変更・キャンセルの場合は早急にご連絡ください
${data.method === 'オンライン' || data.method === 'オンライン（Zoom）' ? '・Zoomの接続テストを事前にお願いいたします\n・相談内容の品質向上・記録のため、原則としてZoom録画を実施いたします' : '・当日は受付にて「中小企業経営診断研究会の相談予約」とお伝えください'}

当日お会いできることを楽しみにしております。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}

/**
 * 担当者向けLINEリマインドメッセージ（拡張版）
 * @param {Object} data - 予約データ（rowData）
 * @param {string} daysBeforeLabel - "1週間前" or "3日前"
 * @param {Array<Object>} memberList - 参加メンバー情報の配列 [{name, term, type}]
 */
function getStaffReminderLine(data, daysBeforeLabel, memberList) {
  const memberNames = memberList ? memberList.map(function(m) { return m.name; }).join(', ') : (data.staff || '');
  const isOnline = data.method === 'オンライン' || data.method === 'オンライン（Zoom）';
  const venue = isOnline ? 'Zoom' : (data.location || '未定');

  return `📋 【${daysBeforeLabel}】担当相談リマインド

日時: ${data.confirmedDate}
会場: ${venue}${isOnline && data.zoomUrl ? '\nZoom: ' + data.zoomUrl : ''}
相談者: ${data.name}様（${data.company}）
電話: ${data.phone || '未登録'}
テーマ: ${data.theme}
${data.companyUrl ? '企業URL: ' + data.companyUrl : ''}
リーダー: ${data.leader || '未選定'}
担当メンバー: ${memberNames}

📖 担当者マニュアル:
https://hirokazusugisugi-hue.github.io/kg-consultation/docs/manual_staff.html

📝 オブザーバー専用ページ:
${CONFIG.CONSENT.WEB_APP_URL}?action=observer

事前準備をお願いします。`;
}

/**
 * 担当者向けメールリマインド（拡張版）
 * @param {Object} data - 予約データ（rowData）
 * @param {string} daysBeforeLabel - "1週間前" or "3日前"
 * @param {Array<Object>} memberList - 参加メンバー情報の配列 [{name, term, type}]
 */
function getStaffReminderEmail(data, daysBeforeLabel, memberList) {
  const isOnline = data.method === 'オンライン' || data.method === 'オンライン（Zoom）';

  let venueInfo = '';
  if (isOnline) {
    venueInfo = 'Zoom' + (data.zoomUrl ? '\nZoom URL：' + data.zoomUrl : '');
  } else {
    venueInfo = data.location || '（未定）';
  }

  let memberSection = '';
  if (memberList && memberList.length > 0) {
    memberSection = memberList.map(function(m) {
      const role = m.term ? '（' + m.term + '）' : '';
      return '  ' + m.name + ' ' + role;
    }).join('\n');
  } else {
    memberSection = '  ' + (data.staff || '未定');
  }

  return `【${daysBeforeLabel}】担当相談のリマインド

${daysBeforeLabel}に以下の相談が予定されています。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 日時・会場
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
日時：${data.confirmedDate}
相談方法：${data.method}
会場：${venueInfo}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 相談者情報
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
氏名：${data.name} 様
企業名：${data.company}
メール：${data.email}
電話番号：${data.phone || '未登録'}
テーマ：${data.theme}
${data.companyUrl ? '企業URL：' + data.companyUrl + '\n※事前リサーチにご活用ください' : ''}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 担当メンバー
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
リーダー：${data.leader || '未選定'}

${memberSection}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 参考リンク
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
担当者マニュアル:
https://hirokazusugisugi-hue.github.io/kg-consultation/docs/manual_staff.html

オブザーバー専用ページ（NDA提出状況確認）:
${CONFIG.CONSENT.WEB_APP_URL}?action=observer

━━━━━━━━━━━━━━━━━━━━━━━━━━━━

事前準備をお願いいたします。`;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 同意書ページHTML テンプレート
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * 同意書確認ページHTML（PDF埋め込み版）
 */
function getConsentPageHtml(data, token) {
  const pdfDirectUrl = CONFIG.CONSENT.PDF_URL || '';
  const pdfFileId = CONFIG.CONSENT.PDF_FILE_ID;
  const pdfViewerUrl = pdfDirectUrl
    ? 'https://docs.google.com/gview?embedded=true&url=' + encodeURIComponent(pdfDirectUrl)
    : 'https://drive.google.com/file/d/' + pdfFileId + '/preview';
  const pdfDownloadUrl = pdfDirectUrl || ('https://drive.google.com/uc?export=download&id=' + pdfFileId);

  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>相談同意書のご確認 - 関西学院大学 中小企業経営診断研究会</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Noto Sans JP', sans-serif;
      background: #f5f5f7;
      color: #1a1a1a;
      line-height: 1.8;
    }
    .container {
      max-width: 800px;
      margin: 0 auto;
      padding: 2rem 1.5rem;
    }
    .header {
      background: #0F2350;
      color: #fff;
      padding: 2rem 0;
      text-align: center;
    }
    .header h1 {
      font-size: 1.3rem;
      font-weight: 700;
      margin-bottom: 0.5rem;
    }
    .header p {
      font-size: 0.85rem;
      opacity: 0.8;
    }
    .card {
      background: #fff;
      border-radius: 12px;
      padding: 2rem;
      margin-bottom: 1.5rem;
      box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .card h2 {
      font-size: 1.1rem;
      font-weight: 700;
      margin-bottom: 1rem;
      padding-bottom: 0.5rem;
      border-bottom: 2px solid #0F2350;
    }
    .applicant-info {
      display: grid;
      grid-template-columns: 100px 1fr;
      gap: 0.5rem;
      font-size: 0.9rem;
    }
    .applicant-info dt {
      font-weight: 600;
      color: #666;
    }
    .pdf-viewer {
      width: 100%;
      height: 500px;
      border: 1px solid #e0e0e0;
      border-radius: 8px;
    }
    .pdf-download {
      display: inline-flex;
      align-items: center;
      gap: 0.5rem;
      margin-top: 0.8rem;
      padding: 0.5rem 1rem;
      background: #f0f0f0;
      border-radius: 6px;
      color: #0F2350;
      text-decoration: none;
      font-size: 0.85rem;
      font-weight: 600;
      transition: background 0.3s;
    }
    .pdf-download:hover { background: #e0e0e0; }
    .consent-section {
      margin-top: 1.5rem;
    }
    .checkbox-group {
      display: flex;
      align-items: flex-start;
      gap: 0.8rem;
      padding: 1rem;
      background: #fff3cd;
      border-radius: 8px;
      margin-bottom: 1rem;
    }
    .checkbox-group input[type="checkbox"] {
      width: 20px;
      height: 20px;
      margin-top: 0.2rem;
      flex-shrink: 0;
    }
    .checkbox-group label {
      font-size: 0.9rem;
      font-weight: 600;
      cursor: pointer;
    }
    .signature-group {
      margin-bottom: 1.5rem;
    }
    .signature-group label {
      display: block;
      font-size: 0.9rem;
      font-weight: 600;
      margin-bottom: 0.5rem;
    }
    .signature-group input {
      width: 100%;
      padding: 0.8rem 1rem;
      border: 1px solid #ddd;
      border-radius: 8px;
      font-size: 1rem;
      font-family: inherit;
    }
    .signature-group input:focus {
      outline: none;
      border-color: #0F2350;
      box-shadow: 0 0 0 3px rgba(15,35,80,0.1);
    }
    .submit-btn {
      width: 100%;
      padding: 1rem;
      background: #0F2350;
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 1rem;
      font-weight: 700;
      cursor: pointer;
      transition: opacity 0.3s;
    }
    .submit-btn:hover { opacity: 0.9; }
    .submit-btn:disabled {
      background: #ccc;
      cursor: not-allowed;
    }
    .error-msg {
      color: #dc3545;
      font-size: 0.85rem;
      margin-top: 0.5rem;
      display: none;
    }
    .success-overlay {
      display: none;
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(0,0,0,0.5);
      z-index: 1000;
      justify-content: center;
      align-items: center;
    }
    .success-overlay.active { display: flex; }
    .success-box {
      background: #fff;
      border-radius: 12px;
      padding: 3rem 2rem;
      text-align: center;
      max-width: 500px;
      margin: 1rem;
    }
    .success-icon {
      width: 60px; height: 60px;
      background: #d4edda;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 1rem;
      font-size: 1.5rem;
      color: #28a745;
    }
  </style>
</head>
<body>
  <div class="header">
    <h1>相談同意書のご確認</h1>
    <p>関西学院大学 中小企業経営診断研究会</p>
  </div>

  <div class="container">
    <div class="card">
      <h2>お申込者情報</h2>
      <dl class="applicant-info">
        <dt>申込ID</dt><dd>${data.id}</dd>
        <dt>お名前</dt><dd>${data.name}</dd>
        <dt>貴社名</dt><dd>${data.company || '（個人）'}</dd>
        <dt>相談テーマ</dt><dd>${data.theme}</dd>
      </dl>
    </div>

    <div class="card">
      <h2>経営相談に関する同意書</h2>
      <p style="font-size:0.85rem; color:#666; margin-bottom:1rem;">関西学院大学 中小企業診断士養成課程（無料経営診断分科会）</p>
      <iframe class="pdf-viewer" src="${pdfViewerUrl}" allow="autoplay"></iframe>
      <a href="${pdfDownloadUrl}" target="_blank" class="pdf-download">PDFをダウンロード</a>

      <div class="consent-section">
        <div class="checkbox-group">
          <input type="checkbox" id="agreeCheck">
          <label for="agreeCheck">上記同意書の内容を確認し、全ての内容に同意します</label>
        </div>

        <div class="signature-group">
          <label for="signature">電子署名（お名前をご入力ください）</label>
          <input type="text" id="signature" placeholder="${data.name}" required>
        </div>

        <div id="errorMsg" class="error-msg"></div>

        <button id="submitBtn" class="submit-btn" disabled onclick="submitConsent()">
          同意して送信
        </button>
      </div>
    </div>
  </div>

  <div id="successOverlay" class="success-overlay">
    <div class="success-box">
      <div class="success-icon">&#10003;</div>
      <h3>同意が完了しました</h3>
      <p style="margin-top: 1rem; color: #666;">確定のメールが届きます。<br>このページを閉じていただいて結構です。</p>
    </div>
  </div>

  <script>
    const checkbox = document.getElementById('agreeCheck');
    const signatureInput = document.getElementById('signature');
    const submitBtn = document.getElementById('submitBtn');

    function updateButtonState() {
      submitBtn.disabled = !(checkbox.checked && signatureInput.value.trim());
    }

    checkbox.addEventListener('change', updateButtonState);
    signatureInput.addEventListener('input', updateButtonState);

    function submitConsent() {
      submitBtn.disabled = true;
      submitBtn.textContent = '送信中...';

      const formData = {
        token: '${token}',
        agreed: 'true',
        signature: signatureInput.value.trim()
      };

      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            document.getElementById('successOverlay').classList.add('active');
          } else {
            document.getElementById('errorMsg').textContent = result.message;
            document.getElementById('errorMsg').style.display = 'block';
            submitBtn.disabled = false;
            submitBtn.textContent = '同意して送信';
          }
        })
        .withFailureHandler(function(error) {
          document.getElementById('errorMsg').textContent = 'エラーが発生しました。もう一度お試しください。';
          document.getElementById('errorMsg').style.display = 'block';
          submitBtn.disabled = false;
          submitBtn.textContent = '同意して送信';
        })
        .submitNdaConsent(formData);
    }
  </script>
</body>
</html>`;
}

/**
 * 同意書エラーページHTML
 */
function getConsentErrorPageHtml() {
  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>エラー - 同意書確認</title>
  <style>
    body { font-family: 'Noto Sans JP', sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; background: #f5f5f7; }
    .error-box { background: #fff; padding: 3rem; border-radius: 12px; text-align: center; max-width: 500px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
    .error-icon { font-size: 3rem; color: #dc3545; margin-bottom: 1rem; }
    h2 { margin-bottom: 1rem; }
    p { color: #666; }
  </style>
</head>
<body>
  <div class="error-box">
    <div class="error-icon">&#9888;</div>
    <h2>無効なリンクです</h2>
    <p>この同意書確認リンクは無効か、既に使用済みです。<br>お心当たりがない場合は、お問い合わせください。</p>
  </div>
</body>
</html>`;
}

/**
 * 同意済ページHTML
 */
function getConsentAlreadyAgreedPageHtml(data) {
  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>同意済 - 関西学院大学 中小企業経営診断研究会</title>
  <style>
    body { font-family: 'Noto Sans JP', sans-serif; display: flex; justify-content: center; align-items: center; min-height: 100vh; background: #f5f5f7; }
    .box { background: #fff; padding: 3rem; border-radius: 12px; text-align: center; max-width: 500px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); }
    .icon { width: 60px; height: 60px; background: #d4edda; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin: 0 auto 1rem; font-size: 1.5rem; color: #28a745; }
    h2 { margin-bottom: 1rem; }
    p { color: #666; }
  </style>
</head>
<body>
  <div class="box">
    <div class="icon">&#10003;</div>
    <h2>同意書への同意は既に完了しています</h2>
    <p>${data.name} 様（申込ID: ${data.id}）<br>相談同意書への同意は既に受領済みです。<br>確定のメールが届きます。</p>
  </div>
</body>
</html>`;
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// アンケートページHTML
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getSurveyPageHtml(tokenData) {
  var webAppUrl = CONFIG.CONSENT.WEB_APP_URL;
  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>相談後アンケート</title>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Noto Sans JP', sans-serif; background: #f5f5f7; color: #1a1a1a; line-height: 1.8; }
    .header { background: #0F2350; color: #fff; padding: 1.5rem; text-align: center; }
    .header h1 { font-size: 1.2rem; font-weight: 500; }
    .container { max-width: 700px; margin: 0 auto; padding: 1.5rem 1rem; }
    .card { background: #fff; border-radius: 12px; padding: 1.5rem; margin-bottom: 1.2rem; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
    .card h2 { font-size: 1rem; color: #0F2350; margin-bottom: 0.8rem; border-left: 4px solid #0F2350; padding-left: 0.8rem; }
    .q-label { font-weight: 500; margin-bottom: 0.5rem; font-size: 0.95rem; }
    .required { color: #c00; font-size: 0.75rem; margin-left: 0.3rem; }
    .radio-group, .check-group { margin: 0.5rem 0 1rem 0; }
    .radio-group label, .check-group label { display: block; padding: 0.5rem 0.8rem; margin: 0.3rem 0; background: #f8f9fa; border-radius: 8px; cursor: pointer; font-size: 0.9rem; transition: background 0.2s; }
    .radio-group label:hover, .check-group label:hover { background: #eef3ff; }
    .radio-group input, .check-group input { margin-right: 0.5rem; }
    .sns-sub { margin-left: 2rem; padding: 0.5rem; background: #f0f4ff; border-radius: 8px; display: none; }
    .sns-sub.show { display: block; }
    .sns-sub label { display: inline-block; margin: 0.2rem 0.5rem; font-size: 0.85rem; }
    .scale-group { display: flex; justify-content: space-between; margin: 0.5rem 0; gap: 0.3rem; }
    .scale-group label { flex: 1; text-align: center; padding: 0.6rem 0.2rem; background: #f8f9fa; border-radius: 8px; cursor: pointer; font-size: 0.8rem; transition: all 0.2s; }
    .scale-group input { display: none; }
    .scale-group input:checked + span { background: #0F2350; color: #fff; display: block; border-radius: 8px; padding: 0.6rem 0.2rem; margin: -0.6rem -0.2rem; }
    .scale-labels { display: flex; justify-content: space-between; font-size: 0.7rem; color: #888; margin-top: 0.2rem; }
    textarea { width: 100%; border: 1px solid #ddd; border-radius: 8px; padding: 0.8rem; font-size: 0.9rem; font-family: inherit; resize: vertical; min-height: 80px; }
    .comment-box { display: none; margin-top: 0.5rem; }
    .comment-box.show { display: block; }
    .btn { display: block; width: 100%; padding: 1rem; background: #0F2350; color: #fff; border: none; border-radius: 12px; font-size: 1rem; font-weight: 500; cursor: pointer; margin-top: 1rem; }
    .btn:hover { background: #1a3570; }
    .btn:disabled { background: #ccc; cursor: not-allowed; }
    .success { text-align: center; padding: 3rem 1rem; }
    .success h2 { color: #0F2350; margin-bottom: 1rem; }
    .info-bar { background: #eef3ff; padding: 0.8rem 1rem; border-radius: 8px; font-size: 0.85rem; margin-bottom: 1rem; }
  </style>
</head>
<body>
  <div class="header">
    <h1>無料経営相談 アンケート</h1>
    <p style="font-size:0.8rem; opacity:0.8; margin-top:0.3rem;">関西学院大学 中小企業経営診断研究会</p>
  </div>

  <div class="container" id="formContainer">
    <div class="info-bar">
      ${tokenData.name} 様（${tokenData.company || ''}）<br>
      ご利用ありがとうございました。今後の改善のため、アンケートにご協力ください。
    </div>

    <!-- Q1 -->
    <div class="card">
      <div class="q-label">Q1. 本プロジェクトを知ったきっかけ<span class="required">※複数選択可</span></div>
      <div class="check-group">
        <label><input type="checkbox" name="q1" value="大学の案内・掲示"> 大学の案内・掲示</label>
        <label><input type="checkbox" name="q1" value="知人・友人の紹介"> 知人・友人の紹介</label>
        <label><input type="checkbox" name="q1" value="SNS" id="q1Sns"> SNS</label>
        <div class="sns-sub" id="snsSub">
          <label><input type="radio" name="q1sns" value="LINE"> LINE</label>
          <label><input type="radio" name="q1sns" value="Twitter/X"> Twitter/X</label>
          <label><input type="radio" name="q1sns" value="Facebook"> Facebook</label>
          <label><input type="radio" name="q1sns" value="Instagram"> Instagram</label>
          <label><input type="radio" name="q1sns" value="その他SNS"> その他</label>
        </div>
        <label><input type="checkbox" name="q1" value="インターネット検索"> インターネット検索</label>
        <label><input type="checkbox" name="q1" value="商工会議所・支援機関からの紹介"> 商工会議所・支援機関からの紹介</label>
        <label><input type="checkbox" name="q1" value="その他"> その他</label>
        <div class="comment-box" id="q1Other"><textarea id="q1OtherText" placeholder="具体的にお聞かせください"></textarea></div>
      </div>
    </div>

    <!-- Q2 -->
    <div class="card">
      <div class="q-label">Q2. 相談までの手続きはスムーズでしたか？<span class="required">※必須</span></div>
      <div class="scale-group" id="q2Scale">
        <label><input type="radio" name="q2" value="1"><span>1<br>スムーズ<br>でない</span></label>
        <label><input type="radio" name="q2" value="2"><span>2</span></label>
        <label><input type="radio" name="q2" value="3"><span>3</span></label>
        <label><input type="radio" name="q2" value="4"><span>4</span></label>
        <label><input type="radio" name="q2" value="5"><span>5<br>とても<br>スムーズ</span></label>
      </div>
      <div class="comment-box" id="q2Comment"><textarea id="q2CommentText" placeholder="改善点をお聞かせください"></textarea></div>
    </div>

    <!-- Q3 -->
    <div class="card">
      <div class="q-label">Q3. 相談内容についての感想</div>
      <textarea id="q3" placeholder="ご自由にお書きください"></textarea>
    </div>

    <!-- Q4 -->
    <div class="card">
      <div class="q-label">Q4. 相談時間について<span class="required">※必須</span></div>
      <div class="radio-group">
        <label><input type="radio" name="q4" value="長すぎた"> 長すぎた</label>
        <label><input type="radio" name="q4" value="やや長い"> やや長い</label>
        <label><input type="radio" name="q4" value="ちょうどよい"> ちょうどよい</label>
        <label><input type="radio" name="q4" value="やや短い"> やや短い</label>
        <label><input type="radio" name="q4" value="短すぎた"> 短すぎた</label>
      </div>
    </div>

    <!-- Q5-Q8 -->
    <div class="card">
      <div class="q-label">Q5. 相談員の説明はわかりやすかったですか？<span class="required">※必須</span></div>
      <div class="scale-group"><label><input type="radio" name="q5" value="1"><span>1</span></label><label><input type="radio" name="q5" value="2"><span>2</span></label><label><input type="radio" name="q5" value="3"><span>3</span></label><label><input type="radio" name="q5" value="4"><span>4</span></label><label><input type="radio" name="q5" value="5"><span>5</span></label></div>
      <div class="scale-labels"><span>わかりにくい</span><span>とてもわかりやすい</span></div>
    </div>

    <div class="card">
      <div class="q-label">Q6. 相談内容は課題解決の参考になりましたか？<span class="required">※必須</span></div>
      <div class="scale-group"><label><input type="radio" name="q6" value="1"><span>1</span></label><label><input type="radio" name="q6" value="2"><span>2</span></label><label><input type="radio" name="q6" value="3"><span>3</span></label><label><input type="radio" name="q6" value="4"><span>4</span></label><label><input type="radio" name="q6" value="5"><span>5</span></label></div>
      <div class="scale-labels"><span>参考にならない</span><span>とても参考になった</span></div>
    </div>

    <div class="card">
      <div class="q-label">Q7. 相談員の対応は誠実・丁寧でしたか？<span class="required">※必須</span></div>
      <div class="scale-group"><label><input type="radio" name="q7" value="1"><span>1</span></label><label><input type="radio" name="q7" value="2"><span>2</span></label><label><input type="radio" name="q7" value="3"><span>3</span></label><label><input type="radio" name="q7" value="4"><span>4</span></label><label><input type="radio" name="q7" value="5"><span>5</span></label></div>
      <div class="scale-labels"><span>不満</span><span>とても満足</span></div>
    </div>

    <div class="card">
      <div class="q-label">Q8. 具体的な行動につながるアドバイスがありましたか？<span class="required">※必須</span></div>
      <div class="scale-group"><label><input type="radio" name="q8" value="1"><span>1</span></label><label><input type="radio" name="q8" value="2"><span>2</span></label><label><input type="radio" name="q8" value="3"><span>3</span></label><label><input type="radio" name="q8" value="4"><span>4</span></label><label><input type="radio" name="q8" value="5"><span>5</span></label></div>
      <div class="scale-labels"><span>なかった</span><span>とてもあった</span></div>
    </div>

    <!-- Q9 -->
    <div class="card">
      <div class="q-label">Q9. また相談を受けてみたいですか？<span class="required">※必須</span></div>
      <div class="scale-group">
        <label><input type="radio" name="q9" value="5"><span>5<br>ぜひ<br>受けたい</span></label>
        <label><input type="radio" name="q9" value="4"><span>4<br>受けたい</span></label>
        <label><input type="radio" name="q9" value="3"><span>3<br>どちらでも<br>ない</span></label>
        <label><input type="radio" name="q9" value="2"><span>2<br>あまり受け<br>たくない</span></label>
        <label><input type="radio" name="q9" value="1"><span>1<br>全く受け<br>たくない</span></label>
      </div>
      <textarea id="q9Reason" placeholder="理由をお聞かせください" style="margin-top:0.5rem;"></textarea>
    </div>

    <!-- Q10 -->
    <div class="card">
      <div class="q-label">Q10. 他の方にすすめたいですか？<span class="required">※必須</span></div>
      <div class="scale-group">
        <label><input type="radio" name="q10" value="5"><span>5<br>強く<br>勧めたい</span></label>
        <label><input type="radio" name="q10" value="4"><span>4<br>勧めたい</span></label>
        <label><input type="radio" name="q10" value="3"><span>3<br>どちらでも<br>ない</span></label>
        <label><input type="radio" name="q10" value="2"><span>2<br>勧め<br>たくない</span></label>
        <label><input type="radio" name="q10" value="1"><span>1<br>絶対に勧め<br>たくない</span></label>
      </div>
      <textarea id="q10Reason" placeholder="理由をお聞かせください" style="margin-top:0.5rem;"></textarea>
    </div>

    <!-- Q11 -->
    <div class="card">
      <div class="q-label">Q11. 後日、終了後レポートを希望しますか？</div>
      <div style="background:#fff8e1; border:1px solid #ffe082; border-radius:8px; padding:0.75rem 1rem; margin-bottom:0.75rem; font-size:0.85rem; color:#795548;">
        <strong>&#9888; ご注意：</strong>レポートは相談日から<strong>3日以内</strong>に作成・配信されます。アンケートのご回答が相談日から3日を過ぎた場合、レポートをお届けできない場合がございます。レポートをご希望の方は、お早めにアンケートにご回答ください。
      </div>
      <div class="radio-group">
        <label><input type="radio" name="q11" value="希望する"> 希望する</label>
        <label><input type="radio" name="q11" value="希望しない"> 希望しない</label>
      </div>
    </div>

    <button class="btn" id="submitBtn" onclick="submitSurvey()">回答を送信する</button>
  </div>

  <div class="container" id="successContainer" style="display:none;">
    <div class="card success">
      <h2>ご回答ありがとうございました</h2>
      <p>今後のサービス向上に役立ててまいります。</p>
    </div>
  </div>

  <script>
    // SNSサブ選択の表示切替
    document.getElementById('q1Sns').addEventListener('change', function() {
      document.getElementById('snsSub').classList.toggle('show', this.checked);
    });

    // Q1その他
    document.querySelectorAll('input[name="q1"][value="その他"]')[0].addEventListener('change', function() {
      document.getElementById('q1Other').classList.toggle('show', this.checked);
    });

    // Q2: 4以下でコメント表示
    document.querySelectorAll('input[name="q2"]').forEach(function(r) {
      r.addEventListener('change', function() {
        document.getElementById('q2Comment').classList.toggle('show', parseInt(this.value) <= 4);
      });
    });

    // 5段階のスタイル
    document.querySelectorAll('.scale-group label').forEach(function(label) {
      label.addEventListener('click', function() {
        var group = this.parentElement;
        group.querySelectorAll('label').forEach(function(l) { l.style.background = '#f8f9fa'; l.style.color = '#1a1a1a'; });
        this.style.background = '#0F2350';
        this.style.color = '#fff';
      });
    });

    function submitSurvey() {
      var btn = document.getElementById('submitBtn');
      btn.disabled = true;
      btn.textContent = '送信中...';

      var q1vals = [];
      document.querySelectorAll('input[name="q1"]:checked').forEach(function(c) { q1vals.push(c.value); });
      var q1other = document.getElementById('q1OtherText').value;
      if (q1other) q1vals.push('その他:' + q1other);

      var formData = {
        applicationId: '${tokenData.applicationId || ""}',
        name: '${tokenData.name || ""}',
        company: '${tokenData.company || ""}',
        q1: q1vals.join(', '),
        q1Sns: (document.querySelector('input[name="q1sns"]:checked') || {}).value || '',
        q2: (document.querySelector('input[name="q2"]:checked') || {}).value || '',
        q2Comment: document.getElementById('q2CommentText').value,
        q3: document.getElementById('q3').value,
        q4: (document.querySelector('input[name="q4"]:checked') || {}).value || '',
        q5: (document.querySelector('input[name="q5"]:checked') || {}).value || '',
        q6: (document.querySelector('input[name="q6"]:checked') || {}).value || '',
        q7: (document.querySelector('input[name="q7"]:checked') || {}).value || '',
        q8: (document.querySelector('input[name="q8"]:checked') || {}).value || '',
        q9: (document.querySelector('input[name="q9"]:checked') || {}).value || '',
        q9Reason: document.getElementById('q9Reason').value,
        q10: (document.querySelector('input[name="q10"]:checked') || {}).value || '',
        q10Reason: document.getElementById('q10Reason').value,
        q11: (document.querySelector('input[name="q11"]:checked') || {}).value || ''
      };

      google.script.run
        .withSuccessHandler(function(result) {
          document.getElementById('formContainer').style.display = 'none';
          document.getElementById('successContainer').style.display = 'block';
        })
        .withFailureHandler(function(err) {
          alert('送信エラー: ' + err.message);
          btn.disabled = false;
          btn.textContent = '回答を送信する';
        })
        .submitSurveyResponse(formData);
    }
  </script>
</body>
</html>`;
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// オブザーバー専用ページHTML（拡張版）
// 企業名・相談予定可能者・オブザーバー表示、サーバーサイドPDF生成
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function getObserverPageHtml(schedules) {
  var schedulesJson = JSON.stringify(schedules);
  return `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>オブザーバー専用ページ</title>
  <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'Noto Sans JP', sans-serif; background: #f5f5f7; color: #1a1a1a; line-height: 1.8; }
    .header { background: #0F2350; color: #fff; padding: 1.5rem; text-align: center; }
    .header h1 { font-size: 1.2rem; font-weight: 500; }
    .container { max-width: 700px; margin: 0 auto; padding: 1.5rem 1rem; }
    .card { background: #fff; border-radius: 12px; padding: 1.5rem; margin-bottom: 1.2rem; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
    .card h2 { font-size: 1rem; color: #0F2350; margin-bottom: 1rem; border-left: 4px solid #0F2350; padding-left: 0.8rem; }

    /* スケジュールカード */
    .schedule-card { border: 1px solid #e0e4ea; border-radius: 10px; padding: 1.2rem; margin-bottom: 1rem; transition: box-shadow 0.2s; }
    .schedule-card:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.08); }
    .schedule-date { font-size: 1.1rem; font-weight: 700; color: #0F2350; margin-bottom: 0.8rem; padding-bottom: 0.5rem; border-bottom: 2px solid #eef3ff; }
    .schedule-detail { display: grid; grid-template-columns: auto 1fr; gap: 0.3rem 0.8rem; font-size: 0.9rem; margin-bottom: 0.8rem; }
    .schedule-label { font-weight: 600; color: #555; white-space: nowrap; }
    .schedule-value { color: #1a1a1a; }
    .observer-badges { display: flex; flex-wrap: wrap; gap: 0.3rem; }
    .badge { display: inline-block; padding: 0.15rem 0.6rem; border-radius: 20px; font-size: 0.8rem; }
    .badge-submitted { background: #d4edda; color: #155724; }
    .badge-scheduled { background: #fff3cd; color: #856404; }
    .badge-none { background: #f0f0f0; color: #888; font-style: italic; }

    .btn { padding: 0.6rem 1.2rem; border: none; border-radius: 8px; font-size: 0.85rem; cursor: pointer; font-family: inherit; font-weight: 500; }
    .btn-primary { background: #0F2350; color: #fff; }
    .btn-primary:hover { background: #1a3570; }
    .btn-secondary { background: #e8ecf1; color: #0F2350; }
    .btn-secondary:hover { background: #d0d8e4; }

    /* 署名モーダル */
    .modal-overlay { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 100; }
    .modal-overlay.show { display: flex; align-items: center; justify-content: center; }
    .modal { background: #fff; border-radius: 16px; padding: 1.5rem; width: 95%; max-width: 500px; max-height: 90vh; overflow-y: auto; }
    .modal h2 { font-size: 1.1rem; color: #0F2350; margin-bottom: 1rem; }
    .modal .field { margin-bottom: 0.8rem; }
    .modal .field label { font-size: 0.85rem; font-weight: 500; display: block; margin-bottom: 0.3rem; }
    .modal .field input { width: 100%; padding: 0.5rem; border: 1px solid #ddd; border-radius: 6px; font-size: 0.9rem; }
    .modal .field input:read-only { background: #f0f0f0; }

    /* 署名Canvas */
    .sig-container { border: 2px solid #0F2350; border-radius: 8px; margin: 0.5rem 0; position: relative; background: #fff; }
    .sig-container canvas { display: block; width: 100%; touch-action: none; }
    .sig-label { font-size: 0.8rem; color: #888; text-align: center; padding: 0.3rem; }
    .sig-clear { position: absolute; top: 0.3rem; right: 0.3rem; background: #e8ecf1; border: none; border-radius: 4px; padding: 0.2rem 0.5rem; font-size: 0.75rem; cursor: pointer; }

    .nda-preview { max-height: 200px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px; padding: 1rem; font-size: 0.8rem; margin: 0.5rem 0; background: #fafafa; }
    .status-msg { text-align: center; padding: 1rem; font-size: 0.9rem; }
    .status-msg.success { color: #28a745; }
    .status-msg.error { color: #c00; }
    .empty-msg { text-align: center; color: #888; padding: 2rem; }
  </style>
</head>
<body>
  <div class="header">
    <h1>オブザーバー専用ページ</h1>
    <p style="font-size:0.8rem; opacity:0.8; margin-top:0.3rem;">関西学院大学 中小企業経営診断研究会</p>
  </div>

  <div class="container">
    <div class="card">
      <h2>相談予定一覧</h2>
      <div id="scheduleList"></div>
    </div>
  </div>

  <!-- 署名モーダル -->
  <div class="modal-overlay" id="signModal">
    <div class="modal">
      <h2>秘密保持誓約書 署名</h2>
      <div class="field">
        <label>相談日</label>
        <input type="text" id="signDate" readonly>
      </div>
      <div class="field">
        <label>相談予定可能者</label>
        <input type="text" id="signStaff" readonly>
      </div>
      <div class="field">
        <label>相談企業名</label>
        <input type="text" id="signCompany" readonly>
      </div>
      <div class="field">
        <label>オブザーバー氏名</label>
        <input type="text" id="signName" placeholder="氏名を入力">
      </div>

      <div class="nda-preview">
        <strong>秘密保持誓約書（養成課程在学生・オブザーバー用）</strong><br><br>
        私は、「経営診断研究会 無料経営相談分科会」（以下「本分科会」といいます）が実施する無料経営相談にオブザーバーとして出席するにあたり、個人の責任として、以下の事項を遵守することを誓約いたします。<br><br>
        <strong>第1条（秘密情報の定義）</strong><br>
        本誓約における「秘密情報」とは、本分科会の活動を通じて知り得た、相談企業の経営・財務・技術等の情報、関係者の個人情報、および活動中に作成された相談資料・録音データ等、一切の情報を指します。<br><br>
        <strong>第2条（遵守事項）</strong><br>
        1. 本分科会の正規メンバー以外の第三者に、秘密情報を開示・漏洩しないこと。<br>
        2. 経営相談およびそれに伴う学術研究・教育以外の目的で、秘密情報を使用しないこと。<br>
        3. SNSやブログ等のインターネット上に、相談企業が特定できる情報や活動内容を投稿しないこと。<br>
        4. 学術・教育目的で事例を利用する場合は、相談企業の事前同意に基づき、企業および個人が特定されないよう厳格な匿名化・統計化処理を施すこと。<br>
        5. 活動終了時または研究会の指示があった際は、秘密情報を含む資料・データ等を速やかに返還または廃棄すること。<br><br>
        <strong>第3条（期間および損害賠償）</strong><br>
        1. 本誓約の義務は、本分科会の活動終了後および養成課程修了後も存続するものとします。<br>
        2. 本誓約に違反し、研究会または相談企業に損害を与えた場合は、法的責任を負うとともに、研究会の処分に従います。
      </div>

      <div class="sig-container">
        <div class="sig-label">署名欄（指またはペンで署名してください）</div>
        <canvas id="sigCanvas" width="460" height="150"></canvas>
        <button class="sig-clear" onclick="clearSignature()">やり直し</button>
      </div>

      <div id="signStatus" class="status-msg"></div>

      <div style="display:flex; gap:0.5rem; justify-content:center; margin-top:1rem;">
        <button class="btn btn-secondary" onclick="closeSignModal()">キャンセル</button>
        <button class="btn btn-primary" id="signSubmitBtn" onclick="submitSignedNda()">署名して提出</button>
      </div>
    </div>
  </div>

  <script>
    var schedules = ${schedulesJson};
    var sigCanvas, sigCtx, isDrawing = false;

    // HTMLエスケープ
    function escHtml(str) {
      var d = document.createElement('div');
      d.textContent = str || '';
      return d.innerHTML;
    }

    // 相談予定一覧の描画（拡張版）
    function renderSchedules() {
      var list = document.getElementById('scheduleList');
      if (!schedules || schedules.length === 0) {
        list.innerHTML = '<div class="empty-msg">現在、相談予定はありません</div>';
        return;
      }
      var html = '';
      schedules.forEach(function(s, idx) {
        var consultants = s.members || s.staff;
        var observerHtml = '';
        var ndaSubmitted = s.ndaSubmitted || [];
        if (s.observers && s.observers.length > 0) {
          s.observers.forEach(function(name) {
            if (ndaSubmitted.indexOf(name) >= 0) {
              observerHtml += '<span class="badge badge-submitted">' + escHtml(name) + '（NDA提出済）</span>';
            } else {
              observerHtml += '<span class="badge badge-scheduled">' + escHtml(name) + '（未提出）</span>';
            }
          });
        } else {
          observerHtml = '<span class="badge badge-none">なし</span>';
        }
        html += '<div class="schedule-card">' +
          '<div class="schedule-date">' + escHtml(s.date) + '</div>' +
          '<div class="schedule-detail">' +
            '<span class="schedule-label">相談企業</span>' +
            '<span class="schedule-value">' + escHtml(s.company) + '</span>' +
            '<span class="schedule-label">相談予定可能者</span>' +
            '<span class="schedule-value">' + escHtml(consultants) + '</span>' +
            '<span class="schedule-label">オブザーバー</span>' +
            '<span class="schedule-value"><span class="observer-badges">' + observerHtml + '</span></span>' +
          '</div>' +
          '<button class="btn btn-primary" onclick="openSignModal(' + idx + ')">署名して提出</button>' +
          '</div>';
      });
      list.innerHTML = html;
    }

    // 署名モーダル（インデックスベース：文字列エスケープ問題を回避）
    function openSignModal(idx) {
      var s = schedules[idx];
      document.getElementById('signDate').value = s.dateRaw;
      document.getElementById('signStaff').value = s.members || s.staff;
      document.getElementById('signCompany').value = s.company;
      document.getElementById('signName').value = '';
      document.getElementById('signStatus').textContent = '';
      document.getElementById('signStatus').className = 'status-msg';
      document.getElementById('signSubmitBtn').disabled = false;
      document.getElementById('signSubmitBtn').textContent = '署名して提出';
      document.getElementById('signModal').classList.add('show');
      initSignatureCanvas();
    }

    function closeSignModal() {
      document.getElementById('signModal').classList.remove('show');
    }

    // 署名Canvas
    function initSignatureCanvas() {
      sigCanvas = document.getElementById('sigCanvas');
      sigCtx = sigCanvas.getContext('2d');
      var rect = sigCanvas.parentElement.getBoundingClientRect();
      sigCanvas.width = rect.width - 4;
      sigCanvas.height = 150;
      sigCtx.strokeStyle = '#000';
      sigCtx.lineWidth = 2;
      sigCtx.lineCap = 'round';
      clearSignature();

      sigCanvas.addEventListener('mousedown', startDraw);
      sigCanvas.addEventListener('mousemove', draw);
      sigCanvas.addEventListener('mouseup', stopDraw);
      sigCanvas.addEventListener('touchstart', startDrawTouch, { passive: false });
      sigCanvas.addEventListener('touchmove', drawTouch, { passive: false });
      sigCanvas.addEventListener('touchend', stopDraw);
    }

    function getPos(e) {
      var rect = sigCanvas.getBoundingClientRect();
      return { x: e.clientX - rect.left, y: e.clientY - rect.top };
    }

    function startDraw(e) { isDrawing = true; var p = getPos(e); sigCtx.beginPath(); sigCtx.moveTo(p.x, p.y); }
    function draw(e) { if (!isDrawing) return; var p = getPos(e); sigCtx.lineTo(p.x, p.y); sigCtx.stroke(); }
    function stopDraw() { isDrawing = false; }

    function startDrawTouch(e) { e.preventDefault(); var t = e.touches[0]; startDraw({ clientX: t.clientX, clientY: t.clientY }); }
    function drawTouch(e) { e.preventDefault(); var t = e.touches[0]; draw({ clientX: t.clientX, clientY: t.clientY }); }

    function clearSignature() {
      if (sigCtx) {
        sigCtx.clearRect(0, 0, sigCanvas.width, sigCanvas.height);
        sigCtx.fillStyle = '#fff';
        sigCtx.fillRect(0, 0, sigCanvas.width, sigCanvas.height);
      }
    }

    function isCanvasBlank() {
      var data = sigCtx.getImageData(0, 0, sigCanvas.width, sigCanvas.height).data;
      for (var i = 0; i < data.length; i += 4) {
        if (data[i] < 250 || data[i+1] < 250 || data[i+2] < 250) return false;
      }
      return true;
    }

    // NDA提出（署名画像のみ送信、PDF生成はサーバー側）
    function submitSignedNda() {
      var name = document.getElementById('signName').value.trim();
      if (!name) { alert('氏名を入力してください'); return; }
      if (isCanvasBlank()) { alert('署名を記入してください'); return; }

      var btn = document.getElementById('signSubmitBtn');
      btn.disabled = true;
      btn.textContent = '送信中...';
      document.getElementById('signStatus').textContent = '署名を送信しています...';
      document.getElementById('signStatus').className = 'status-msg';

      // 署名をPNG Base64として取得（data URI prefixを除去）
      var sigDataUrl = sigCanvas.toDataURL('image/png');
      var signatureBase64 = sigDataUrl.split(',')[1];

      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            document.getElementById('signStatus').textContent = '提出完了しました';
            document.getElementById('signStatus').className = 'status-msg success';
            btn.textContent = '提出完了';
          } else {
            document.getElementById('signStatus').textContent = 'エラー: ' + result.message;
            document.getElementById('signStatus').className = 'status-msg error';
            btn.disabled = false;
            btn.textContent = '署名して提出';
          }
        })
        .withFailureHandler(function(err) {
          document.getElementById('signStatus').textContent = 'エラー: ' + err.message;
          document.getElementById('signStatus').className = 'status-msg error';
          btn.disabled = false;
          btn.textContent = '署名して提出';
        })
        .saveSignedNda({
          observerName: name,
          consultDate: document.getElementById('signDate').value,
          company: document.getElementById('signCompany').value,
          staff: document.getElementById('signStaff').value,
          signatureBase64: signatureBase64
        });
    }

    renderSchedules();
  </script>
</body>
</html>`;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// レポート配信メールテンプレート
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

/**
 * リーダーへのレポート作成依頼メール
 */
function getReportRequestEmailBody(data) {
  return `${data.leaderName} 様

お疲れ様です。
下記の経営相談について、診断報告書の作成をお願いいたします。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 相談概要
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.applicationId}
企業名：${data.company}
業種：${data.industry || ''}
テーマ：${data.theme || ''}
相談日：${data.confirmedDate || ''}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【レポートアップロード】
以下のリンクからPDFまたはWordファイルをアップロードしてください。

アップロードページ: ${data.uploadUrl}

【提出期限】
${data.deadlineStr}（相談日から3日以内）

※ファイルサイズは5MB以内でお願いいたします。
※アップロード後、相談者様に自動配信されます。

ご不明な点がございましたら、お気軽にお問い合わせください。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}

/**
 * 相談者へのレポート配信メール
 */
function getReportDeliveryEmailBody(data) {
  return `${data.name} 様

先日は、関西学院大学 中小企業経営診断研究会の無料経営相談をご利用いただき、
誠にありがとうございました。

相談内容をもとに作成した診断報告書をお届けいたします。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 報告書情報
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.applicationId}
貴社名：${data.company}
相談テーマ：${data.theme || ''}
相談日：${data.confirmedDate || ''}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【診断報告書ダウンロード】
${data.fileUrl}

※上記リンクから報告書をダウンロードいただけます。
※本報告書の内容は秘密情報として取り扱われます。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 今後について
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
報告書の内容についてご不明な点がございましたら、
お気軽にお問い合わせください。

また、追加のご相談も随時受け付けております。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
URL: ${CONFIG.ORG.URL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}

/**
 * リーダーへのリマインドメール
 */
function getReportReminderEmailBody(data) {
  return `${data.leaderName} 様

お疲れ様です。
下記の診断報告書について、提出期限のリマインドです。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 対象案件
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.applicationId}
企業名：${data.company}
提出期限：${data.deadlineStr}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

報告書の提出がまだお済みでない場合は、
お早めにアップロードをお願いいたします。

※依頼メールに記載のアップロードリンクからご提出ください。

ご不明な点がございましたら、お気軽にお問い合わせください。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}
