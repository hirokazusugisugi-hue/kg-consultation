/**
 * メールテンプレート（拡張版）
 * 同意書ページHTML、リマインド3日前/2日前、担当者向けテンプレート、当日受付用を追加
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
ご連絡先：${data.email} / ${data.phone}
ご希望日時：${data.date1}${data.date2 ? '\n第二希望：' + data.date2 : ''}
相談方法：${data.method}
相談テーマ：${data.theme}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${consentSection}
【次のステップ】
1. 上記リンクから相談同意書をご確認・ご同意ください
2. 添付の「ヒアリングシート」にご記入の上、
   相談日の3営業日前までにご返送ください

返送先：${CONFIG.REPLY_TO}
件名：【ヒアリングシート返送】${data.id} ${data.name}

ヒアリングシートをご返送いただき次第、
担当者より日程確定のご連絡をさせていただきます。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご注意事項
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
・本メールは自動送信です
・同意書へのご同意およびヒアリングシートのご返送をもって予約受付となります
・日程の変更・キャンセルは上記メールアドレスまでご連絡ください

ご不明な点がございましたら、お気軽にお問い合わせください。

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
ヒアリングシート受領後、日程を調整し、
ステータスを「確定」に変更すると確定メールが自動送信されます。`;
}

/**
 * 予約確定メール（申込者向け）- 企業URL・事前リサーチ案内付き
 */
function getConfirmedEmailBody(data) {
  let locationInfo = '';

  if (data.method === 'オンライン' || data.method === 'オンライン（Zoom）') {
    locationInfo = `【オンライン相談】
Zoom URL：${data.zoomUrl || '（後日ご連絡いたします）'}

※開始時刻の5分前を目安にご参加ください
※接続に不具合がある場合はお電話にてご連絡ください`;
  } else {
    locationInfo = `【対面相談】
場所：関西学院大学 西宮上ケ原キャンパス
    （詳細は別途ご案内いたします）

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
   - ヒアリングシートをもとにお話を伺います

2. 課題の整理・ディスカッション（30〜45分）
   - 課題を整理し、解決の方向性を一緒に考えます

3. 今後のアクション整理（15分程度）
   - 次のステップを明確にします

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ ご準備いただくもの
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
・ヒアリングシートの控え
・関連資料（決算書、事業計画書等）があればお持ちください
  ※必須ではありません

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ キャンセル・変更について
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
ご都合が悪くなった場合は、できるだけ早めにご連絡ください。
連絡先：${CONFIG.ORG.EMAIL}

ご不明な点がございましたら、お気軽にお問い合わせください。
当日お会いできることを楽しみにしております。

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
・ヒアリングシートの控えをお手元にご用意ください
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
 * リマインドメール（2日前・最終確認）
 */
function getReminderEmail2DaysBefore(data) {
  return `${data.name} 様

明後日のご相談について、最終確認のご連絡です。

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
・変更・キャンセルの場合は本日中にご連絡ください
${data.method === 'オンライン' || data.method === 'オンライン（Zoom）' ? '・Zoomの接続テストを事前にお願いいたします' : '・当日は受付にて「中小企業経営診断研究会の相談予約」とお伝えください'}

当日お会いできることを楽しみにしております。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
${CONFIG.ORG.NAME}
Email: ${CONFIG.ORG.EMAIL}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
}

/**
 * 担当者向けLINEリマインドメッセージ
 */
function getStaffReminderLine(data, daysBeforeLabel) {
  return `📋 ${daysBeforeLabel}リマインド

申込ID: ${data.id}
お名前: ${data.name}様
貴社名: ${data.company}
日時: ${data.confirmedDate}
方法: ${data.method}
テーマ: ${data.theme}
${data.companyUrl ? '企業URL: ' + data.companyUrl : ''}
事前準備をお願いします。`;
}

/**
 * 担当者向けメールリマインド（フォールバック用）
 */
function getStaffReminderEmail(data, daysBeforeLabel) {
  return `【${daysBeforeLabel}】担当相談のリマインド

${daysBeforeLabel}に以下の相談が予定されています。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 相談内容
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
申込ID：${data.id}
お名前：${data.name}様
貴社名：${data.company}
日時：${data.confirmedDate}
相談方法：${data.method}
テーマ：${data.theme}
${data.companyUrl ? '企業URL：' + data.companyUrl + '\n※事前リサーチにご活用ください' : ''}
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
  const pdfFileId = CONFIG.CONSENT.PDF_FILE_ID;
  const pdfViewerUrl = 'https://drive.google.com/file/d/' + pdfFileId + '/preview';
  const pdfDownloadUrl = 'https://drive.google.com/uc?export=download&id=' + pdfFileId;

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
      <p style="margin-top: 1rem; color: #666;">担当者より日程確定のご連絡をいたします。<br>このページを閉じていただいて結構です。</p>
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
    <p>${data.name} 様（申込ID: ${data.id}）<br>相談同意書への同意は既に受領済みです。<br>担当者より日程確定のご連絡をいたします。</p>
  </div>
</body>
</html>`;
}
