/**
 * 通知処理（LINE Messaging API、担当者個別LINE DM、Notion）
 */

/**
 * LINE Messaging API で通知送信（グループ向け）
 */
function sendLineNotification(data) {
  const token = CONFIG.LINE.CHANNEL_ACCESS_TOKEN;
  const groupId = CONFIG.LINE.GROUP_ID;

  if (!token || token === 'ここにチャネルアクセストークンを入力') {
    console.log('LINE Messaging API: トークン未設定のためスキップ');
    return;
  }

  if (!groupId || groupId === 'ここにグループIDまたはユーザーIDを入力') {
    console.log('LINE Messaging API: グループID未設定のためスキップ');
    return;
  }

  const message = `📋 新規相談申込

申込ID: ${data.id}
お名前: ${data.name}様
貴社名: ${data.company}
相談テーマ: ${data.theme}
希望日時: ${data.date1}
相談方法: ${data.method}
${data.companyUrl ? '企業URL: ' + data.companyUrl : ''}
スプレッドシートを確認してください。`;

  sendLineMessage(groupId, message);
}

/**
 * LINE Messaging API でステータス変更通知
 */
function sendLineStatusNotification(data, newStatus) {
  const token = CONFIG.LINE.CHANNEL_ACCESS_TOKEN;
  const groupId = CONFIG.LINE.GROUP_ID;

  if (!token || token === 'ここにチャネルアクセストークンを入力') {
    return;
  }

  if (!groupId || groupId === 'ここにグループIDまたはユーザーIDを入力') {
    return;
  }

  let emoji = '📝';
  if (newStatus === STATUS.CONFIRMED) emoji = '✅';
  if (newStatus === STATUS.COMPLETED) emoji = '🎉';
  if (newStatus === STATUS.CANCELLED) emoji = '❌';

  const message = `${emoji} ステータス更新

申込ID: ${data.id}
お名前: ${data.name}様
新ステータス: ${newStatus}`;

  sendLineMessage(groupId, message);
}

/**
 * LINE Messaging API でメッセージ送信
 * @param {string} to - 送信先ID（グループID or ユーザーID）
 * @param {string} text - メッセージ本文
 */
function sendLineMessage(to, text) {
  const token = CONFIG.LINE.CHANNEL_ACCESS_TOKEN;

  if (!token || token === 'ここにチャネルアクセストークンを入力') {
    console.log('LINE: トークン未設定のためスキップ');
    return false;
  }

  if (!to || to === 'ここにグループIDまたはユーザーIDを入力') {
    console.log('LINE: 送信先未設定のためスキップ');
    return false;
  }

  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: to,
    messages: [
      {
        type: 'text',
        text: text
      }
    ]
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      console.log('LINE通知送信成功');
      return true;
    } else {
      console.error('LINE通知送信エラー:', response.getContentText());
      return false;
    }
  } catch (e) {
    console.error('LINE通知送信エラー:', e);
    return false;
  }
}

/**
 * 担当者個別LINE DM送信
 * LINE IDが設定されている場合はLINEで送信、未設定の場合はメールにフォールバック
 * @param {string} staffName - 担当者名
 * @param {string} lineMessage - LINEメッセージ本文
 * @param {string} emailSubject - メール件名（フォールバック用）
 * @param {string} emailBody - メール本文（フォールバック用）
 */
function sendStaffNotification(staffName, lineMessage, emailSubject, emailBody) {
  const lineId = getStaffLineId(staffName);
  const staffEmail = getStaffEmail(staffName);

  // LINE優先で送信
  if (lineId) {
    const lineSuccess = sendLineMessage(lineId, lineMessage);
    if (lineSuccess) {
      console.log(`担当者 ${staffName} へのLINE DM送信成功`);
      return;
    }
    console.log(`担当者 ${staffName} へのLINE DM送信失敗、メールにフォールバック`);
  }

  // メールフォールバック
  if (staffEmail) {
    GmailApp.sendEmail(staffEmail, emailSubject, emailBody, {
      name: CONFIG.SENDER_NAME
    });
    console.log(`担当者 ${staffName} へのメール送信成功`);
  } else {
    console.log(`担当者 ${staffName}: LINE ID・メールアドレスともに未設定`);
  }
}

/**
 * 複数担当者への一括通知
 * カンマ区切りの担当者名に対して個別通知
 * @param {string} staffNames - カンマ区切りの担当者名
 * @param {string} lineMessage - LINEメッセージ本文
 * @param {string} emailSubject - メール件名
 * @param {string} emailBody - メール本文
 */
function sendStaffNotifications(staffNames, lineMessage, emailSubject, emailBody) {
  if (!staffNames) return;

  const names = staffNames.split(',').map(n => n.trim()).filter(n => n);
  names.forEach(name => {
    sendStaffNotification(name, lineMessage, emailSubject, emailBody);
  });

  // グループにも一括通知
  sendLineMessage(CONFIG.LINE.GROUP_ID, lineMessage);
}

/**
 * Notionにエントリを作成
 */
function createNotionEntry(data) {
  if (!CONFIG.NOTION.ENABLED) {
    return;
  }

  const url = 'https://api.notion.com/v1/pages';

  const payload = {
    parent: {
      database_id: CONFIG.NOTION.DATABASE_ID
    },
    properties: {
      '申込ID': {
        title: [
          {
            text: {
              content: data.id
            }
          }
        ]
      },
      'お名前': {
        rich_text: [
          {
            text: {
              content: data.name
            }
          }
        ]
      },
      '貴社名': {
        rich_text: [
          {
            text: {
              content: data.company
            }
          }
        ]
      },
      'メールアドレス': {
        email: data.email
      },
      '電話番号': {
        phone_number: data.phone
      },
      '相談テーマ': {
        select: {
          name: data.theme
        }
      },
      'ステータス': {
        select: {
          name: STATUS.PENDING
        }
      },
      '相談方法': {
        select: {
          name: data.method
        }
      },
      '希望日時': {
        rich_text: [
          {
            text: {
              content: data.date1
            }
          }
        ]
      }
    }
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.NOTION.API_KEY,
      'Content-Type': 'application/json',
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    console.log('Notion登録成功:', response.getContentText());
  } catch (e) {
    console.error('Notion登録エラー:', e);
  }
}

/**
 * Notionのステータスを更新
 */
function updateNotionStatus(pageId, newStatus) {
  if (!CONFIG.NOTION.ENABLED) {
    return;
  }

  const url = `https://api.notion.com/v1/pages/${pageId}`;

  const payload = {
    properties: {
      'ステータス': {
        select: {
          name: newStatus
        }
      }
    }
  };

  const options = {
    method: 'patch',
    headers: {
      'Authorization': 'Bearer ' + CONFIG.NOTION.API_KEY,
      'Content-Type': 'application/json',
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payload)
  };

  try {
    UrlFetchApp.fetch(url, options);
  } catch (e) {
    console.error('Notionステータス更新エラー:', e);
  }
}

/**
 * LINE Messaging API テスト
 */
function testLineMessage() {
  const testData = {
    id: 'TEST-001',
    name: 'テスト太郎',
    company: 'テスト株式会社',
    theme: '経営全般',
    date1: '2026-02-15 10:00',
    method: 'オンライン',
    companyUrl: ''
  };

  sendLineNotification(testData);
}
