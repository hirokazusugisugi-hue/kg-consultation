/**
 * 通知処理（メール、Notion）
 */

/**
 * 担当者へメール通知
 * @param {string} staffName - 担当者名
 * @param {string} emailSubject - メール件名
 * @param {string} emailBody - メール本文
 */
function sendStaffNotification(staffName, emailSubject, emailBody) {
  const staffEmail = getStaffEmail(staffName);

  if (staffEmail) {
    GmailApp.sendEmail(staffEmail, emailSubject, emailBody, {
      name: CONFIG.SENDER_NAME
    });
    console.log(`担当者 ${staffName} へのメール送信成功`);
  } else {
    console.log(`担当者 ${staffName}: メールアドレス未設定`);
  }
}

/**
 * 複数担当者への一括メール通知
 * カンマ区切りの担当者名に対して個別通知
 * @param {string} staffNames - カンマ区切りの担当者名
 * @param {string} emailSubject - メール件名
 * @param {string} emailBody - メール本文
 */
function sendStaffNotifications(staffNames, emailSubject, emailBody) {
  if (!staffNames) return;

  const names = staffNames.split(',').map(n => n.trim()).filter(n => n);
  names.forEach(name => {
    sendStaffNotification(name, emailSubject, emailBody);
  });
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

