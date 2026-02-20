/**
 * 設定ファイル
 * ※実際の運用時に値を設定してください
 */

const CONFIG = {
  // スプレッドシートID（URLの/d/の後の部分）
  SPREADSHEET_ID: '1tD6-0WVXsud8_APnR4IxlI7yR085I47WoTjiFZ8oFBc',

  // シート名
  SHEET_NAME: '予約管理',

  // メンバーマスタシート名
  MEMBER_SHEET_NAME: 'メンバーマスタ',

  // 日程設定シート名
  SCHEDULE_SHEET_NAME: '日程設定',

  // 担当者メールアドレス（カンマ区切りで複数可）
  ADMIN_EMAILS: [
    'kgibaconsultant@gmail.com'
  ],

  // 送信元表示名
  SENDER_NAME: '関西学院大学 中小企業経営診断研究会',

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // LINE Messaging API 設定
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  LINE: {
    // チャネルアクセストークン（LINE Developersで取得）
    CHANNEL_ACCESS_TOKEN: 'ここにチャネルアクセストークンを入力',

    // 通知先グループID または ユーザーID
    GROUP_ID: 'ここにグループIDまたはユーザーIDを入力'
  },

  // Zoom API 設定（Server-to-Server OAuth）
  // ※認証情報はGASエディタの「スクリプト プロパティ」に設定してください
  // ZOOM_ACCOUNT_ID, ZOOM_CLIENT_ID, ZOOM_CLIENT_SECRET
  ZOOM: {
    ACCOUNT_ID: '',
    CLIENT_ID: '',
    CLIENT_SECRET: ''
  },

  // ヒアリングシートのGoogle DriveファイルID
  HEARING_SHEET_FILE_ID: 'ここにヒアリングシートのファイルIDを入力',

  // 同意書ページ設定
  CONSENT: {
    // GAS WebアプリのURL（デプロイ後に設定）
    WEB_APP_URL: 'https://script.google.com/macros/s/AKfycbzR7l1lyRF9dNZ0qqIov8LZwxDvkkyT4NNo2LSJKbQR_i46iqLfSRg4EuqQRflP76elAg/exec',
    // Google DriveにアップロードしたPDFのファイルID
    PDF_FILE_ID: '1X_q73Mgns_MKhku_b16x6oysnVs3dnP9',
    // PDF直接URL（GitHub raw）
    PDF_URL: 'https://raw.githubusercontent.com/hirokazusugisugi-hue/kg-consultation/main/consent_updated.pdf'
  },

  // 当日受付コード
  WALK_IN_CODE: '9999',

  // Notion連携（オプション）
  NOTION: {
    ENABLED: false,
    API_KEY: '',
    DATABASE_ID: ''
  },

  // フォームポーリング設定（v2: オプトアウト方式・2回送信）
  FORM_POLLING: {
    FIRST_SEND_DAY: 25,       // 1回目送信日（毎月N日）
    FIRST_SEND_HOUR: 10,      // 1回目送信時刻
    REMINDER_SEND_DAY: 15,    // 2回目（確認）送信日（翌月N日）
    REMINDER_SEND_HOUR: 10,   // 2回目送信時刻
    FINALIZE_DAY: 21,         // 日程確定日（翌月N日）
    FINALIZE_HOUR: 10,        // 日程確定時刻
    FIRST_DEADLINE_DAYS: 18,  // 1回目の回答期限（25日→翌月15日の約18日後）
    FINAL_DEADLINE_DAYS: 5    // 2回目の最終期限（送信からN日後）
  },

  // お知らせシート名
  NEWS_SHEET_NAME: 'お知らせ',

  // 回答集計シート名
  SUMMARY_SHEET_NAME: '回答集計',

  // アンケートシート名
  SURVEY_SHEET_NAME: 'アンケート',

  // オブザーバーNDAシート名
  OBSERVER_NDA_SHEET_NAME: 'オブザーバーNDA',

  // オブザーバーNDA設定
  OBSERVER_NDA: {
    // 署名済みNDA保存先DriveフォルダID（未設定時はルート）
    DRIVE_FOLDER_ID: '',
    // NDA PDFテンプレートURL
    PDF_URL: 'https://raw.githubusercontent.com/hirokazusugisugi-hue/kg-consultation/main/observer_nda.pdf'
  },

  // アンケート自動送信設定
  SURVEY: {
    DELAY_HOURS: 2  // 相談終了後N時間後に送信
  },

  // リーダー履歴シート名
  LEADER_HISTORY_SHEET_NAME: 'リーダー履歴',

  // レポート管理シート名
  REPORT_SHEET_NAME: 'レポート管理',

  // レポート配信設定
  REPORT: {
    DRIVE_FOLDER_ID: '',  // レポート保存先DriveフォルダID（未設定時はルート）
    DEADLINE_DAYS: 3,     // レポート提出期限（日数）
    MAX_FILE_SIZE: 5 * 1024 * 1024  // 5MB
  },

  // 返信先メールアドレス
  REPLY_TO: 'kgibaconsultant@gmail.com',

  // 組織情報
  ORG: {
    NAME: '関西学院大学 中小企業経営診断研究会',
    EMAIL: 'kgibaconsultant@gmail.com',
    URL: 'https://example.com'
  }
};

/**
 * Driveフォルダを取得（ScriptProperties優先、CONFIG fallback）
 * @param {string} propKey - ScriptPropertiesのキー
 * @param {string} configFallback - CONFIGの値（フォールバック）
 * @returns {Folder} Google Driveフォルダ
 */
function getDriveFolder(propKey, configFallback) {
  var folderId = PropertiesService.getScriptProperties().getProperty(propKey);
  if (!folderId && configFallback) folderId = configFallback;
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) { /* fallback to root */ }
  }
  return DriveApp.getRootFolder();
}

/**
 * スプレッドシートの列定義
 */
const COLUMNS = {
  TIMESTAMP: 0,      // A: タイムスタンプ
  ID: 1,             // B: 申込ID
  NAME: 2,           // C: お名前
  COMPANY: 3,        // D: 貴社名
  EMAIL: 4,          // E: メールアドレス
  PHONE: 5,          // F: 電話番号
  POSITION: 6,       // G: 役職
  INDUSTRY: 7,       // H: 業種
  THEME: 8,          // I: 相談テーマ
  CONTENT: 9,        // J: 相談内容
  DATE1: 10,         // K: 希望日時1
  DATE2: 11,         // L: 希望日時2
  METHOD: 12,        // M: 相談方法
  LOCATION: 13,      // N: 場所
  STATUS: 14,        // O: ステータス
  STAFF: 15,         // P: 担当者
  CONFIRMED_DATE: 16,// Q: 確定日時
  ZOOM_URL: 17,      // R: ZoomURL
  HEARING_SHEET: 18, // S: ヒアリングシート
  NOTES: 19,         // T: 備考
  NDA_STATUS: 20,    // U: 同意書同意（済/未）
  NDA_DATE: 21,      // V: 同意日時
  COMPANY_URL: 22,   // W: 企業URL
  WALK_IN_FLAG: 23,  // X: 当日受付フラグ
  LEADER: 24,        // Y: リーダー
  REPORT_STATUS: 25  // Z: レポート状態
};

/**
 * 場所の選択肢
 */
const LOCATION_OPTIONS = [
  'アプローズタワー',
  'スミセスペース',
  'ナレッジサロン',
  'その他'
];

/**
 * ステータス定義
 */
const STATUS = {
  PENDING: '仮予約',
  NDA_AGREED: 'NDA同意済',
  RECEIVED: '書類受領',
  CONFIRMED: '確定',
  COMPLETED: '完了',
  CANCELLED: 'キャンセル'
};

/**
 * メンバーマスタの列定義
 */
const MEMBER_COLUMNS = {
  NAME: 0,      // A: 氏名
  TERM: 1,      // B: 期
  CERT: 2,      // C: 資格
  TYPE: 3,      // D: 区分
  EMAIL: 4,     // E: メール
  PHONE: 5,     // F: 電話番号
  LINE_ID: 6,   // G: LINE ID
  NOTES: 7,         // H: 備考
  SPECIALTIES: 8,   // I: 得意業種
  THEMES: 9         // J: 得意テーマ
};

/**
 * 日程設定シートの列定義（拡張版）
 */
const SCHEDULE_COLUMNS = {
  DATE: 0,           // A: 日付
  TIME: 1,           // B: 時間帯
  AVAILABLE: 2,      // C: 対応可否
  METHOD: 3,         // D: 対応方法
  STAFF: 4,          // E: 担当者
  BOOKING_STATUS: 5, // F: 予約状況
  NOTES: 6,          // G: 備考
  MEMBERS: 7,        // H: 参加メンバー（カンマ区切り）
  SCORE: 8,          // I: 配置点数
  SPECIAL_FLAG: 9,   // J: 特別対応フラグ
  BOOKABLE: 10       // K: 予約可能判定
};

/**
 * お知らせシートの列定義
 */
const NEWS_COLUMNS = {
  DATE: 0,       // A: 日付
  CONTENT: 1,    // B: 内容
  VISIBLE: 2,    // C: 表示フラグ
  CREATED_AT: 3  // D: 作成日時
};

/**
 * アンケートシートの列定義
 */
const SURVEY_COLUMNS = {
  TIMESTAMP: 0,    // A: 回答日時
  APP_ID: 1,       // B: 申込ID
  NAME: 2,         // C: 氏名
  COMPANY: 3,      // D: 企業名
  Q1: 4,           // E: きっかけ（複数選択）
  Q1_SNS: 5,       // F: SNS種別
  Q2: 6,           // G: 手続きスムーズ（5段階）
  Q2_COMMENT: 7,   // H: Q2コメント
  Q3: 8,           // I: 感想（自由記述）
  Q4: 9,           // J: 時間（5択）
  Q5: 10,          // K: 説明わかりやすさ（5段階）
  Q6: 11,          // L: 課題解決参考（5段階）
  Q7: 12,          // M: 対応誠実（5段階）
  Q8: 13,          // N: 行動アドバイス（5段階）
  Q9: 14,          // O: また受けたい（5段階）
  Q9_REASON: 15,   // P: Q9理由
  Q10: 16,         // Q: 勧めたい（5段階）
  Q10_REASON: 17,  // R: Q10理由
  Q11: 18          // S: レポート希望
};

/**
 * オブザーバーNDA管理の列定義
 */
const OBSERVER_NDA_COLUMNS = {
  TIMESTAMP: 0,      // A: 提出日時
  OBSERVER_NAME: 1,  // B: オブザーバー氏名
  CONSULT_DATE: 2,   // C: 相談日
  COMPANY: 3,        // D: 相談企業名
  STAFF: 4,          // E: 相談担当者
  FILE_ID: 5,        // F: Drive上のファイルID
  FILE_URL: 6        // G: ダウンロードURL
};

/**
 * レポート状態定数
 */
const REPORT_STATUS = {
  NOT_REQUESTED: '未依頼',
  REQUESTED: '依頼済',
  UPLOADED: 'アップロード済',
  DELIVERED: '配信済',
  OVERDUE: '期限超過'
};

/**
 * リーダー履歴シートの列定義
 */
const LEADER_HISTORY_COLUMNS = {
  STATUS: 0,         // A: ステータス
  CONSULT_DATE: 1,   // B: 相談日時
  APP_ID: 2,         // C: 申込ID
  COMPANY: 3,        // D: 相談企業
  LEADER: 4,         // E: リーダー
  INDUSTRY: 5,       // F: 業種
  THEME: 6,          // G: テーマ
  MEMBERS: 7,        // H: 参加メンバー
  MATCH_SCORE: 8,    // I: マッチスコア
  REASON: 9,         // J: 選定理由
  DATE: 10           // K: 登録日
};

/**
 * レポート管理シートの列定義
 */
const REPORT_COLUMNS = {
  APP_ID: 0,         // A: 申込ID
  CONSULT_DATE: 1,   // B: 相談日
  COMPANY: 2,        // C: 相談企業
  LEADER: 3,         // D: リーダー
  LEADER_EMAIL: 4,   // E: リーダーメール
  REQUEST_DATE: 5,   // F: 依頼日時
  DEADLINE: 6,       // G: 期限
  UPLOAD_DATE: 7,    // H: アップロード日時
  FILE_ID: 8,        // I: ファイルID
  FILE_URL: 9,       // J: ファイルURL
  DELIVERY_DATE: 10, // K: 配信日時
  STATUS: 11         // L: ステータス
};
