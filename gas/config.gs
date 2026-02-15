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
  STATUS: 13,        // N: ステータス
  STAFF: 14,         // O: 担当者
  CONFIRMED_DATE: 15,// P: 確定日時
  ZOOM_URL: 16,      // Q: ZoomURL
  HEARING_SHEET: 17, // R: ヒアリングシート
  NOTES: 18,         // S: 備考
  NDA_STATUS: 19,    // T: 同意書同意（済/未）
  NDA_DATE: 20,      // U: 同意日時
  COMPANY_URL: 21,   // V: 企業URL
  WALK_IN_FLAG: 22   // W: 当日受付フラグ
};

/**
 * ステータス定義
 */
const STATUS = {
  PENDING: '仮予約',
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
  NOTES: 7      // H: 備考
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
