/**
 * 古本買取価格調査システム
 * Main Configuration and Constants
 * 
 * ファイル名: Config.gs
 */

// ========================================
// システム設定
// ========================================

const CONFIG = {
  // スプレッドシート設定
  SHEET_NAMES: {
    ISBN_LIST: 'ISBNリスト',
    COMPLETED: '買取完了',
    PRICE_HISTORY: '価格履歴',
    DASHBOARD: 'ダッシュボード',
    ERROR_LOG: 'エラーログ'
  },
  
  // スクレイピング設定
  SCRAPING: {
    VALUEBOOKS_BASE_URL: 'https://www.valuebooks.jp',
    VALUEBOOKS_SEARCH_URL: 'https://www.valuebooks.jp/estimate/guide',
    CHARIBON_NEWS_URL: 'https://www.charibon.jp/news/',
    REQUEST_INTERVAL: 2000,  // リクエスト間隔（ミリ秒）
    TIMEOUT: 30000,          // タイムアウト（ミリ秒）
    MAX_RETRIES: 3,          // 最大リトライ回数
    RETRY_INTERVAL: 5000     // リトライ間隔（ミリ秒）
  },
  
  // Google Books API設定
  GOOGLE_BOOKS: {
    API_ENDPOINT: 'https://www.googleapis.com/books/v1/volumes',
    // API Keyは後でスクリプトプロパティに設定
  },
  
  // 列番号定義（ISBNリストシート）
  ISBN_LIST_COLUMNS: {
    ISBN: 1,           // A列
    TITLE: 2,          // B列
    AUTHOR: 3,         // C列
    PUBLISHER: 4,      // D列
    PRICE: 5,          // E列
    UPDATED: 6,        // F列
    STATUS: 7,         // G列
    PRICE_CHANGE: 8,   // H列
    CHECKBOX: 9        // I列
  },
  
  // 列番号定義（買取完了シート）
  COMPLETED_COLUMNS: {
    ISBN: 1,           // A列
    TITLE: 2,          // B列
    AUTHOR: 3,         // C列
    PUBLISHER: 4,      // D列
    ESTIMATE: 5,       // E列
    ACTUAL: 6,         // F列
    DIFFERENCE: 7,     // G列
    DATE: 8            // H列
  },
  
  // 列番号定義（価格履歴シート）
  HISTORY_COLUMNS: {
    ISBN: 1,           // A列
    TITLE: 2,          // B列
    DATETIME: 3,       // C列
    PRICE: 4,          // D列
    CHANGE: 5          // E列
  },
  
  // 列番号定義（エラーログシート）
  ERROR_COLUMNS: {
    DATETIME: 1,       // A列
    TYPE: 2,           // B列
    ISBN: 3,           // C列
    MESSAGE: 4,        // D列
    STATUS: 5          // E列
  },
  
  // ステータス定義
  STATUS: {
    NOT_SOLD: '未買取',
    SOLD: '買取済'
  },
  
  // エラー種別
  ERROR_TYPES: {
    SCRAPING_FAILED: 'スクレイピング失敗',
    API_FAILED: 'API失敗',
    DATA_ERROR: 'データエラー',
    NETWORK_ERROR: 'ネットワークエラー',
    OTHER: 'その他'
  },
  
  // エラーログステータス
  ERROR_STATUS: {
    PENDING: '未対応',
    RESOLVED: '対応済'
  },
  
  // メール通知設定
  EMAIL: {
    SUBJECT_PREFIX: '[古本買取システム] ',
    // 送信先メールアドレスはスクリプトプロパティに設定
  },
  
  // ダッシュボード設定
  DASHBOARD: {
    TOP_ITEMS_COUNT: 10,  // トップ10表示
    PRICE_CHANGE_DAYS: 7  // 価格変動の集計日数
  }
};

// ========================================
// スプレッドシート取得関数
// ========================================

/**
 * アクティブなスプレッドシートを取得
 * @returns {Spreadsheet} スプレッドシート
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 指定したシートを取得
 * @param {string} sheetName - シート名
 * @returns {Sheet} シート
 */
function getSheet(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`シート "${sheetName}" が見つかりません`);
  }
  return sheet;
}

/**
 * 各シートを取得する関数群
 */
function getIsbnListSheet() {
  return getSheet(CONFIG.SHEET_NAMES.ISBN_LIST);
}

function getCompletedSheet() {
  return getSheet(CONFIG.SHEET_NAMES.COMPLETED);
}

function getPriceHistorySheet() {
  return getSheet(CONFIG.SHEET_NAMES.PRICE_HISTORY);
}

function getDashboardSheet() {
  return getSheet(CONFIG.SHEET_NAMES.DASHBOARD);
}

function getErrorLogSheet() {
  return getSheet(CONFIG.SHEET_NAMES.ERROR_LOG);
}

// ========================================
// スクリプトプロパティ管理
// ========================================

/**
 * スクリプトプロパティを取得
 * @param {string} key - プロパティキー
 * @returns {string|null} プロパティ値
 */
function getScriptProperty(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * スクリプトプロパティを設定
 * @param {string} key - プロパティキー
 * @param {string} value - プロパティ値
 */
function setScriptProperty(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, value);
}

/**
 * 通知先メールアドレスを取得
 * @returns {string} メールアドレス
 */
function getNotificationEmail() {
  let email = getScriptProperty('NOTIFICATION_EMAIL');
  if (!email) {
    // 未設定の場合は実行ユーザーのメールアドレスを使用
    email = Session.getActiveUser().getEmail();
    setScriptProperty('NOTIFICATION_EMAIL', email);
  }
  return email;
}

/**
 * Google Books API Keyを取得
 * @returns {string|null} API Key
 */
function getGoogleBooksApiKey() {
  return getScriptProperty('GOOGLE_BOOKS_API_KEY');
}

/**
 * 初期設定を行う（初回実行時）
 */
function initialSetup() {
  const email = Session.getActiveUser().getEmail();
  setScriptProperty('NOTIFICATION_EMAIL', email);
  
  Logger.log('初期設定が完了しました');
  Logger.log(`通知先メールアドレス: ${email}`);
  Logger.log('');
  Logger.log('【重要】Google Books APIを使用する場合:');
  Logger.log('1. Google Cloud Consoleでプロジェクトを作成');
  Logger.log('2. Books APIを有効化');
  Logger.log('3. APIキーを取得');
  Logger.log('4. 以下のコマンドを実行してAPIキーを設定:');
  Logger.log('   setScriptProperty("GOOGLE_BOOKS_API_KEY", "YOUR_API_KEY")');
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * 指定時間待機
 * @param {number} milliseconds - 待機時間（ミリ秒）
 */
function sleep(milliseconds) {
  Utilities.sleep(milliseconds);
}

/**
 * 現在日時を取得
 * @returns {Date} 現在日時
 */
function getCurrentDateTime() {
  return new Date();
}

/**
 * 日付を文字列にフォーマット
 * @param {Date} date - 日付
 * @returns {string} フォーマット済み文字列
 */
function formatDateTime(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
}

/**
 * 日付を日付のみの文字列にフォーマット
 * @param {Date} date - 日付
 * @returns {string} フォーマット済み文字列
 */
function formatDate(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd');
}

/**
 * ISBNの形式をチェック
 * @param {string} isbn - ISBN
 * @returns {boolean} 正しい形式かどうか
 */
function isValidISBN(isbn) {
  if (!isbn || typeof isbn !== 'string') {
    return false;
  }
  
  // ハイフンを除去
  const cleanIsbn = isbn.replace(/-/g, '');
  
  // ISBN-10 または ISBN-13 の形式チェック
  return /^\d{10}$/.test(cleanIsbn) || /^\d{13}$/.test(cleanIsbn);
}

/**
 * 数値を円表記にフォーマット
 * @param {number} value - 数値
 * @returns {string} 円表記の文字列
 */
function formatCurrency(value) {
  if (value === null || value === undefined || isNaN(value)) {
    return '¥0';
  }
  return '¥' + value.toLocaleString('ja-JP');
}

/**
 * テキストから数値を抽出
 * @param {string} text - テキスト
 * @returns {number|null} 抽出した数値
 */
function extractNumber(text) {
  if (!text) return null;
  const match = text.toString().match(/\d+/);
  return match ? parseInt(match[0]) : null;
}

// ========================================
// ログ出力関数
// ========================================

/**
 * 情報ログを出力
 * @param {string} message - メッセージ
 */
function logInfo(message) {
  Logger.log(`[INFO] ${message}`);
}

/**
 * エラーログを出力
 * @param {string} message - メッセージ
 */
function logError(message) {
  Logger.log(`[ERROR] ${message}`);
}

/**
 * 警告ログを出力
 * @param {string} message - メッセージ
 */
function logWarning(message) {
  Logger.log(`[WARNING] ${message}`);
}