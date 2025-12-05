/**
 * 古本買取価格調査システム
 * Main Configuration and Constants
 * 
 * ファイル名: Config.gs
 */

// =================================/**
 * 古本買取価格調査システム
 * Utility Functions
 * 
 * ファイル名: Utils.gs
 */

// ========================================
// HTTPリクエスト関連
// ========================================

/**
 * HTTPリクエストを実行（リトライ機能付き）
 * @param {string} url - リクエストURL
 * @param {Object} options - リクエストオプション
 * @returns {HTTPResponse} レスポンス
 */
function fetchWithRetry(url, options = {}) {
  const maxRetries = CONFIG.SCRAPING.MAX_RETRIES;
  const retryInterval = CONFIG.SCRAPING.RETRY_INTERVAL;
  
  // デフォルトオプション - より人間らしいブラウザヘッダー
  const defaultOptions = {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
      'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8',
      'Accept-Encoding': 'gzip, deflate, br',
      'Connection': 'keep-alive',
      'Upgrade-Insecure-Requests': '1',
      'Sec-Fetch-Dest': 'document',
      'Sec-Fetch-Mode': 'navigate',
      'Sec-Fetch-Site': 'none',
      'Cache-Control': 'max-age=0'
    }
  };
  
  const requestOptions = { ...defaultOptions, ...options };
  
  for (let i = 0; i < maxRetries; i++) {
    try {
      logInfo(`リクエスト送信: ${url} (試行 ${i + 1}/${maxRetries})`);
      const response = UrlFetchApp.fetch(url, requestOptions);
      const statusCode = response.getResponseCode();
      
      if (statusCode === 200) {
        logInfo(`リクエスト成功: ${url}`);
        return response;
      } else if (statusCode === 403) {
        logWarning(`アクセス拒否(403): ${url} - Botとして判断された可能性があります`);
        if (i < maxRetries - 1) {
          // 403エラーの場合は少し長めに待つ
          sleep(retryInterval * 2);
        }
      } else {
        logWarning(`HTTPエラー: ${statusCode} - ${url}`);
        if (i < maxRetries - 1) {
          sleep(retryInterval);
        }
      }
    } catch (error) {
      logError(`リクエストエラー (試行 ${i + 1}/${maxRetries}): ${error.message}`);
      if (i < maxRetries - 1) {
        sleep(retryInterval);
      } else {
        throw new Error(`リクエスト失敗: ${error.message}`);
      }
    }
  }
  
  throw new Error(`リクエスト失敗: 最大リトライ回数に到達しました`);
}

/**
 * HTMLからテキストを抽出（正規表現使用）
 * @param {string} html - HTML文字列
 * @param {string} pattern - 抽出パターン（正規表現）
 * @param {number} groupIndex - キャプチャグループのインデックス
 * @returns {string|null} 抽出したテキスト
 */
function extractTextFromHtml(html, pattern, groupIndex = 1) {
  try {
    const regex = new RegExp(pattern, 's');
    const match = html.match(regex);
    return match ? match[groupIndex] : null;
  } catch (error) {
    logError(`HTML抽出エラー: ${error.message}`);
    return null;
  }
}

/**
 * HTMLエンティティをデコード
 * @param {string} text - テキスト
 * @returns {string} デコードされたテキスト
 */
function decodeHtmlEntities(text) {
  if (!text) return text;
  
  const entities = {
    '&amp;': '&',
    '&lt;': '<',
    '&gt;': '>',
    '&quot;': '"',
    '&#39;': "'",
    '&nbsp;': ' '
  };
  
  let decoded = text;
  for (const [entity, char] of Object.entries(entities)) {
    decoded = decoded.replace(new RegExp(entity, 'g'), char);
  }
  
  return decoded;
}

// ========================================
// スプレッドシート操作関連
// ========================================

/**
 * シートの最終行を取得
 * @param {Sheet} sheet - シート
 * @returns {number} 最終行番号
 */
function getLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  return lastRow > 1 ? lastRow : 1; // 最低でも1を返す（ヘッダー行）
}

/**
 * シートの最終列を取得
 * @param {Sheet} sheet - シート
 * @returns {number} 最終列番号
 */
function getLastColumn(sheet) {
  const lastColumn = sheet.getLastColumn();
  return lastColumn > 0 ? lastColumn : 1;
}

/**
 * シートの全データを取得（ヘッダー除く）
 * @param {Sheet} sheet - シート
 * @returns {Array<Array>} データ配列
 */
function getSheetData(sheet) {
  const lastRow = getLastRow(sheet);
  if (lastRow <= 1) {
    return [];
  }
  
  const lastColumn = getLastColumn(sheet);
  return sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
}

/**
 * 指定した列の値を取得
 * @param {Sheet} sheet - シート
 * @param {number} columnIndex - 列番号（1始まり）
 * @returns {Array} 列の値の配列
 */
function getColumnValues(sheet, columnIndex) {
  const lastRow = getLastRow(sheet);
  if (lastRow <= 1) {
    return [];
  }
  
  return sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues().flat();
}

/**
 * 条件に合致する行を検索
 * @param {Sheet} sheet - シート
 * @param {number} columnIndex - 検索する列番号
 * @param {*} value - 検索値
 * @returns {number|null} 行番号（見つからない場合はnull）
 */
function findRow(sheet, columnIndex, value) {
  const values = getColumnValues(sheet, columnIndex);
  const index = values.indexOf(value);
  return index >= 0 ? index + 2 : null; // +2はヘッダー行とインデックス調整
}

/**
 * 行データを取得
 * @param {Sheet} sheet - シート
 * @param {number} rowIndex - 行番号
 * @returns {Array} 行データ
 */
function getRowData(sheet, rowIndex) {
  const lastColumn = getLastColumn(sheet);
  return sheet.getRange(rowIndex, 1, 1, lastColumn).getValues()[0];
}

/**
 * 行データを設定
 * @param {Sheet} sheet - シート
 * @param {number} rowIndex - 行番号
 * @param {Array} data - 設定するデータ
 */
function setRowData(sheet, rowIndex, data) {
  sheet.getRange(rowIndex, 1, 1, data.length).setValues([data]);
}

/**
 * セルの値を取得
 * @param {Sheet} sheet - シート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @returns {*} セルの値
 */
function getCellValue(sheet, row, col) {
  return sheet.getRange(row, col).getValue();
}

/**
 * セルの値を設定
 * @param {Sheet} sheet - シート
 * @param {number} row - 行番号
 * @param {number} col - 列番号
 * @param {*} value - 設定する値
 */
function setCellValue(sheet, row, col, value) {
  sheet.getRange(row, col).setValue(value);
}

/**
 * 行を追加
 * @param {Sheet} sheet - シート
 * @param {Array} data - 追加するデータ
 * @returns {number} 追加した行番号
 */
function appendRow(sheet, data) {
  sheet.appendRow(data);
  return getLastRow(sheet);
}

/**
 * 行を削除
 * @param {Sheet} sheet - シート
 * @param {number} rowIndex - 削除する行番号
 */
function deleteRow(sheet, rowIndex) {
  sheet.deleteRow(rowIndex);
}

/**
 * チェックボックスがチェックされている行を取得
 * @param {Sheet} sheet - シート
 * @param {number} checkboxColumn - チェックボックスの列番号
 * @returns {Array<number>} チェックされている行番号の配列
 */
function getCheckedRows(sheet, checkboxColumn) {
  const lastRow = getLastRow(sheet);
  if (lastRow <= 1) {
    return [];
  }
  
  const values = sheet.getRange(2, checkboxColumn, lastRow - 1, 1).getValues();
  const checkedRows = [];
  
  values.forEach((row, index) => {
    if (row[0] === true) {
      checkedRows.push(index + 2); // +2はヘッダー行とインデックス調整
    }
  });
  
  return checkedRows;
}

// ========================================
// 日付・時刻関連
// ========================================

/**
 * 指定日数前の日付を取得
 * @param {number} days - 日数
 * @returns {Date} 日付
 */
function getDaysAgo(days) {
  const date = new Date();
  date.setDate(date.getDate() - days);
  return date;
}

/**
 * 2つの日付の差分（日数）を取得
 * @param {Date} date1 - 日付1
 * @param {Date} date2 - 日付2
 * @returns {number} 差分日数
 */
function getDaysDifference(date1, date2) {
  const oneDay = 24 * 60 * 60 * 1000;
  return Math.round(Math.abs((date1 - date2) / oneDay));
}

/**
 * 日付が指定期間内かチェック
 * @param {Date} date - チェックする日付
 * @param {Date} startDate - 開始日
 * @param {Date} endDate - 終了日
 * @returns {boolean} 期間内かどうか
 */
function isDateInRange(date, startDate, endDate) {
  return date >= startDate && date <= endDate;
}

// ========================================
// データ処理関連
// ========================================

/**
 * 配列の重複を除去
 * @param {Array} array - 配列
 * @returns {Array} 重複除去後の配列
 */
function uniqueArray(array) {
  return [...new Set(array)];
}

/**
 * 配列をソート（降順）
 * @param {Array} array - 配列
 * @returns {Array} ソート済み配列
 */
function sortDescending(array) {
  return array.sort((a, b) => b - a);
}

/**
 * オブジェクト配列を指定キーでソート
 * @param {Array<Object>} array - オブジェクト配列
 * @param {string} key - ソートキー
 * @param {boolean} descending - 降順かどうか
 * @returns {Array<Object>} ソート済み配列
 */
function sortByKey(array, key, descending = false) {
  return array.sort((a, b) => {
    if (descending) {
      return b[key] - a[key];
    } else {
      return a[key] - b[key];
    }
  });
}

/**
 * 数値配列の合計を計算
 * @param {Array<number>} array - 数値配列
 * @returns {number} 合計
 */
function sum(array) {
  return array.reduce((acc, val) => acc + (val || 0), 0);
}

/**
 * 数値配列の平均を計算
 * @param {Array<number>} array - 数値配列
 * @returns {number} 平均
 */
function average(array) {
  if (!array || array.length === 0) return 0;
  return sum(array) / array.length;
}

/**
 * 空の値をチェック
 * @param {*} value - 値
 * @returns {boolean} 空かどうか
 */
function isEmpty(value) {
  return value === null || value === undefined || value === '';
}

/**
 * 全て空でないかチェック
 * @param {Array} values - 値の配列
 * @returns {boolean} 全て空でないかどうか
 */
function allNotEmpty(values) {
  return values.every(val => !isEmpty(val));
}=======
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