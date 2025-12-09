/**
 * ダッシュボード用カスタム関数
 * 
 * ファイル名: Dashboard.gs
 */

/**
 * 総買取冊数を取得
 * @param {Date} lastUpdate - 最終更新日時（再計算トリガー用）
 * @returns {number} 総冊数
 * @customfunction
 */
function getTotalBuyCount() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    let totalCount = 0;
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      // "買取完了_" で始まるシート名のみ対象
      if (name.startsWith('買取完了_')) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          totalCount += lastRow - 1; // ヘッダー行を除く
        }
      }
    });
    
    return totalCount;
    
  } catch (error) {
    logError(`getTotalBuyCount エラー: ${error.message}`);
    return 0;
  }
}

/**
 * 全買取完了シートの総利益を取得
 * @param {Date} lastUpdate - 最終更新日時（再計算トリガー用）
 * @returns {number} 総利益
 * @customfunction
 */
function getTotalProfit() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    let totalProfit = 0;
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      // "買取完了_" で始まるシート名のみ対象
      if (name.startsWith('買取完了_')) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          // G列（利益）を取得
          const profits = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
          profits.forEach(row => {
            const profit = parseFloat(row[0]);
            if (!isNaN(profit)) {
              totalProfit += profit;
            }
          });
        }
      }
    });
    
    return Math.round(totalProfit);
    
  } catch (error) {
    logError(`getTotalProfit エラー: ${error.message}`);
    return 0;
  }
}

/**
 * 全買取完了シートの最高利益を取得
 * @param {Date} lastUpdate - 最終更新日時（再計算トリガー用）
 * @returns {number} 最高利益
 * @customfunction
 */
function getMaxProfit() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    let maxProfit = 0;
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      // "買取完了_" で始まるシート名のみ対象
      if (name.startsWith('買取完了_')) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          // G列（利益）を取得
          const profits = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
          profits.forEach(row => {
            const profit = parseFloat(row[0]);
            if (!isNaN(profit) && profit > maxProfit) {
              maxProfit = profit;
            }
          });
        }
      }
    });
    
    return Math.round(maxProfit);
    
  } catch (error) {
    logError(`getMaxProfit エラー: ${error.message}`);
    return 0;
  }
}

/**
 * 今月の買取冊数を取得
 * @param {Date} lastUpdate - 最終更新日時（再計算トリガー用）
 * @returns {number} 今月の買取冊数
 * @customfunction
 */
function getMonthlyBuyCount() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1; // 1-12
    
    let monthlyCount = 0;
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      // "買取完了_YYYY-MM-DD" 形式のシート名をチェック
      if (name.startsWith('買取完了_')) {
        const dateStr = name.replace('買取完了_', '');
        
        try {
          const sheetDate = new Date(dateStr);
          const sheetYear = sheetDate.getFullYear();
          const sheetMonth = sheetDate.getMonth() + 1;
          
          // 今月のシートのみカウント
          if (sheetYear === currentYear && sheetMonth === currentMonth) {
            const lastRow = sheet.getLastRow();
            if (lastRow > 1) {
              monthlyCount += lastRow - 1;
            }
          }
        } catch (e) {
          // 日付パースエラーは無視
        }
      }
    });
    
    return monthlyCount;
    
  } catch (error) {
    logError(`getMonthlyBuyCount エラー: ${error.message}`);
    return 0;
  }
}

/**
 * 今月の利益を取得
 * @param {Date} lastUpdate - 最終更新日時（再計算トリガー用）
 * @returns {number} 今月の利益
 * @customfunction
 */
function getMonthlyProfit() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;
    
    let monthlyProfit = 0;
    
    sheets.forEach(sheet => {
      const name = sheet.getName();
      // "買取完了_YYYY-MM-DD" 形式のシート名をチェック
      if (name.startsWith('買取完了_')) {
        const dateStr = name.replace('買取完了_', '');
        
        try {
          const sheetDate = new Date(dateStr);
          const sheetYear = sheetDate.getFullYear();
          const sheetMonth = sheetDate.getMonth() + 1;
          
          // 今月のシートのみ集計
          if (sheetYear === currentYear && sheetMonth === currentMonth) {
            const lastRow = sheet.getLastRow();
            if (lastRow > 1) {
              // G列（利益）を取得
              const profits = sheet.getRange(2, 7, lastRow - 1, 1).getValues();
              profits.forEach(row => {
                const profit = parseFloat(row[0]);
                if (!isNaN(profit)) {
                  monthlyProfit += profit;
                }
              });
            }
          }
        } catch (e) {
          // 日付パースエラーは無視
        }
      }
    });
    
    return Math.round(monthlyProfit);
    
  } catch (error) {
    logError(`getMonthlyProfit エラー: ${error.message}`);
    return 0;
  }
}

// ========================================
// 価格変動アラート用関数
// ========================================

/**
 * 価格変動データを取得
 * @returns {Array} 価格変動データの配列
 */
function getPriceChanges() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const isbnSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISBN_LIST);
    const historySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PRICE_HISTORY);
    
    if (!isbnSheet || !historySheet) {
      return [];
    }
    
    // ISBNリストから現在のデータを取得
    const isbnLastRow = isbnSheet.getLastRow();
    if (isbnLastRow < 2) {
      return [];
    }
    
    const isbnData = isbnSheet.getRange(2, 1, isbnLastRow - 1, 5).getValues();
    // [ISBN, タイトル, 著者, 出版社, 最新見積価格]
    
    // 価格履歴から初回価格を取得
    const historyLastRow = historySheet.getLastRow();
    if (historyLastRow < 2) {
      return [];
    }
    
    const historyData = historySheet.getRange(2, 1, historyLastRow - 1, 4).getValues();
    // [ISBN, タイトル, 更新日時, 価格]
    
    // ISBNごとの初回価格を取得
    const firstPrices = {};
    historyData.forEach(row => {
      const isbn = String(row[0]).trim();
      const price = parseFloat(row[3]);
      
      if (isbn && !isNaN(price)) {
        // まだ記録されていないか、より古い日時の場合
        if (!firstPrices[isbn]) {
          firstPrices[isbn] = {
            title: row[1],
            price: price,
            datetime: row[2]
          };
        } else {
          // より古い日時のデータで更新
          if (row[2] < firstPrices[isbn].datetime) {
            firstPrices[isbn] = {
              title: row[1],
              price: price,
              datetime: row[2]
            };
          }
        }
      }
    });
    
    // 価格変動を計算
    const priceChanges = [];
    
    isbnData.forEach(row => {
      const isbn = String(row[0]).trim();
      const title = row[1];
      const latestPrice = parseFloat(row[4]);
      
      if (!isbn || !title) {
        return;
      }
      
      // 最新価格が数値でない場合はスキップ
      if (isNaN(latestPrice)) {
        return;
      }
      
      // 初回価格が存在する場合のみ計算
      if (firstPrices[isbn]) {
        const firstPrice = firstPrices[isbn].price;
        const change = latestPrice - firstPrice;
        
        // 変動がある場合のみ追加
        if (change !== 0) {
          priceChanges.push({
            isbn: isbn,
            title: title,
            firstPrice: firstPrice,
            latestPrice: latestPrice,
            change: change
          });
        }
      }
    });
    
    return priceChanges;
    
  } catch (error) {
    logError(`getPriceChanges エラー: ${error.message}`);
    return [];
  }
}

/**
 * 価格上昇TOP5を取得
 * @returns {Array} [タイトル, 初回価格, 最新価格, 変動額]の配列
 * @customfunction
 */
function getPriceIncreasesTop5() {
  try {
    const priceChanges = getPriceChanges();
    
    // 上昇のみフィルタ（change > 0）
    const increases = priceChanges.filter(item => item.change > 0);
    
    // 変動額の大きい順にソート
    increases.sort((a, b) => b.change - a.change);
    
    // TOP 5を取得
    const top5 = increases.slice(0, 5);
    
    // 表示用の配列に変換
    const result = top5.map(item => [
      item.title,
      item.firstPrice,
      item.latestPrice,
      item.change
    ]);
    
    // 5件未満の場合は空行で埋める
    while (result.length < 5) {
      result.push(['', '', '', '']);
    }
    
    return result;
    
  } catch (error) {
    logError(`getPriceIncreasesTop5 エラー: ${error.message}`);
    return [['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', '']];
  }
}

/**
 * 価格下落TOP5を取得
 * @returns {Array} [タイトル, 初回価格, 最新価格, 変動額]の配列
 * @customfunction
 */
function getPriceDecreasesTop5() {
  try {
    const priceChanges = getPriceChanges();
    
    // 下落のみフィルタ（change < 0）
    const decreases = priceChanges.filter(item => item.change < 0);
    
    // 変動額の大きい順にソート（絶対値）
    decreases.sort((a, b) => a.change - b.change);
    
    // TOP 5を取得
    const top5 = decreases.slice(0, 5);
    
    // 表示用の配列に変換
    const result = top5.map(item => [
      item.title,
      item.firstPrice,
      item.latestPrice,
      item.change
    ]);
    
    // 5件未満の場合は空行で埋める
    while (result.length < 5) {
      result.push(['', '', '', '']);
    }
    
    return result;
    
  } catch (error) {
    logError(`getPriceDecreasesTop5 エラー: ${error.message}`);
    return [['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', ''], ['', '', '', '']];
  }
}

/**
 * 0円になった書籍を取得
 * @returns {Array} [タイトル, ISBN]の配列
 * @customfunction
 */
function getZeroPriceBooks() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const isbnSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISBN_LIST);
    
    if (!isbnSheet) {
      return [['', '']];
    }
    
    const lastRow = isbnSheet.getLastRow();
    if (lastRow < 2) {
      return [['', '']];
    }
    
    const data = isbnSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    // [ISBN, タイトル, 著者, 出版社, 最新見積価格]
    
    const zeroBooks = [];
    
    data.forEach(row => {
      const isbn = String(row[0]).trim();
      const title = row[1];
      const price = parseFloat(row[4]);
      
      if (isbn && title && price === 0) {
        zeroBooks.push([title, isbn]);
      }
    });
    
    // 1件もない場合
    if (zeroBooks.length === 0) {
      return [['なし', '']];
    }
    
    return zeroBooks;
    
  } catch (error) {
    logError(`getZeroPriceBooks エラー: ${error.message}`);
    return [['エラー', '']];
  }
}

// ========================================
// 高利益書籍ランキング用関数
// ========================================

/**
 * 高利益書籍TOP10を取得
 * @returns {Array} [タイトル, 見積価格, 売却価格, 利益]の配列
 * @customfunction
 */
function getTopProfitBooks() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    const bookData = [];
    
    // 全「買取完了_*」シートからデータを取得
    sheets.forEach(sheet => {
      const name = sheet.getName();
      
      if (name.startsWith('買取完了_')) {
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) {
          return; // データなし
        }
        
        // A列: ISBN, B列: タイトル, E列: 見積価格, F列: 売却価格, G列: 利益
        const data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
        
        data.forEach(row => {
          const title = row[1];
          const estimatePrice = parseFloat(row[4]);
          const salePrice = parseFloat(row[5]);
          const profit = parseFloat(row[6]);
          
          // タイトルと利益が有効な場合のみ追加
          if (title && !isNaN(profit) && profit > 0) {
            bookData.push({
              title: title,
              estimatePrice: isNaN(estimatePrice) ? 0 : estimatePrice,
              salePrice: isNaN(salePrice) ? 0 : salePrice,
              profit: profit
            });
          }
        });
      }
    });
    
    // データがない場合
    if (bookData.length === 0) {
      const emptyRow = ['', '', '', ''];
      return Array(10).fill(emptyRow);
    }
    
    // 利益の大きい順にソート
    bookData.sort((a, b) => b.profit - a.profit);
    
    // TOP 10を取得
    const top10 = bookData.slice(0, 10);
    
    // 表示用の配列に変換
    const result = top10.map(item => [
      item.title,
      item.estimatePrice,
      item.salePrice,
      item.profit
    ]);
    
    // 10件未満の場合は空行で埋める
    while (result.length < 10) {
      result.push(['', '', '', '']);
    }
    
    return result;
    
  } catch (error) {
    logError(`getTopProfitBooks エラー: ${error.message}`);
    const emptyRow = ['', '', '', ''];
    return Array(10).fill(emptyRow);
  }
}
