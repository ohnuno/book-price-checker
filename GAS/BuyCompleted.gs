/**
 * カスタムメニューを追加（onOpenトリガー）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 古本買取システム')
    .addItem('✅ 買取完了に移行', 'moveToBuyCompleted')
    .addSeparator()
    .addItem('📊 ダッシュボードをセットアップ', 'setupDashboardSheet')
    .addItem('🔄 ダッシュボードを更新', 'refreshDashboard')
    .addItem('🎁 キャンペーン情報を更新', 'updateCampaignInfo')
    .addToUi();
}

/**
 * ダッシュボードを強制更新
 */
function refreshDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
    
    if (!dashboardSheet) {
      SpreadsheetApp.getUi().alert(
        'エラー',
        'ダッシュボードシートが見つかりません',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    logInfo('ダッシュボードの強制更新を開始');
    
    // カスタム関数を含むセルのリスト
    const formulaCells = [
      'B6', 'B7', 'D7',           // 現在の状況
      'B8', 'B9', 'B10', 'B11',   // 価格統計
      'B14', 'B15', 'B16', 'B17', // 買取実績
      'B20', 'B21', 'B22'         // 今月の実績
    ];
    
    // 価格上昇TOP5 (F7:I11)
    for (let i = 7; i <= 11; i++) {
      formulaCells.push(`F${i}`, `G${i}`, `H${i}`, `I${i}`);
    }
    
    // 価格下落TOP5 (F15:I19)
    for (let i = 15; i <= 19; i++) {
      formulaCells.push(`F${i}`, `G${i}`, `H${i}`, `I${i}`);
    }
    
    // 0円書籍 (F22, G22)
    formulaCells.push('F22', 'G22');
    
    // 高利益TOP10 (F32:I41)
    for (let i = 32; i <= 41; i++) {
      formulaCells.push(`F${i}`, `G${i}`, `H${i}`, `I${i}`);
    }
    
    // Step 1: 各セルの数式を一時的に保存
    const formulas = {};
    formulaCells.forEach(cell => {
      const range = dashboardSheet.getRange(cell);
      const formula = range.getFormula();
      if (formula) {
        formulas[cell] = formula;
      }
    });
    
    logInfo(`${Object.keys(formulas).length}個の数式を保存しました`);
    
    // Step 2: 数式をクリア
    formulaCells.forEach(cell => {
      dashboardSheet.getRange(cell).clear();
    });
    
    // 強制的にスプレッドシートをフラッシュ（変更を確定）
    SpreadsheetApp.flush();
    
    // Step 3: 数式を再設定
    Object.keys(formulas).forEach(cell => {
      dashboardSheet.getRange(cell).setFormula(formulas[cell]);
    });
    
    // 再度フラッシュ
    SpreadsheetApp.flush();
    
    logInfo('すべての数式を再設定しました');
    
    // Step 4: 最終更新時刻を更新（日本時間）
    const now = new Date();
    const jstTime = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    dashboardSheet.getRange('B25').setValue(jstTime);
    
    // 完了通知
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'ダッシュボードを更新しました（全カスタム関数を再計算）',
      '✅ 更新完了',
      3
    );
    
    logInfo('ダッシュボードの強制更新が完了しました');
    
  } catch (error) {
    logError(`refreshDashboard エラー: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `更新中にエラーが発生しました: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * チェックされた書籍を買取完了シートに移行
 */
function moveToBuyCompleted() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const isbnSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISBN_LIST);
    
    if (!isbnSheet) {
      showAlert('エラー', '必要なシートが見つかりません');
      return;
    }
    
    // データを取得（ヘッダー除く）
    const dataRange = isbnSheet.getRange(2, 1, isbnSheet.getLastRow() - 1, CONFIG.ISBN_LIST_COLUMNS.CHECKBOX);
    const data = dataRange.getValues();
    
    // チェックされた書籍を確認
    let checkedCount = 0;
    for (let i = 0; i < data.length; i++) {
      const checkbox = data[i][CONFIG.ISBN_LIST_COLUMNS.CHECKBOX - 1];
      if (checkbox === true) {
        checkedCount++;
      }
    }
    
    if (checkedCount === 0) {
      showToast('チェックされた書籍がありません', 'ℹ️ 情報', 3);
      return;
    }
    
    // 記録先シートを選択
    const targetSheetName = selectTargetSheet(ss, checkedCount);
    
    if (!targetSheetName) {
      // ユーザーがキャンセルした
      return;
    }
    
    // ターゲットシートを取得または作成
    let targetSheet = ss.getSheetByName(targetSheetName);
    
    if (!targetSheet) {
      // 新規シート作成
      targetSheet = createNewBuySheet(ss, targetSheetName);
      logInfo(`新しいシートを作成: ${targetSheetName}`);
    }
    
    let movedCount = 0;
    const rowsToDelete = [];
    
    // 下から上に処理（行削除の影響を受けないように）
    for (let i = data.length - 1; i >= 0; i--) {
      const row = i + 2; // ヘッダー行を考慮
      const isbn = data[i][CONFIG.ISBN_LIST_COLUMNS.ISBN - 1];
      const checkbox = data[i][CONFIG.ISBN_LIST_COLUMNS.CHECKBOX - 1];
      
      // チェックボックスがtrueの行を処理
      if (checkbox === true && isbn) {
        const title = data[i][CONFIG.ISBN_LIST_COLUMNS.TITLE - 1];
        const author = data[i][CONFIG.ISBN_LIST_COLUMNS.AUTHOR - 1];
        const publisher = data[i][CONFIG.ISBN_LIST_COLUMNS.PUBLISHER - 1];
        const estimatePrice = data[i][CONFIG.ISBN_LIST_COLUMNS.PRICE - 1];
        
        // 選択したシートに追加
        const newRow = [
          isbn,
          title,
          author,
          publisher,
          estimatePrice || 0,  // E列: 見積価格
          '',                  // F列: 実際の買取価格（空白）
          '',                  // G列: 差額（空白）
          formatDateTime(new Date())  // H列: 登録日
        ];
        
        targetSheet.appendRow(newRow);
        
        // 差額の計算式を設定（G列 = F列 - E列）
        const lastRow = targetSheet.getLastRow();
        const diffCell = targetSheet.getRange(lastRow, 7);  // G列
        diffCell.setFormula(`=F${lastRow}-E${lastRow}`);
        
        // 価格履歴から該当ISBNの全履歴を削除
        deletePriceHistory(ss, isbn);
        
        // ISBNリストから削除対象としてマーク
        rowsToDelete.push(row);
        movedCount++;
        
        logInfo(`買取完了に移行: ${isbn} - ${title} → ${targetSheetName}`);
      }
    }
    
    // 行を削除（下から順に削除）
    for (let row of rowsToDelete) {
      isbnSheet.deleteRow(row);
    }
    
    // データ追加後に列幅を自動調整
    if (movedCount > 0) {
      adjustColumnWidths(targetSheet);
    }
    
    // 結果を表示
    if (movedCount > 0) {
      showToast(
        `${movedCount}件の書籍を「${targetSheetName}」に移行しました`,
        '✅ 移行完了',
        5
      );
    }
    
  } catch (error) {
    logError(`買取完了移行エラー: ${error.message}`);
    showAlert('エラー', `処理中にエラーが発生しました: ${error.message}`);
  }
}

/**
 * シートの列幅を自動調整
 * @param {Sheet} sheet - 調整対象のシート
 */
function adjustColumnWidths(sheet) {
  try {
    const lastColumn = sheet.getLastColumn();
    
    // 各列を自動調整
    for (let col = 1; col <= lastColumn; col++) {
      sheet.autoResizeColumn(col);
    }
    
    // 調整後、最小幅と最大幅を設定
    const columnSettings = {
      1: { min: 130, max: 150 },  // ISBN
      2: { min: 200, max: 400 },  // タイトル
      3: { min: 100, max: 200 },  // 著者
      4: { min: 100, max: 200 },  // 出版社
      5: { min: 100, max: 150 },  // 最新見積価格
      6: { min: 100, max: 150 },  // 売却価格
      7: { min: 80, max: 120 },   // 利益
      8: { min: 150, max: 200 }   // 登録日
    };
    
    for (let col = 1; col <= lastColumn; col++) {
      const currentWidth = sheet.getColumnWidth(col);
      const settings = columnSettings[col];
      
      if (settings) {
        if (currentWidth < settings.min) {
          sheet.setColumnWidth(col, settings.min);
        } else if (currentWidth > settings.max) {
          sheet.setColumnWidth(col, settings.max);
        }
      }
    }
    
    logInfo(`列幅自動調整完了: ${sheet.getName()}`);
    
  } catch (error) {
    logError(`列幅調整エラー: ${error.message}`);
    // エラーが発生しても処理は継続
  }
}

/**
 * 記録先シートを選択するダイアログを表示
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {number} bookCount - 処理する書籍の件数
 * @returns {string|null} 選択されたシート名（キャンセル時はnull）
 */
function selectTargetSheet(ss, bookCount) {
  const ui = SpreadsheetApp.getUi();
  
  // 直近1ヶ月以内の買取完了シートを取得
  const recentSheets = getRecentBuySheets(ss);
  
  // 当日の日付でデフォルトのシート名を生成
  const todaySheetName = getTodaySheetName();
  
  // ダイアログメッセージを作成
  let message = `${bookCount}件の書籍を処理します。\n記録先シートを選択してください。\n\n`;
  
  if (recentSheets.length > 0) {
    message += '【既存のシート】\n';
    recentSheets.forEach((sheetName, index) => {
      message += `${index + 1}. ${sheetName}\n`;
    });
    message += `\n${recentSheets.length + 1}. 新しいシート（${todaySheetName}）\n\n`;
    message += `番号を入力してください（1-${recentSheets.length + 1}）:`;
  } else {
    message += '既存の記録シートがありません。\n';
    message += `新しいシート「${todaySheetName}」を作成します。\n\n`;
    message += 'よろしいですか？（OK = 作成 / Cancel = キャンセル）';
  }
  
  const response = ui.prompt(
    '📋 記録先シートの選択',
    message,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;  // キャンセル
  }
  
  const userInput = response.getResponseText().trim();
  
  if (recentSheets.length === 0) {
    // 既存シートがない場合は新規作成
    return todaySheetName;
  }
  
  // 入力を数値に変換
  const selection = parseInt(userInput);
  
  if (isNaN(selection) || selection < 1 || selection > recentSheets.length + 1) {
    showAlert('エラー', `無効な選択です: ${userInput}\n1〜${recentSheets.length + 1}の番号を入力してください。`);
    return null;
  }
  
  if (selection === recentSheets.length + 1) {
    // 新しいシートを作成
    return todaySheetName;
  } else {
    // 既存のシートを選択
    return recentSheets[selection - 1];
  }
}

/**
 * 直近1ヶ月以内の買取完了シートを取得
 * @param {Spreadsheet} ss - スプレッドシート
 * @returns {Array<string>} シート名の配列（新しい順）
 */
function getRecentBuySheets(ss) {
  const sheets = ss.getSheets();
  const oneMonthAgo = new Date();
  oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
  
  const recentSheets = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    
    // "買取完了_YYYY-MM-DD" 形式のシート名のみ対象
    if (name.startsWith('買取完了_')) {
      const dateStr = name.replace('買取完了_', '');
      
      // 日付文字列をパース（YYYY-MM-DD形式）
      try {
        const sheetDate = new Date(dateStr);
        
        // 有効な日付で、1ヶ月以内のもののみ追加
        if (!isNaN(sheetDate.getTime()) && sheetDate >= oneMonthAgo) {
          recentSheets.push({
            name: name,
            date: sheetDate
          });
        }
      } catch (e) {
        // 日付パースエラーは無視
        logWarning(`日付パースエラー: ${name}`);
      }
    }
  });
  
  // 日付の新しい順にソート
  recentSheets.sort((a, b) => b.date - a.date);
  
  return recentSheets.map(item => item.name);
}

/**
 * 当日の日付でシート名を生成
 * @returns {string} シート名（例: "買取完了_2025-12-05"）
 */
function getTodaySheetName() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  
  return `買取完了_${year}-${month}-${day}`;
}

/**
 * 新しい買取完了シートを作成
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} sheetName - シート名
 * @returns {Sheet} 作成されたシート
 */
function createNewBuySheet(ss, sheetName) {
  // エラーログシートの位置を取得
  const errorLogSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ERROR_LOG);
  
  let newSheet;
  
  if (errorLogSheet) {
    // エラーログの右に挿入（エラーログのインデックス位置に挿入）
    const errorLogIndex = errorLogSheet.getIndex();
    newSheet = ss.insertSheet(sheetName, errorLogIndex);
    logInfo(`新規シート作成: ${sheetName}（位置: ${errorLogIndex}）`);
  } else {
    // エラーログシートがない場合は末尾に追加
    newSheet = ss.insertSheet(sheetName);
    logWarning('エラーログシートが見つかりません。シートを末尾に作成しました。');
  }
  
  // ヘッダー行を設定
  const headers = [
    'ISBN',
    'タイトル',
    '著者',
    '出版社',
    '最新見積価格',
    '売却価格',
    '利益',
    '登録日'
  ];
  
  const headerRange = newSheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  
  // ヘッダー行の書式設定
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setHorizontalAlignment('center');
  
  // 初期列幅を設定（ヘッダーに合わせて最小限の幅）
  // データ追加後に自動調整されるため、ここでは基本的な幅のみ設定
  const defaultWidths = [130, 300, 150, 150, 120, 120, 100, 180];
  for (let i = 0; i < defaultWidths.length; i++) {
    newSheet.setColumnWidth(i + 1, defaultWidths[i]);
  }
  
  logInfo(`ヘッダー設定完了: ${sheetName}`);
  
  return newSheet;
}

/**
 * 価格履歴から該当ISBNの全行を削除
 * @param {Spreadsheet} ss - スプレッドシート
 * @param {string} isbn - ISBN
 */
function deletePriceHistory(ss, isbn) {
  try {
    logInfo(`価格履歴削除開始: ISBN ${isbn}`);
    
    const historySheet = ss.getSheetByName(CONFIG.SHEET_NAMES.PRICE_HISTORY);
    
    if (!historySheet) {
      logWarning('価格履歴シートが見つかりません');
      return;
    }
    
    // 全データを取得（ヘッダー除く）
    const lastRow = historySheet.getLastRow();
    if (lastRow < 2) {
      logInfo('価格履歴シートにデータがありません');
      return;
    }
    
    const data = historySheet.getRange(2, 1, lastRow - 1, 1).getValues(); // A列（ISBN）のみ取得
    
    let deletedCount = 0;
    
    // 後ろから削除（行番号がずれないように）
    for (let i = data.length - 1; i >= 0; i--) {
      const rowIsbn = String(data[i][0]).trim();
      const rowNumber = i + 2; // ヘッダー行を考慮
      
      if (rowIsbn === String(isbn).trim()) {
        historySheet.deleteRow(rowNumber);
        deletedCount++;
        logInfo(`  行${rowNumber}を削除: ISBN ${rowIsbn}`);
      }
    }
    
    logInfo(`価格履歴削除完了: ${deletedCount}件削除 (ISBN: ${isbn})`);
    
  } catch (error) {
    logError(`価格履歴削除エラー: ${error.message}`);
    // エラーが発生しても処理は継続（買取完了への移行は実行される）
  }
}

/**
 * アラートを表示
 * @param {string} title - タイトル
 * @param {string} message - メッセージ
 */
function showAlert(title, message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(title, message, ui.ButtonSet.OK);
}

/**
 * トースト通知を表示
 * @param {string} message - メッセージ
 * @param {string} title - タイトル
 * @param {number} timeout - 表示時間（秒）
 */
function showToast(message, title, timeout) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeout);
}
