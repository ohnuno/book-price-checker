/**
 * Google Books APIから書籍情報を取得（日本対応版）
 */
function fetchBookInfoFromGoogleBooks(isbn) {
  try {
    const apiKey = getScriptProperty('GOOGLE_BOOKS_API_KEY');
    let url = `https://www.googleapis.com/books/v1/volumes?q=isbn:${isbn}&country=JP`;
    
    if (apiKey) {
      url += `&key=${apiKey}`;
    }
    
    logInfo(`Google Books APIから書籍情報を取得: ${isbn}`);
    
    const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    const data = JSON.parse(response.getContentText());
    
    if (!data.items || data.items.length === 0) {
      logWarning(`Google Books APIで書籍が見つかりませんでした: ${isbn}`);
      return null;
    }
    
    const volumeInfo = data.items[0].volumeInfo;
    
    const bookInfo = {
      isbn: isbn,
      title: volumeInfo.title || '（タイトル不明）',
      author: volumeInfo.authors ? volumeInfo.authors.join(', ') : '',
      publisher: volumeInfo.publisher || ''
    };
    
    logInfo(`Google Books API取得成功: ${bookInfo.title}`);
    return bookInfo;
    
  } catch (error) {
    logError(`Google Books API取得エラー: ${error.message}`);
    return null;
  }
}

/**
 * スプレッドシート起動時にカスタムメニューを追加
 * 
 * トリガー設定:
 * 1. Apps Scriptエディタで「トリガー」をクリック
 * 2. 「トリガーを追加」をクリック
 * 3. 実行する関数: createBookInfoMenu
 * 4. イベントのソース: スプレッドシートから
 * 5. イベントの種類: 起動時
 * 6. 保存
 */
function createBookInfoMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 書籍情報')
    .addItem('📝 全ISBNの情報を取得', 'processAllISBNs')
    .addItem('🔄 選択範囲のISBNを処理', 'processSelectedISBNs')
    .addSeparator()
    .addItem('✅ 全行にチェックボックス設定', 'addCheckboxesToAll')
    .addToUi();
}

/**
 * 全ISBNに対して書籍情報を取得
 */
function processAllISBNs() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '全ISBN処理の確認',
    'ISBNリストシートの全ISBNに対して書籍情報を取得します。\n既に情報がある行はスキップされます。\n\n処理を開始しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISBN_LIST);
  
  if (!sheet) {
    ui.alert('❌ エラー', 'ISBNリストシートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  
  // データ範囲を取得（ヘッダー行を除く）
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ 情報', '処理するISBNがありません。', ui.ButtonSet.OK);
    return;
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues();
  
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  
  // 進捗表示用のトースト
  ss.toast('処理を開始します...', '📚 書籍情報取得', 3);
  
  for (let i = 0; i < values.length; i++) {
    const row = i + 2; // 実際の行番号（ヘッダー除く）
    const isbn = values[i][0]; // A列（ISBN）
    const existingTitle = values[i][1]; // B列（書籍名）
    
    // ISBNを文字列に変換
    const isbnStr = isbn ? String(isbn).trim() : '';
    
    // 空白または無効なISBNはスキップ
    if (!isbnStr || !isValidISBN(isbnStr)) {
      continue;
    }
    
    // 既に書籍情報がある場合はスキップ（チェックボックスのみ確認）
    if (existingTitle) {
      ensureCheckbox(sheet, row);
      skippedCount++;
      continue;
    }
    
    // 進捗表示（10件ごと）
    if (processedCount % 10 === 0) {
      ss.toast(
        `処理中: ${processedCount + 1}/${values.length}`,
        '📚 書籍情報取得',
        1
      );
    }
    
    // Google Books APIから書籍情報を取得
    try {
      const bookInfo = fetchBookInfoFromGoogleBooks(isbnStr);
      
      if (bookInfo) {
        // 書籍情報を設定
        sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.TITLE).setValue(bookInfo.title);
        sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.AUTHOR).setValue(bookInfo.author);
        sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.PUBLISHER).setValue(bookInfo.publisher);
        sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.STATUS).setValue(CONFIG.STATUS.NOT_SOLD);
        
        // チェックボックスを設定
        ensureCheckbox(sheet, row);
        
        processedCount++;
      } else {
        // API失敗時もチェックボックスは設定
        ensureCheckbox(sheet, row);
        errorCount++;
      }
      
      // APIレート制限を考慮して待機（100ms）
      Utilities.sleep(100);
      
    } catch (error) {
      logError(`ISBN処理エラー (行${row}): ${error.message}`);
      ensureCheckbox(sheet, row);
      errorCount++;
    }
  }
  
  // 完了メッセージ
  const message = `処理完了\n\n✅ 取得成功: ${processedCount}件\n⏭️ スキップ: ${skippedCount}件\n❌ 失敗: ${errorCount}件`;
  ss.toast(message, '📚 完了', 10);
  ui.alert('✅ 処理完了', message, ui.ButtonSet.OK);
}

/**
 * 選択範囲のISBNに対して書籍情報を取得
 */
function processSelectedISBNs() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  // ISBNリストシート以外は処理しない
  if (sheet.getName() !== CONFIG.SHEET_NAMES.ISBN_LIST) {
    ui.alert('❌ エラー', 'ISBNリストシートで実行してください。', ui.ButtonSet.OK);
    return;
  }
  
  const selection = sheet.getActiveRange();
  if (!selection) {
    ui.alert('❌ エラー', 'セル範囲を選択してください。', ui.ButtonSet.OK);
    return;
  }
  
  // A列のみを処理
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  if (startRow === 1) {
    ui.alert('❌ エラー', 'ヘッダー行以外を選択してください。', ui.ButtonSet.OK);
    return;
  }
  
  let processedCount = 0;
  let skippedCount = 0;
  let errorCount = 0;
  
  ss.toast('処理を開始します...', '📚 書籍情報取得', 3);
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const isbn = sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.ISBN).getValue();
    const existingTitle = sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.TITLE).getValue();
    
    // ISBNを文字列に変換
    const isbnStr = isbn ? String(isbn).trim() : '';
    
    // 空白または無効なISBNはスキップ
    if (!isbnStr || !isValidISBN(isbnStr)) {
      continue;
    }
    
    // 既に書籍情報がある場合はスキップ
    if (existingTitle) {
      ensureCheckbox(sheet, row);
      skippedCount++;
      continue;
    }
    
    // Google Books APIから書籍情報を取得
    try {
      const bookInfo = fetchBookInfoFromGoogleBooks(isbnStr);
      
      if (bookInfo) {
        // 書籍情報を設定
        sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.TITLE).setValue(bookInfo.title);
        sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.AUTHOR).setValue(bookInfo.author);
        sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.PUBLISHER).setValue(bookInfo.publisher);
        sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.STATUS).setValue(CONFIG.STATUS.NOT_SOLD);
        
        // チェックボックスを設定
        ensureCheckbox(sheet, row);
        
        processedCount++;
      } else {
        ensureCheckbox(sheet, row);
        errorCount++;
      }
      
      // APIレート制限を考慮して待機（100ms）
      Utilities.sleep(100);
      
    } catch (error) {
      logError(`ISBN処理エラー (行${row}): ${error.message}`);
      ensureCheckbox(sheet, row);
      errorCount++;
    }
  }
  
  // 完了メッセージ
  const message = `処理完了\n\n✅ 取得成功: ${processedCount}件\n⏭️ スキップ: ${skippedCount}件\n❌ 失敗: ${errorCount}件`;
  ss.toast(message, '📚 完了', 10);
  ui.alert('✅ 処理完了', message, ui.ButtonSet.OK);
}

/**
 * 全行にチェックボックスを設定
 */
function addCheckboxesToAll() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISBN_LIST);
  
  if (!sheet) {
    ui.alert('❌ エラー', 'ISBNリストシートが見つかりません。', ui.ButtonSet.OK);
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('ℹ️ 情報', '処理する行がありません。', ui.ButtonSet.OK);
    return;
  }
  
  ss.toast('チェックボックスを設定中...', '✅ 処理中', 3);
  
  let addedCount = 0;
  
  for (let row = 2; row <= lastRow; row++) {
    const isbn = sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.ISBN).getValue();
    const isbnStr = isbn ? String(isbn).trim() : '';
    
    // ISBNがある行のみ処理
    if (isbnStr && isValidISBN(isbnStr)) {
      ensureCheckbox(sheet, row);
      addedCount++;
    }
  }
  
  ss.toast(`${addedCount}行にチェックボックスを設定しました`, '✅ 完了', 5);
  ui.alert('✅ 完了', `${addedCount}行にチェックボックスを設定しました。`, ui.ButtonSet.OK);
}

/**
 * ISBN入力時の自動処理（onEditトリガー）
 */
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    
    // ISBNリストシート以外は無視
    if (sheet.getName() !== CONFIG.SHEET_NAMES.ISBN_LIST) {
      return;
    }
    
    // A列（ISBN列）以外は無視
    if (range.getColumn() !== CONFIG.ISBN_LIST_COLUMNS.ISBN) {
      return;
    }
    
    // ヘッダー行は無視
    if (range.getRow() === 1) {
      return;
    }
    
    const isbn = range.getValue();
    const row = range.getRow();
    
    // ISBNを文字列に変換
    const isbnStr = isbn ? String(isbn).trim() : '';
    
    // 空白または無効なISBNは無視
    if (!isbnStr || !isValidISBN(isbnStr)) {
      // ISBNが削除された場合、チェックボックスも削除
      if (!isbnStr) {
        const checkboxCell = sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.CHECKBOX);
        checkboxCell.clearContent();
        checkboxCell.clearDataValidations();
      }
      return;
    }
    
    // 既に書籍名が入力されている場合はスキップ
    const existingTitle = sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.TITLE).getValue();
    
    if (existingTitle) {
      logInfo(`既に書籍情報が登録されています: 行${row}`);
      // チェックボックスが未設定なら設定
      ensureCheckbox(sheet, row);
      return;
    }
    
    // Google Books APIから書籍情報を取得
    logInfo(`ISBN入力検知: ${isbnStr} (行${row})`);
    
    const bookInfo = fetchBookInfoFromGoogleBooks(isbnStr);
    
    if (bookInfo) {
      // 書籍情報を設定
      sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.TITLE).setValue(bookInfo.title);
      sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.AUTHOR).setValue(bookInfo.author);
      sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.PUBLISHER).setValue(bookInfo.publisher);
      sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.STATUS).setValue(CONFIG.STATUS.NOT_SOLD);
      
      // チェックボックスを設定
      ensureCheckbox(sheet, row);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `書籍情報を取得しました: ${bookInfo.title}`,
        '✅ 取得成功',
        5
      );
    } else {
      // API失敗時もチェックボックスは設定
      ensureCheckbox(sheet, row);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `書籍情報の取得に失敗しました: ${isbnStr}`,
        '❌ 取得失敗',
        5
      );
    }
    
  } catch (error) {
    logError(`onEditエラー: ${error.message}`);
    
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `エラーが発生しました: ${error.message}`,
        '❌ エラー',
        10
      );
    } catch (toastError) {
      Logger.log(`トースト表示もエラー: ${toastError.message}`);
    }
  }
}

/**
 * チェックボックスが未設定なら設定する
 */
function ensureCheckbox(sheet, row) {
  try {
    const checkboxCell = sheet.getRange(row, CONFIG.ISBN_LIST_COLUMNS.CHECKBOX);
    
    // 既にチェックボックスが設定されているかチェック
    const validations = checkboxCell.getDataValidations();
    if (validations && validations[0] && validations[0][0]) {
      return;
    }
    
    // チェックボックスを設定
    checkboxCell.insertCheckboxes();
    checkboxCell.setValue(false);
    
    logInfo(`チェックボックスを設定しました: 行${row}`);
    
  } catch (error) {
    logError(`チェックボックス設定エラー: ${error.message}`);
  }
}