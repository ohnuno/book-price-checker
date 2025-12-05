/**
 * カスタムメニューを追加（onOpenトリガー）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📚 古本買取システム')
    .addItem('✅ 買取完了に移行', 'moveToBuyCompleted')
    .addToUi();
}

/**
 * チェックされた書籍を買取完了シートに移行
 */
function moveToBuyCompleted() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const isbnSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.ISBN_LIST);
    const completedSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.COMPLETED);
    
    if (!isbnSheet || !completedSheet) {
      showAlert('エラー', '必要なシートが見つかりません');
      return;
    }
    
    // データを取得（ヘッダー除く）
    const dataRange = isbnSheet.getRange(2, 1, isbnSheet.getLastRow() - 1, CONFIG.ISBN_LIST_COLUMNS.CHECKBOX);
    const data = dataRange.getValues();
    
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
        
        // 買取完了シートに追加
        const newRow = [
          isbn,
          title,
          author,
          publisher,
          estimatePrice || 0,  // E列: 見積価格
          '',                  // F列: 実際の買取価格（手動入力）
          '',                  // G列: 差額（数式で自動計算）
          formatDateTime(new Date())  // H列: 買取完了日時
        ];
        
        completedSheet.appendRow(newRow);
        
        // 差額の計算式を設定（G列 = F列 - E列）
        const lastRow = completedSheet.getLastRow();
        const diffCell = completedSheet.getRange(lastRow, CONFIG.COMPLETED_COLUMNS.DIFFERENCE);
        diffCell.setFormula(`=F${lastRow}-E${lastRow}`);
        
        // ISBNリストから削除対象としてマーク
        rowsToDelete.push(row);
        movedCount++;
        
        logInfo(`買取完了に移行: ${isbn} - ${title}`);
      }
    }
    
    // 行を削除（下から順に削除）
    for (let row of rowsToDelete) {
      isbnSheet.deleteRow(row);
    }
    
    // 結果を表示
    if (movedCount > 0) {
      showToast(
        `${movedCount}件の書籍を買取完了に移行しました`,
        '✅ 移行完了',
        5
      );
      
      // 実際の買取価格入力を促すメッセージ
      showToast(
        '買取完了シートのF列に実際の買取価格を入力してください',
        'ℹ️ 次のステップ',
        10
      );
    } else {
      showToast(
        'チェックされた書籍がありません',
        'ℹ️ 情報',
        3
      );
    }
    
  } catch (error) {
    logError(`買取完了移行エラー: ${error.message}`);
    showAlert('エラー', `処理中にエラーが発生しました: ${error.message}`);
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