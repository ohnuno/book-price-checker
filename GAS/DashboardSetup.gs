/**
 * ダッシュボードシート自動セットアップスクリプト
 * 
 * 使い方:
 * 1. このスクリプトをApps Scriptエディタに追加
 * 2. setupDashboardSheet() 関数を実行
 * 
 * ファイル名: DashboardSetup.gs
 */

/**
 * ダッシュボードシートを自動セットアップ
 */
function setupDashboardSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ダッシュボードシートを取得または作成
    let dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
    
    if (!dashboardSheet) {
      logInfo('ダッシュボードシートを新規作成します');
      dashboardSheet = ss.insertSheet(CONFIG.SHEET_NAMES.DASHBOARD, 0); // 先頭に配置
    } else {
      logInfo('既存のダッシュボードシートをクリアします');
      dashboardSheet.clear();
    }
    
    // シートをセットアップ
    setupLayout(dashboardSheet);
    setupFormulas(dashboardSheet);
    setupFormatting(dashboardSheet);
    
    // 完了メッセージ
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'ダッシュボードシートのセットアップが完了しました',
      '✅ セットアップ完了',
      5
    );
    
    logInfo('ダッシュボードシートのセットアップが完了しました');
    
  } catch (error) {
    logError(`setupDashboardSheet エラー: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `セットアップ中にエラーが発生しました: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * レイアウトを設定
 * @param {Sheet} sheet - ダッシュボードシート
 */
function setupLayout(sheet) {
  logInfo('レイアウトを設定中...');
  
  // データを配列で定義（A-D列: 基本統計、E列: 空列、F-I列: 価格変動アラート）
  const data = [
    // Row 1-2: ヘッダー（左右に分割）
    ['📊 古本買取システム ダッシュボード', '', '', '', '', '📈 価格変動アラート', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    
    // Row 3: 空行
    ['', '', '', '', '', '', '', '', ''],
    
    // Row 4: 空行
    ['', '', '', '', '', '', '', '', ''],
    
    // Row 5-11: 現在の状況 | 価格上昇TOP5ヘッダー
    ['【現在の状況】', '', '', '', '', '【価格上昇 TOP 5】', '', '', ''],
    ['登録中の書籍数', '', '冊', '', '', 'タイトル', '初回価格', '最新価格', '変動額'],
    ['本日更新済み', '', '冊', '', '', '', '', '', ''],
    ['平均見積価格', '', '', '', '', '', '', '', ''],
    ['最高見積価格', '', '', '', '', '', '', '', ''],
    ['最低見積価格', '', '', '', '', '', '', '', ''],
    ['見積額総額', '', '', '', '', '', '', '', ''],
    
    // Row 12: 空行 | データ行
    ['', '', '', '', '', '', '', '', ''],
    
    // Row 13-17: 買取実績 | 価格下落TOP5ヘッダー
    ['【買取実績】', '', '', '', '', '【価格下落 TOP 5】', '', '', ''],
    ['総買取冊数', '', '冊', '', '', 'タイトル', '初回価格', '最新価格', '変動額'],
    ['総利益', '', '', '', '', '', '', '', ''],
    ['平均利益', '', '/冊', '', '', '', '', '', ''],
    ['最高利益', '', '/冊', '', '', '', '', '', ''],
    
    // Row 18: 空行 | データ行
    ['', '', '', '', '', '', '', '', ''],
    
    // Row 19-22: 今月の実績 | データ行
    ['【今月の実績】', '', '', '', '', '', '', '', ''],
    ['買取冊数', '', '冊', '', '', '', '', '', ''],
    ['今月利益', '', '', '', '', '【0円になった書籍】', '', '', ''],
    ['平均利益', '', '/冊', '', '', 'タイトル', 'ISBN', '', ''],
    
    // Row 23-25: 区切り線、最終更新 | データ行
    ['', '', '', '', '', '', '', '', ''],
    ['───────────────────────────────────────', '', '', '', '', '', '', '', ''],
    ['最終更新', '', '', '', '', '', '', '', ''],
    
    // Row 26-27: 空行
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    
    // Row 28-40: キャンペーン情報エリア（A~D列）| 高利益書籍ランキング（F~I列）
    ['', '', '', '', '', '💰 高利益書籍ランキング', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '【高利益 TOP 10】', '', '', ''],
    ['', '', '', '', '', 'タイトル', '見積価格', '売却価格', '利益'],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '']
  ];
  
  // データを一括で書き込み
  sheet.getRange(1, 1, data.length, 9).setValues(data);
  
  logInfo('レイアウト設定完了');
}

/**
 * 数式を設定
 * @param {Sheet} sheet - ダッシュボードシート
 */
function setupFormulas(sheet) {
  logInfo('数式を設定中...');
  
  // === 基本統計サマリ ===
  
  // B6: 登録中の書籍数
  sheet.getRange('B6').setFormula('=COUNTA(ISBNリスト!A:A)-1');
  
  // B7: 本日更新済み
  sheet.getRange('B7').setFormula('=COUNTIF(ISBNリスト!F:F,TEXT(TODAY(),"yyyy/mm/dd")&"*")');
  
  // D7: 更新率
  sheet.getRange('D7').setFormula('=IF(B6>0,B7/B6,"0%")');
  
  // B8: 平均見積価格
  sheet.getRange('B8').setFormula('=IFERROR(ROUND(AVERAGE(ISBNリスト!E:E),0),0)');
  
  // B9: 最高見積価格
  sheet.getRange('B9').setFormula('=IFERROR(MAX(ISBNリスト!E:E),0)');
  
  // B10: 最低見積価格（0円除く）
  sheet.getRange('B10').setFormula('=IFERROR(MINIFS(ISBNリスト!E:E,ISBNリスト!E:E,">0"),0)');
  
  // B11: 見積額総額
  sheet.getRange('B11').setFormula('=IFERROR(SUM(ISBNリスト!E:E),0)');
  
  // B14: 総買取冊数
  sheet.getRange('B14').setFormula('=getTotalBuyCount()');
  
  // B15: 総利益
  sheet.getRange('B15').setFormula('=getTotalProfit()');
  
  // B16: 平均利益
  sheet.getRange('B16').setFormula('=IF(B14>0,ROUND(B15/B14,0),0)');
  
  // B17: 最高利益
  sheet.getRange('B17').setFormula('=getMaxProfit()');
  
  // B20: 今月買取冊数
  sheet.getRange('B20').setFormula('=getMonthlyBuyCount()');
  
  // B21: 今月利益
  sheet.getRange('B21').setFormula('=getMonthlyProfit()');
  
  // B22: 今月平均利益
  sheet.getRange('B22').setFormula('=IF(B20>0,ROUND(B21/B20,0),0)');
  
  // B25: 最終更新（日本時間）
  const now = new Date();
  const jstTime = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  sheet.getRange('B25').setValue(jstTime);
  
  // === 価格変動アラート ===
  
  // 価格上昇TOP5: F7-I11
  for (let i = 0; i < 5; i++) {
    const row = 7 + i;
    sheet.getRange(`F${row}`).setFormula(`=INDEX(getPriceIncreasesTop5(),${i+1},1)`);
    sheet.getRange(`G${row}`).setFormula(`=INDEX(getPriceIncreasesTop5(),${i+1},2)`);
    sheet.getRange(`H${row}`).setFormula(`=INDEX(getPriceIncreasesTop5(),${i+1},3)`);
    sheet.getRange(`I${row}`).setFormula(`=INDEX(getPriceIncreasesTop5(),${i+1},4)`);
  }
  
  // 価格下落TOP5: F15-I19
  for (let i = 0; i < 5; i++) {
    const row = 15 + i;
    sheet.getRange(`F${row}`).setFormula(`=INDEX(getPriceDecreasesTop5(),${i+1},1)`);
    sheet.getRange(`G${row}`).setFormula(`=INDEX(getPriceDecreasesTop5(),${i+1},2)`);
    sheet.getRange(`H${row}`).setFormula(`=INDEX(getPriceDecreasesTop5(),${i+1},3)`);
    sheet.getRange(`I${row}`).setFormula(`=INDEX(getPriceDecreasesTop5(),${i+1},4)`);
  }
  
  // 0円書籍: F22, G22
  sheet.getRange('F22').setFormula('=INDEX(getZeroPriceBooks(),1,1)');
  sheet.getRange('G22').setFormula('=INDEX(getZeroPriceBooks(),1,2)');
  
  // === 高利益書籍ランキング ===
  
  // 高利益TOP10: F32-I41
  for (let i = 0; i < 10; i++) {
    const row = 32 + i;
    sheet.getRange(`F${row}`).setFormula(`=INDEX(getTopProfitBooks(),${i+1},1)`);
    sheet.getRange(`G${row}`).setFormula(`=INDEX(getTopProfitBooks(),${i+1},2)`);
    sheet.getRange(`H${row}`).setFormula(`=INDEX(getTopProfitBooks(),${i+1},3)`);
    sheet.getRange(`I${row}`).setFormula(`=INDEX(getTopProfitBooks(),${i+1},4)`);
  }
  
  logInfo('数式設定完了');
}

/**
 * 書式を設定
 * @param {Sheet} sheet - ダッシュボードシート
 */
function setupFormatting(sheet) {
  logInfo('書式を設定中...');
  
  // === 左側ヘッダー（A1:D2）===
  const leftHeaderRange = sheet.getRange('A1:D2');
  leftHeaderRange.merge();
  leftHeaderRange.setBackground('#4a86e8');
  leftHeaderRange.setFontColor('#ffffff');
  leftHeaderRange.setFontWeight('bold');
  leftHeaderRange.setFontSize(14);
  leftHeaderRange.setHorizontalAlignment('center');
  leftHeaderRange.setVerticalAlignment('middle');
  
  // === 右側ヘッダー（F1:I2）===
  const rightHeaderRange = sheet.getRange('F1:I2');
  rightHeaderRange.merge();
  rightHeaderRange.setBackground('#e8f4e8');
  rightHeaderRange.setFontColor('#000000');
  rightHeaderRange.setFontWeight('bold');
  rightHeaderRange.setFontSize(14);
  rightHeaderRange.setHorizontalAlignment('center');
  rightHeaderRange.setVerticalAlignment('middle');
  
  // === セクションヘッダー（左側：A5, A13, A19）===
  const leftSectionHeaders = ['A5', 'A13', 'A19'];
  leftSectionHeaders.forEach(cell => {
    const range = sheet.getRange(cell);
    range.setFontWeight('bold');
    range.setFontSize(11);
    range.setBackground('#f3f3f3');
  });
  
  // === セクションヘッダー（右側：E5, E13, E21, E30）===
  const rightSectionHeaders = ['F5', 'F13', 'F21', 'F30'];
  rightSectionHeaders.forEach(cell => {
    const range = sheet.getRange(cell);
    range.setFontWeight('bold');
    range.setFontSize(11);
    range.setBackground('#f3f3f3');
  });
  
  // === 高利益書籍ランキングヘッダー（F28:I29）===
  const profitHeaderRange = sheet.getRange('F28:I29');
  profitHeaderRange.merge();
  profitHeaderRange.setBackground('#fff3cd');
  profitHeaderRange.setFontWeight('bold');
  profitHeaderRange.setFontSize(12);
  profitHeaderRange.setHorizontalAlignment('center');
  profitHeaderRange.setVerticalAlignment('middle');
  
  // === テーブルヘッダー（E6:H6, E14:H14, E22:F22, E31:H31）===
  const tableHeaders = ['F6:I6', 'F14:I14', 'F22:G22', 'F31:I31'];
  tableHeaders.forEach(range => {
    sheet.getRange(range).setFontWeight('bold').setBackground('#f3f3f3');
  });
  
  // === ラベル列（A列）===
  sheet.getRange('A:A').setHorizontalAlignment('right');
  
  // === 数値列（B列）===
  sheet.getRange('B:B').setHorizontalAlignment('right');
  
  // === D7: 更新率をパーセンテージ表示 ===
  sheet.getRange('D7').setNumberFormat('0.0%');
  
  // === 価格に通貨記号を追加（表示形式）===
  const priceRanges = ['B8', 'B9', 'B10', 'B11', 'B15', 'B16', 'B17', 'B21', 'B22'];
  priceRanges.forEach(cell => {
    sheet.getRange(cell).setNumberFormat('"¥"#,##0');
  });
  
  // === 価格変動テーブルの数値書式 ===
  // 初回価格・最新価格（上昇）
  sheet.getRange('G7:H11').setNumberFormat('"¥"#,##0');
  // 変動額（上昇）+/-記号付き
  sheet.getRange('I7:I11').setNumberFormat('"+¥"#,##0;"-¥"#,##0');
  
  // 初回価格・最新価格（下落）
  sheet.getRange('G15:H19').setNumberFormat('"¥"#,##0');
  // 変動額（下落）+/-記号付き
  sheet.getRange('I15:I19').setNumberFormat('"+¥"#,##0;"-¥"#,##0');
  
  // === 高利益書籍ランキングの数値書式 ===
  // 見積価格・売却価格・利益
  sheet.getRange('G32:I41').setNumberFormat('"¥"#,##0');
  
  // === B25: 日時表示形式 ===
  sheet.getRange('B25').setNumberFormat('yyyy/mm/dd hh:mm:ss');
  
  // === 区切り線（A24:D24）===
  const separatorRange = sheet.getRange('A24:D24');
  separatorRange.setBorder(
    true, null, null, null, null, null,
    '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // === 列幅設定 ===
  sheet.setColumnWidth(1, 150);  // A列: 150px
  sheet.setColumnWidth(2, 100);  // B列: 100px
  sheet.setColumnWidth(3, 50);   // C列: 50px
  sheet.setColumnWidth(4, 80);   // D列: 80px
  sheet.setColumnWidth(5, 50);  // E列（空列）  // E列: 250px（タイトル用）
  sheet.setColumnWidth(6, 250);  // F列  // F列: 100px
  sheet.setColumnWidth(8, 100);  // H列
  sheet.setColumnWidth(9, 100);  // I列
  
  // === 行の高さ調整 ===
  sheet.setRowHeight(1, 50);   // ヘッダー行を高く
  sheet.setRowHeight(28, 40);  // 高利益書籍ランキングヘッダー
  
  // === グリッド線を非表示 ===
  sheet.setHiddenGridlines(true);
  
  // === 条件付き書式: 価格上昇は緑背景 ===
  const increaseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#d4edda')
    .setRanges([sheet.getRange('I7:I11')])
    .build();
  
  // === 条件付き書式: 価格下落は赤背景 ===
  const decreaseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setBackground('#f8d7da')
    .setRanges([sheet.getRange('I15:I19')])
    .build();
  
  // === 条件付き書式: 高利益は金色背景 ===
  const highProfitRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(200)
    .setBackground('#fff9e6')
    .setRanges([sheet.getRange('I32:I41')])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(increaseRule);
  rules.push(decreaseRule);
  rules.push(highProfitRule);
  sheet.setConditionalFormatRules(rules);
  
  logInfo('書式設定完了');
}

/**
 * テスト用: ダッシュボード関数の動作確認
 */
function testDashboardFunctions() {
  Logger.log('=== ダッシュボード関数テスト ===');
  
  Logger.log(`総買取冊数: ${getTotalBuyCount()}`);
  Logger.log(`総利益: ${getTotalProfit()}`);
  Logger.log(`最高利益: ${getMaxProfit()}`);
  Logger.log(`今月買取冊数: ${getMonthlyBuyCount()}`);
  Logger.log(`今月利益: ${getMonthlyProfit()}`);
  
  Logger.log('=== テスト完了 ===');
}
