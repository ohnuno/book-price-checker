/**
 * キャンペーン情報取得スクリプト
 * 
 * ファイル名: CampaignFetcher.gs
 */

/**
 * キャンペーン情報を取得してダッシュボードに表示
 */
function fetchCampaignInfo() {
  try {
    logInfo('キャンペーン情報の取得を開始');
    
    const url = 'https://www.charibon.jp/news/';
    
    // ウェブページを取得
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    
    if (response.getResponseCode() !== 200) {
      logError(`キャンペーンページの取得に失敗: ${response.getResponseCode()}`);
      return null;
    }
    
    const html = response.getContentText();
    
    // キャンペーン記事を抽出
    const campaigns = extractCampaigns(html);
    
    if (campaigns.length === 0) {
      logInfo('該当するキャンペーンが見つかりませんでした');
      return {
        hasCampaign: false,
        message: '現在実施中のキャンペーンはありません',
        lastUpdate: new Date()
      };
    }
    
    logInfo(`${campaigns.length}件のキャンペーンを検出`);
    
    // ダッシュボードに書き込み
    writeCampaignToDashboard(campaigns);
    
    return {
      hasCampaign: true,
      campaigns: campaigns,
      lastUpdate: new Date()
    };
    
  } catch (error) {
    logError(`fetchCampaignInfo エラー: ${error.message}`);
    return null;
  }
}

/**
 * HTMLからキャンペーン記事を抽出
 * @param {string} html - HTMLコンテンツ
 * @returns {Array} キャンペーン情報の配列
 */
function extractCampaigns(html) {
  const campaigns = [];
  
  try {
    // 記事タイトルとリンクを抽出するパターン
    // charibon.jpのニュースページの構造に合わせて調整が必要
    
    // パターン1: <a>タグ内のテキストを検索
    const titlePattern = /<a[^>]*href="([^"]*)"[^>]*>([^<]*(?:キャンペーン|寄付)[^<]*)<\/a>/gi;
    let match;
    
    while ((match = titlePattern.exec(html)) !== null) {
      const link = match[1];
      const title = match[2].trim();
      
      // 重複チェック
      if (!campaigns.some(c => c.title === title)) {
        campaigns.push({
          title: title,
          link: link.startsWith('http') ? link : `https://www.charibon.jp${link}`,
          date: extractDateFromTitle(title)
        });
      }
    }
    
    // パターン2: <h2>や<h3>タグ内のテキスト
    const headingPattern = /<h[23][^>]*>([^<]*(?:キャンペーン|寄付)[^<]*)<\/h[23]>/gi;
    
    while ((match = headingPattern.exec(html)) !== null) {
      const title = match[1].trim();
      
      // 重複チェック
      if (!campaigns.some(c => c.title === title)) {
        campaigns.push({
          title: title,
          link: 'https://www.charibon.jp/news/',
          date: extractDateFromTitle(title)
        });
      }
    }
    
  } catch (error) {
    logError(`extractCampaigns エラー: ${error.message}`);
  }
  
  return campaigns;
}

/**
 * タイトルから日付を抽出
 * @param {string} title - 記事タイトル
 * @returns {string} 日付文字列
 */
function extractDateFromTitle(title) {
  // YYYY/MM/DD または YYYY-MM-DD 形式を検索
  const datePattern = /(\d{4})[-/](\d{1,2})[-/](\d{1,2})/;
  const match = title.match(datePattern);
  
  if (match) {
    return `${match[1]}/${match[2].padStart(2, '0')}/${match[3].padStart(2, '0')}`;
  }
  
  return '';
}

/**
 * キャンペーン情報をダッシュボードに書き込み
 * @param {Array} campaigns - キャンペーン情報の配列
 */
function writeCampaignToDashboard(campaigns) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
    
    if (!dashboardSheet) {
      logError('ダッシュボードシートが見つかりません');
      return;
    }
    
    // キャンペーン情報を書き込む位置（右下：I列以降）
    const startRow = 5;
    const startCol = 9; // I列
    
    // ヘッダー
    dashboardSheet.getRange(startRow, startCol, 1, 2).setValues([['🎁 キャンペーン情報', '']]);
    dashboardSheet.getRange(startRow, startCol, 1, 2).merge();
    dashboardSheet.getRange(startRow, startCol).setBackground('#fff3cd');
    dashboardSheet.getRange(startRow, startCol).setFontWeight('bold');
    dashboardSheet.getRange(startRow, startCol).setFontSize(12);
    dashboardSheet.getRange(startRow, startCol).setHorizontalAlignment('center');
    
    let currentRow = startRow + 2;
    
    if (campaigns.length === 0) {
      // キャンペーンなし
      dashboardSheet.getRange(currentRow, startCol, 1, 2).setValues([['現在実施中のキャンペーンはありません', '']]);
      dashboardSheet.getRange(currentRow, startCol, 1, 2).merge();
    } else {
      // セクションヘッダー
      dashboardSheet.getRange(currentRow, startCol, 1, 2).setValues([['【現在実施中のキャンペーン】', '']]);
      dashboardSheet.getRange(currentRow, startCol, 1, 2).merge();
      dashboardSheet.getRange(currentRow, startCol).setFontWeight('bold');
      dashboardSheet.getRange(currentRow, startCol).setBackground('#f3f3f3');
      currentRow++;
      
      // 空行
      currentRow++;
      
      // キャンペーン一覧
      campaigns.forEach((campaign, index) => {
        if (index < 5) { // 最大5件まで表示
          // タイトル
          dashboardSheet.getRange(currentRow, startCol, 1, 2).setValues([[`◆ ${campaign.title}`, '']]);
          dashboardSheet.getRange(currentRow, startCol, 1, 2).merge();
          dashboardSheet.getRange(currentRow, startCol).setFontWeight('bold');
          currentRow++;
          
          // 日付（あれば）
          if (campaign.date) {
            dashboardSheet.getRange(currentRow, startCol, 1, 2).setValues([[`  期間: ${campaign.date}`, '']]);
            dashboardSheet.getRange(currentRow, startCol, 1, 2).merge();
            currentRow++;
          }
          
          // リンク
          const linkFormula = `=HYPERLINK("${campaign.link}", "  [詳細を見る]")`;
          dashboardSheet.getRange(currentRow, startCol).setFormula(linkFormula);
          dashboardSheet.getRange(currentRow, startCol).setFontColor('#1155cc');
          currentRow++;
          
          // 空行
          currentRow++;
        }
      });
    }
    
    // 最終更新
    currentRow++;
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
    dashboardSheet.getRange(currentRow, startCol, 1, 2).setValues([[`📢 最終更新: ${now}`, '']]);
    dashboardSheet.getRange(currentRow, startCol, 1, 2).merge();
    dashboardSheet.getRange(currentRow, startCol).setFontSize(9);
    dashboardSheet.getRange(currentRow, startCol).setFontColor('#666666');
    
    logInfo('キャンペーン情報をダッシュボードに書き込みました');
    
  } catch (error) {
    logError(`writeCampaignToDashboard エラー: ${error.message}`);
  }
}

/**
 * 手動でキャンペーン情報を更新
 */
function updateCampaignInfo() {
  try {
    const result = fetchCampaignInfo();
    
    if (result) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'キャンペーン情報を更新しました',
        '✅ 更新完了',
        3
      );
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'キャンペーン情報の取得に失敗しました',
        '❌ エラー',
        3
      );
    }
    
  } catch (error) {
    logError(`updateCampaignInfo エラー: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'エラー',
      `キャンペーン情報の更新中にエラーが発生しました: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 毎日自動実行用（時間主導型トリガーで設定）
 */
function dailyCampaignUpdate() {
  fetchCampaignInfo();
}