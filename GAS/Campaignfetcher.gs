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
    
    const newsUrl = 'https://www.charibon.jp/news/';
    
    // ニュース一覧ページを取得
    logInfo(`ニュースページを取得: ${newsUrl}`);
    const newsResponse = UrlFetchApp.fetch(newsUrl, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    
    if (newsResponse.getResponseCode() !== 200) {
      logError(`ニュースページの取得に失敗: ${newsResponse.getResponseCode()}`);
      return null;
    }
    
    const newsHtml = newsResponse.getContentText();
    logInfo(`ニュースページ取得成功 (${newsHtml.length} 文字)`);
    
    // Step 1: news-content内のh3タグから最新のキャンペーン記事を特定
    const campaignSection = findLatestCampaignSection(newsHtml);
    
    if (!campaignSection) {
      logInfo('該当するキャンペーンが見つかりませんでした');
      return {
        hasCampaign: false,
        message: '現在実施中のキャンペーンはありません',
        lastUpdate: new Date()
      };
    }
    
    logInfo(`キャンペーン記事を検出: ${campaignSection.title}`);
    
    // Step 2: セクション内のaタグからリンク先を取得
    const articleUrl = campaignSection.url;
    logInfo(`記事URL: ${articleUrl}`);
    
    // Step 3-6: 記事詳細を取得
    const campaignDetails = fetchCampaignDetailsFromArticle(articleUrl);
    
    if (!campaignDetails) {
      logWarning('キャンペーン詳細の取得に失敗しました');
      return null;
    }
    
    // Step 5: 期間を判定
    if (!campaignDetails.isActive) {
      logInfo('キャンペーン期間外のため、表示しません');
      return {
        hasCampaign: false,
        message: '現在実施中のキャンペーンはありません',
        lastUpdate: new Date()
      };
    }
    
    logInfo('有効なキャンペーンを検出しました');
    
    // Step 6: ダッシュボードに反映
    writeCampaignToDashboard({
      title: campaignSection.title,
      bannerImage: campaignDetails.bannerImage,
      content: campaignDetails.content,
      period: campaignDetails.period,
      target: campaignDetails.target,
      url: articleUrl
    });
    
    return {
      hasCampaign: true,
      campaign: campaignSection,
      lastUpdate: new Date()
    };
    
  } catch (error) {
    logError(`fetchCampaignInfo エラー: ${error.message}`);
    logError(`スタックトレース: ${error.stack}`);
    return null;
  }
}

/**
 * Step 1: ニュースページから最新のキャンペーン記事セクションを検出
 * @param {string} html - ニュースページのHTML
 * @returns {Object|null} {title, url}
 */
function findLatestCampaignSection(html) {
  try {
    logInfo('[Step 1] news-contentセクションを検索中...');
    
    // news-contentクラスを持つsectionタグを抽出
    const sectionPattern = /<section[^>]*class="[^"]*news-content[^"]*"[^>]*>([\s\S]*?)<\/section>/gi;
    let sectionMatch;
    
    while ((sectionMatch = sectionPattern.exec(html)) !== null) {
      const sectionHtml = sectionMatch[1];
      
      // section内のh3タグを検索
      const h3Pattern = /<h3[^>]*>([\s\S]*?)<\/h3>/i;
      const h3Match = sectionHtml.match(h3Pattern);
      
      if (h3Match) {
        const h3Content = h3Match[1];
        
        // h3内のテキストを取得（タグを除去）
        const titleText = h3Content.replace(/<[^>]+>/g, '').trim();
        
        logInfo(`  h3タグ検出: "${titleText}"`);
        
        // 「キャンペーン」または「寄付」を含むかチェック
        if (titleText.includes('キャンペーン') || titleText.includes('寄付')) {
          logInfo(`  ✅ キャンペーン記事を検出: "${titleText}"`);
          
          // section内のaタグからURLを取得
          const linkPattern = /<a[^>]*href="([^"]+)"[^>]*>/i;
          const linkMatch = sectionHtml.match(linkPattern);
          
          if (linkMatch) {
            let url = linkMatch[1];
            
            // 相対URLの場合は絶対URLに変換
            if (!url.startsWith('http')) {
              url = `https://www.charibon.jp${url}`;
            }
            
            logInfo(`  記事URL: ${url}`);
            
            return {
              title: titleText,
              url: url
            };
          }
        }
      }
    }
    
    logWarning('[Step 1] キャンペーン記事が見つかりませんでした');
    return null;
    
  } catch (error) {
    logError(`[Step 1] エラー: ${error.message}`);
    return null;
  }
}

/**
 * Step 3-6: 記事ページから詳細情報を取得
 * @param {string} url - 記事URL
 * @returns {Object|null} {bannerImage, content, period, target, isActive}
 */
function fetchCampaignDetailsFromArticle(url) {
  try {
    logInfo(`[Step 3] 記事ページを取得: ${url}`);
    
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true
    });
    
    if (response.getResponseCode() !== 200) {
      logError(`記事ページの取得に失敗: ${response.getResponseCode()}`);
      return null;
    }
    
    const html = response.getContentText();
    logInfo(`記事ページ取得成功 (${html.length} 文字)`);
    
    // Step 3: news-contentセクション内のfigure > img から画像URLを取得
    const bannerImage = extractBannerImageFromArticle(html);
    
    // Step 4: nuxt-contentクラスのdiv内からp タグの内容を取得
    const campaignInfo = extractCampaignInfoFromArticle(html);
    
    // Step 5: 期間を判定
    const isActive = checkIfCampaignIsActive(campaignInfo.period);
    
    return {
      bannerImage: bannerImage,
      content: campaignInfo.content,
      period: campaignInfo.period,
      target: campaignInfo.target,
      isActive: isActive
    };
    
  } catch (error) {
    logError(`fetchCampaignDetailsFromArticle エラー: ${error.message}`);
    return null;
  }
}

/**
 * Step 3: バナー画像URLを抽出
 * @param {string} html - 記事ページHTML
 * @returns {string} 画像URL
 */
function extractBannerImageFromArticle(html) {
  try {
    logInfo('[Step 3] バナー画像を検索中...');
    
    // news-contentクラスのsection内を検索
    const sectionPattern = /<section[^>]*class="[^"]*news-content[^"]*"[^>]*>([\s\S]*?)<\/section>/i;
    const sectionMatch = html.match(sectionPattern);
    
    if (sectionMatch) {
      const sectionHtml = sectionMatch[1];
      
      // figure > img を検索
      const imgPattern = /<figure[^>]*>[\s\S]*?<img[^>]*src="([^"]+)"[^>]*>[\s\S]*?<\/figure>/i;
      const imgMatch = sectionHtml.match(imgPattern);
      
      if (imgMatch) {
        let imageUrl = imgMatch[1];
        
        // 相対URLの場合は絶対URLに変換
        if (!imageUrl.startsWith('http')) {
          imageUrl = `https://www.charibon.jp${imageUrl}`;
        }
        
        logInfo(`  ✅ バナー画像URL: ${imageUrl}`);
        return imageUrl;
      }
    }
    
    logWarning('[Step 3] バナー画像が見つかりませんでした');
    return '';
    
  } catch (error) {
    logError(`[Step 3] エラー: ${error.message}`);
    return '';
  }
}

/**
 * Step 4: キャンペーン情報を抽出
 * @param {string} html - 記事ページHTML
 * @returns {Object} {content, period, target}
 */
function extractCampaignInfoFromArticle(html) {
  try {
    logInfo('[Step 4] キャンペーン情報を抽出中...');
    
    // nuxt-contentクラスのdiv内を検索
    const divPattern = /<div[^>]*class="[^"]*nuxt-content[^"]*"[^>]*>([\s\S]*?)<\/div>/i;
    const divMatch = html.match(divPattern);
    
    if (!divMatch) {
      logWarning('[Step 4] nuxt-contentが見つかりませんでした');
      return { content: '', period: '', target: '' };
    }
    
    const divHtml = divMatch[1];
    
    // pタグの内容を全て抽出
    const pPattern = /<p[^>]*>([\s\S]*?)<\/p>/gi;
    let content = '';
    let period = '';
    let target = '';
    let pMatch;
    
    while ((pMatch = pPattern.exec(divHtml)) !== null) {
      const pHtml = pMatch[1];
      const pText = pHtml.replace(/<[^>]+>/g, '').trim();
      
      logInfo(`  pタグ検出: "${pText.substring(0, 50)}..."`);
      
      // 「内容」を含むpタグからstrongタグ内を抽出
      if (!content && pText.includes('内容')) {
        const strongPattern = /<strong[^>]*>([\s\S]*?)<\/strong>/i;
        const strongMatch = pHtml.match(strongPattern);
        if (strongMatch) {
          content = strongMatch[1].replace(/<[^>]+>/g, '').trim();
          logInfo(`  ✅ 内容: "${content}"`);
        }
      }
      
      // 「期間」を含むpタグからstrongタグ内を抽出
      if (!period && pText.includes('期間')) {
        const strongPattern = /<strong[^>]*>([\s\S]*?)<\/strong>/i;
        const strongMatch = pHtml.match(strongPattern);
        if (strongMatch) {
          period = strongMatch[1].replace(/<[^>]+>/g, '').trim();
          logInfo(`  ✅ 期間: "${period}"`);
        }
      }
      
      // 「対象」を含むpタグからstrongタグ内を抽出
      if (!target && pText.includes('対象')) {
        const strongPattern = /<strong[^>]*>([\s\S]*?)<\/strong>/i;
        const strongMatch = pHtml.match(strongPattern);
        if (strongMatch) {
          target = strongMatch[1].replace(/<[^>]+>/g, '').trim();
          logInfo(`  ✅ 対象: "${target}"`);
        }
      }
    }
    
    return {
      content: content,
      period: period,
      target: target
    };
    
  } catch (error) {
    logError(`[Step 4] エラー: ${error.message}`);
    return { content: '', period: '', target: '' };
  }
}

/**
 * Step 5: キャンペーン期間を判定
 * @param {string} periodText - 期間テキスト（例: "2025.12.1(月) – 12.31(水)"）
 * @returns {boolean} 有効期間内かどうか
 */
function checkIfCampaignIsActive(periodText) {
  try {
    logInfo(`[Step 5] 期間判定: "${periodText}"`);
    
    if (!periodText) {
      logWarning('[Step 5] 期間情報がありません');
      return false;
    }
    
    // まず年付き日付パターンを抽出（YYYY.MM.DD形式）
    const fullDatePattern = /(\d{4})\.(\d{1,2})\.(\d{1,2})/g;
    const fullDates = [];
    let match;
    
    while ((match = fullDatePattern.exec(periodText)) !== null) {
      const year = parseInt(match[1]);
      const month = parseInt(match[2]);
      const day = parseInt(match[3]);
      fullDates.push({
        year: year,
        month: month,
        day: day,
        date: new Date(year, month - 1, day)
      });
      logInfo(`  検出日付（年付き）: ${year}/${month}/${day}`);
    }
    
    // 年付き日付を文字列から削除してから、年なし日付を抽出
    let remainingText = periodText.replace(/\d{4}\.\d{1,2}\.\d{1,2}/g, '');
    
    // 年なし日付パターンを抽出（MM.DD形式）
    const shortDatePattern = /(\d{1,2})\.(\d{1,2})/g;
    const shortDates = [];
    
    while ((match = shortDatePattern.exec(remainingText)) !== null) {
      const month = parseInt(match[1]);
      const day = parseInt(match[2]);
      
      // 妥当な月日かチェック
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        shortDates.push({
          month: month,
          day: day
        });
        logInfo(`  検出日付（年なし）: ${month}/${day}`);
      }
    }
    
    // 開始日・終了日を決定
    let startDate, endDate;
    
    if (fullDates.length >= 1) {
      // 最初の年付き日付を開始日とする
      startDate = fullDates[0].date;
      const baseYear = fullDates[0].year;
      
      if (fullDates.length >= 2) {
        // 2つ目の年付き日付を終了日とする
        endDate = fullDates[1].date;
        logInfo(`  終了日（年付き）: ${baseYear}/${fullDates[1].month}/${fullDates[1].day}`);
      } else if (shortDates.length >= 1) {
        // 年なし日付を終了日とする（開始日の年を使用）
        const lastShort = shortDates[shortDates.length - 1];
        endDate = new Date(baseYear, lastShort.month - 1, lastShort.day);
        logInfo(`  年なし終了日に開始日の年を適用: ${baseYear}/${lastShort.month}/${lastShort.day}`);
      } else {
        logWarning('[Step 5] 終了日が検出できませんでした');
        return false;
      }
    } else {
      logWarning('[Step 5] 開始日が検出できませんでした');
      return false;
    }
    
    // 日本時間の今日の日付を取得
    const jst = new Date();
    const today = new Date(jst.getFullYear(), jst.getMonth(), jst.getDate());
    
    logInfo(`  開始日: ${Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy/MM/dd')}`);
    logInfo(`  終了日: ${Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy/MM/dd')}`);
    logInfo(`  今日: ${Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd')}`);
    
    // 期間内かチェック
    const isActive = (today >= startDate && today <= endDate);
    
    logInfo(`  判定結果: ${isActive ? '✅ 期間内' : '❌ 期間外'}`);
    
    return isActive;
    
  } catch (error) {
    logError(`[Step 5] エラー: ${error.message}`);
    return false;
  }
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
 * Step 6: キャンペーン情報をダッシュボードに書き込み
 * @param {Object} campaign - キャンペーン情報 {title, bannerImage, content, period, target, url}
 */
function writeCampaignToDashboard(campaign) {
  try {
    logInfo('[Step 6] ダッシュボードに書き込み中...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
    
    if (!dashboardSheet) {
      logError('ダッシュボードシートが見つかりません');
      return;
    }
    
    // キャンペーン情報を書き込む位置（A~D列、Row 28以降）
    const startRow = 28;
    const startCol = 1;  // A列
    
    // 既存のキャンペーン情報をクリア（Row 28-40）
    dashboardSheet.getRange(startRow, startCol, 13, 4).clear();
    dashboardSheet.getRange(startRow, startCol, 13, 4).clearFormat();
    
    let currentRow = startRow;
    
    if (!campaign) {
      // キャンペーンなし
      dashboardSheet.getRange(currentRow, startCol, 1, 4).setValues([['現在実施中のキャンペーンはありません', '', '', '']]);
      dashboardSheet.getRange(currentRow, startCol, 1, 4).merge();
      dashboardSheet.getRange(currentRow, startCol).setHorizontalAlignment('center');
      dashboardSheet.getRange(currentRow, startCol).setFontColor('#999999');
      logInfo('[Step 6] キャンペーンなしメッセージを表示');
      return;
    }
    
    // ヘッダー
    dashboardSheet.getRange(currentRow, startCol, 1, 4).setValues([['🎁 実施中のキャンペーン', '', '', '']]);
    dashboardSheet.getRange(currentRow, startCol, 1, 4).merge();
    dashboardSheet.getRange(currentRow, startCol).setBackground('#fff3cd');
    dashboardSheet.getRange(currentRow, startCol).setFontWeight('bold');
    dashboardSheet.getRange(currentRow, startCol).setFontSize(11);
    dashboardSheet.getRange(currentRow, startCol).setHorizontalAlignment('center');
    currentRow++;
    
    // 空行
    currentRow++;
    
    // タイトル
    dashboardSheet.getRange(currentRow, startCol, 1, 4).setValues([[`◆ ${campaign.title}`, '', '', '']]);
    dashboardSheet.getRange(currentRow, startCol, 1, 4).merge();
    dashboardSheet.getRange(currentRow, startCol).setFontWeight('bold');
    dashboardSheet.getRange(currentRow, startCol).setFontSize(10);
    logInfo(`  タイトル: ${campaign.title}`);
    currentRow++;
    
    // 空行
    currentRow++;
    
    // バナー画像
    if (campaign.bannerImage) {
      try {
        // 画像を挿入（A列に配置）
        const imageFormula = `=IMAGE("${campaign.bannerImage}", 1)`;
        dashboardSheet.getRange(currentRow, startCol, 1, 4).setValues([[imageFormula, '', '', '']]);
        dashboardSheet.getRange(currentRow, startCol, 1, 4).merge();
        dashboardSheet.setRowHeight(currentRow, 120);  // 画像用の行高さ
        logInfo(`  バナー画像: ${campaign.bannerImage}`);
        currentRow++;
        
        // 画像の後に空行
        currentRow++;
      } catch (imageError) {
        logWarning(`画像挿入エラー: ${imageError.message}`);
      }
    }
    
    // 内容
    if (campaign.content) {
      dashboardSheet.getRange(currentRow, startCol).setValue('内容:');
      dashboardSheet.getRange(currentRow, startCol).setFontWeight('bold');
      dashboardSheet.getRange(currentRow, startCol + 1, 1, 3).setValues([[campaign.content, '', '']]);
      dashboardSheet.getRange(currentRow, startCol + 1, 1, 3).merge();
      logInfo(`  内容: ${campaign.content}`);
      currentRow++;
    }
    
    // 期間
    if (campaign.period) {
      dashboardSheet.getRange(currentRow, startCol).setValue('期間:');
      dashboardSheet.getRange(currentRow, startCol).setFontWeight('bold');
      dashboardSheet.getRange(currentRow, startCol + 1, 1, 3).setValues([[campaign.period, '', '']]);
      dashboardSheet.getRange(currentRow, startCol + 1, 1, 3).merge();
      logInfo(`  期間: ${campaign.period}`);
      currentRow++;
    }
    
    // 対象
    if (campaign.target) {
      dashboardSheet.getRange(currentRow, startCol).setValue('対象:');
      dashboardSheet.getRange(currentRow, startCol).setFontWeight('bold');
      dashboardSheet.getRange(currentRow, startCol + 1, 1, 3).setValues([[campaign.target, '', '']]);
      dashboardSheet.getRange(currentRow, startCol + 1, 1, 3).merge();
      logInfo(`  対象: ${campaign.target}`);
      currentRow++;
    }
    
    logInfo('[Step 6] ✅ ダッシュボードへの書き込み完了');
    
  } catch (error) {
    logError(`[Step 6] エラー: ${error.message}`);
    logError(`スタックトレース: ${error.stack}`);
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
