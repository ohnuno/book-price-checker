"""
Cloud Functions エントリーポイント
古本買取価格調査システム
"""
import functions_framework
from book_price_fetcher import ValueBooksScraper
import os
import logging

# ログ設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@functions_framework.http
def update_prices(request):
    """
    HTTPトリガーで価格更新を実行
    
    Args:
        request: HTTPリクエスト
        
    Returns:
        tuple: (メッセージ, ステータスコード)
    """
    logger.info("=" * 60)
    logger.info("価格更新処理を開始")
    logger.info("=" * 60)
    
    # 環境変数からスプレッドシートIDを取得
    spreadsheet_id = os.environ.get('SPREADSHEET_ID')
    
    if not spreadsheet_id:
        error_msg = 'Error: SPREADSHEET_ID environment variable not set'
        logger.error(error_msg)
        return error_msg, 500
    
    logger.info(f"スプレッドシートID: {spreadsheet_id}")
    
    scraper = None
    try:
        # スクレイパーを初期化
        logger.info("スクレイパーを初期化中...")
        scraper = ValueBooksScraper(
            credentials_file='credentials.json',
            headless=True
        )
        
        # スプレッドシートを更新
        logger.info("スプレッドシート更新開始...")
        scraper.update_spreadsheet(spreadsheet_id)
        
        logger.info("=" * 60)
        logger.info("価格更新処理が完了しました")
        logger.info("=" * 60)
        
        return 'Success: Prices updated successfully', 200
        
    except Exception as e:
        error_msg = f'Error: {str(e)}'
        logger.error(error_msg)
        logger.exception("詳細なエラー情報:")
        return error_msg, 500
        
    finally:
        if scraper:
            scraper.close()
            logger.info("リソースをクリーンアップしました")
