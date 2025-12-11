"""
ValueBooks.jpから古本買取価格を取得するスクレイパー

ISBNリストシート列構成:
- A列(1): ISBN
- B列(2): 書籍名
- C列(3): 著者
- D列(4): 初回見積価格
- E列(5): 最新見積価格
- F列(6): 価格更新日時
- G列(7): 価格増減
- H列(8): チェックボックス
"""

import time
import logging
from datetime import datetime
import pytz
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

# Google Cloud Logging setup
import google.cloud.logging
from google.cloud.logging.handlers import CloudLoggingHandler

# Setup Cloud Logging client
try:
    client = google.cloud.logging.Client()
    handler = CloudLoggingHandler(client)
    
    # Logger configuration
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)
    
    # Also output to stdout (for local testing)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    logger.info("Google Cloud Logging initialized successfully")
except Exception as e:
    # Fallback to local logging if Cloud Logging initialization fails
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    logger = logging.getLogger(__name__)
    logger.warning(f"Cloud Logging initialization failed, using standard logging: {e}")



class ValueBooksScraper:
    """Scraper to fetch used book purchase prices from ValueBooks.jp"""
    
    def __init__(self, credentials_file, headless=True):
        """
        Initialize
        
        Args:
            credentials_file: Google Sheets API credentials file
            headless: Whether to run in headless mode
        """
        self.driver = self._setup_driver(headless)
        self.sheet_client = self._setup_google_sheets(credentials_file)
    
    def _setup_driver(self, headless=True):
        """
        Setup Selenium driver
        
        Args:
            headless: Whether to run in headless mode
            
        Returns:
            WebDriver: Chrome driver
        """
        logger.info("=== CHROME DRIVER SETUP START ===")
        logger.info(f"Headless mode: {headless}")
        
        options = Options()
        
        if headless:
            options.add_argument('--headless=new')
            logger.info("[OPTION] Added: --headless=new")
        
        # Bot detection evasion settings
        logger.info("Configuring bot detection evasion options...")
        options.add_argument('--no-sandbox')
        logger.info("[OPTION] Added: --no-sandbox")
        options.add_argument('--disable-dev-shm-usage')
        logger.info("[OPTION] Added: --disable-dev-shm-usage")
        options.add_argument('--disable-blink-features=AutomationControlled')
        logger.info("[OPTION] Added: --disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        logger.info("[OPTION] Added experimental: excludeSwitches=['enable-automation']")
        options.add_experimental_option('useAutomationExtension', False)
        logger.info("[OPTION] Added experimental: useAutomationExtension=False")
        
        # Memory and crash countermeasures
        logger.info("Configuring memory and crash countermeasures...")
        options.add_argument('--disable-gpu')
        logger.info("[OPTION] Added: --disable-gpu")
        options.add_argument('--disable-software-rasterizer')
        logger.info("[OPTION] Added: --disable-software-rasterizer")
        options.add_argument('--disable-extensions')
        logger.info("[OPTION] Added: --disable-extensions")
        options.add_argument('--disable-logging')
        logger.info("[OPTION] Added: --disable-logging")
        options.add_argument('--disable-web-security')
        logger.info("[OPTION] Added: --disable-web-security")
        options.add_argument('--disable-features=VizDisplayCompositor')
        logger.info("[OPTION] Added: --disable-features=VizDisplayCompositor")
        options.add_argument('--disable-setuid-sandbox')
        logger.info("[OPTION] Added: --disable-setuid-sandbox")
        
        # Set User-Agent
        user_agent = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        options.add_argument(f'user-agent={user_agent}')
        logger.info(f"[OPTION] Set User-Agent: {user_agent}")
        
        # Set window size
        options.add_argument('--window-size=1280,720')
        logger.info("[OPTION] Set window size: 1280x720")
        
        logger.info("Creating Chrome WebDriver instance...")
        try:
            driver = webdriver.Chrome(options=options)
            logger.info("✅ Chrome WebDriver instance created successfully")
        except Exception as e:
            logger.error(f"❌ Failed to create Chrome WebDriver: {e}")
            raise
        
        # Avoid WebDriver detection
        logger.info("Applying WebDriver detection evasion script...")
        try:
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            logger.info("✅ WebDriver detection evasion script applied successfully")
        except Exception as e:
            logger.warning(f"⚠️ Failed to apply WebDriver detection evasion: {e}")
        
        logger.info("=== CHROME DRIVER SETUP COMPLETED ===")
        return driver
    
    def _get_jst_now(self):
        """
        Get current Japan time
        
        Returns:
            datetime: Japan time
        """
        jst = pytz.timezone('Asia/Tokyo')
        return datetime.now(jst)
    
    def _setup_google_sheets(self, credentials_file):
        """
        Setup Google Sheets API
        
        Args:
            credentials_file: Credentials file path
            
        Returns:
            gspread.Client: Google Sheets client
        """
        logger.info("Setting up Google Sheets API...")
        scope = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive'
        ]
        credentials = ServiceAccountCredentials.from_json_keyfile_name(credentials_file, scope)
        client = gspread.authorize(credentials)
        logger.info("Google Sheets API configured successfully")
        return client
    
    def search_isbn_estimate(self, isbn):
        """
        Search ISBN on ValueBooks.jp purchase estimate page
        
        Args:
            isbn: ISBN to search
            
        Returns:
            dict: Book information {isbn, title, author, publisher, price, price_date}
                  None if not found
        """
        logger.info(f"==============================")
        logger.info(f"ISBN PURCHASE ESTIMATE START: {isbn}")
        logger.info(f"==============================")
        
        try:
            # Access purchase estimate page
            estimate_url = "https://www.valuebooks.jp/estimate/guide"
            logger.info(f"[STEP 1] Accessing purchase estimate page: {estimate_url}")
            try:
                self.driver.get(estimate_url)
                logger.info(f"✅ Successfully navigated to: {self.driver.current_url}")
            except Exception as e:
                logger.error(f"❌ Failed to navigate to estimate page: {e}")
                raise
            
            # Wait for page to load
            logger.info("[STEP 2] Waiting for page to load (3 seconds)...")
            time.sleep(3)
            logger.info("✅ Page load wait completed")
            
            # Find ISBN input form
            try:
                logger.info("[STEP 3] Searching for ISBN input form...")
                
                # Try multiple selectors
                input_selectors = [
                    "input[placeholder*='気になる本']",
                    "input[placeholder*='検索']",
                    "input[type='search']",
                    "input[id^='input-']",
                    ".v-autocomplete input",
                    ".search-input input",
                    "input[type='text']",
                ]
                
                isbn_input = None
                for idx, selector in enumerate(input_selectors, 1):
                    try:
                        logger.info(f"  Trying selector {idx}/{len(input_selectors)}: {selector}")
                        isbn_input = WebDriverWait(self.driver, 5).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                        )
                        logger.info(f"✅ ISBN input form found with selector: {selector}")
                        break
                    except TimeoutException:
                        logger.info(f"  ⏭️  Selector {idx} failed, trying next...")
                        continue
                
                if not isbn_input:
                    logger.error("❌ Input form not found with any selector")
                    return None
                
                placeholder = isbn_input.get_attribute('placeholder')
                logger.info(f"Input form placeholder: '{placeholder}'")
                
                # Input ISBN
                logger.info("[STEP 4] Clearing input form...")
                isbn_input.clear()
                time.sleep(0.5)
                logger.info("✅ Input form cleared")
                
                # Input character by character (more human-like)
                logger.info(f"[STEP 5] Inputting ISBN character by character: {isbn}")
                for i, char in enumerate(isbn):
                    isbn_input.send_keys(char)
                    time.sleep(0.1)
                    if (i + 1) % 3 == 0:  # Log every 3 characters
                        logger.info(f"  Progress: {i + 1}/{len(isbn)} characters input")
                
                logger.info(f"✅ ISBN input completed: {isbn}")
                
                # Execute search with Enter key
                logger.info("[STEP 6] Waiting 1 second before executing search...")
                time.sleep(1)
                logger.info("Executing search with Enter key...")
                isbn_input.send_keys(Keys.RETURN)
                logger.info("✅ Enter key sent")
                
                # Wait for result page transition
                logger.info("[STEP 7] Waiting for search results (5 seconds)...")
                time.sleep(5)
                current_url = self.driver.current_url
                logger.info(f"✅ Search completed. Current URL: {current_url}")
                
                # Extract book information
                logger.info("[STEP 8] Extracting book information...")
                book_info = self._extract_estimate_result(isbn)
                
                if book_info:
                    logger.info(f"✅ Successfully retrieved: {book_info['title']} - ¥{book_info['price']}")
                else:
                    logger.warning(f"⚠️ Failed to extract book information")
                
                logger.info(f"==============================")
                return book_info
                
            except TimeoutException:
                logger.error("❌ Input form not found (timeout)")
                return None
            
        except Exception as e:
            logger.error(f"❌ Estimate error for ISBN {isbn}: {str(e)}")
            logger.error(f"Error details:", exc_info=True)
            return None
    
    def _extract_estimate_result(self, isbn):
        """
        Extract book information from estimate result page
        
        Args:
            isbn: ISBN
            
        Returns:
            dict: Book information (price=0 if no matching product)
        """
        try:
            # Log current URL
            current_url = self.driver.current_url
            logger.info(f"[EXTRACT] Starting result extraction - URL: {current_url}")
            
            # Log page source for debugging
            page_source = self.driver.page_source
            page_title = self.driver.title
            logger.info(f"[EXTRACT] Page title: {page_title}")
            logger.info(f"[EXTRACT] Page source length: {len(page_source)} characters")
            
            # Check for "No matching products found" message
            not_found_selectors = [
                ".v-card__text",
                "[class*='no-result']",
                "[class*='not-found']"
            ]
            
            logger.info("[EXTRACT] Checking for 'no matching product' message...")
            is_not_found = False
            for selector in not_found_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    logger.info(f"  Found {len(elements)} elements with selector: {selector}")
                    for element in elements:
                        # Check if element is displayed
                        if element.is_displayed():
                            text = element.text.strip()
                            if any(msg in text for msg in [
                                "該当する商品は見つかりませんでした",
                                "商品は見つかりませんでした",
                                "該当する商品がありません"
                            ]):
                                logger.warning(f"⚠️ 'No matching product' message detected: '{text}' - ISBN: {isbn}")
                                is_not_found = True
                                break
                    if is_not_found:
                        break
                except Exception as e:
                    logger.debug(f"Error during no-match check: {e}")
                    continue
            
            if is_not_found:
                logger.info("[EXTRACT] Returning 0 yen for 'no matching product'")
                return {
                    'isbn': isbn,
                    'title': f'No match (ISBN: {isbn})',
                    'author': '',
                    'publisher': '',
                    'price': 0,
                    'price_date': self._get_jst_now().strftime('%Y/%m/%d %H:%M:%S')
                }
            
            logger.info("[EXTRACT] 'No matching product' message not displayed")
            
            # Get title
            title = None
            title_selectors = [
                "h1", "h2", "h3",
                ".book-title", ".title", 
                "[class*='title']", "[class*='book']",
                ".v-card__title"
            ]
            
            logger.info("[EXTRACT] Searching for title elements...")
            for selector in title_selectors:
                try:
                    elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    logger.info(f"  Selector '{selector}': {len(elements)} elements found")
                    
                    for idx, element in enumerate(elements):
                        text = element.text.strip()
                        if text and len(text) > 3:  # 3+ characters
                            logger.info(f"    Candidate {idx+1}: '{text}' (length: {len(text)})")
                            if not title:  # Use first found
                                title = text
                                logger.info(f"✅ Adopted as title: '{title}'")
                except Exception as e:
                    logger.warning(f"  Error with selector '{selector}': {e}")
                    continue
                
                if title:
                    break
            
            # Get purchase price
            logger.info("[EXTRACT] Searching for purchase price...")
            price = 0
            
            # Find element with class="buy-price" (most accurate)
            logger.info("[EXTRACT] Searching for <span class='buy-price'>...")
            try:
                buy_price_selectors = [
                    "span.buy-price",
                    ".buy-price",
                    "[class*='buy-price']"
                ]
                
                for selector in buy_price_selectors:
                    try:
                        logger.info(f"  Trying selector: {selector}")
                        buy_price_elements = self.driver.find_elements(By.CSS_SELECTOR, selector)
                        logger.info(f"  Found {len(buy_price_elements)} elements with selector '{selector}'")
                        
                        for idx, element in enumerate(buy_price_elements):
                            if element.is_displayed():
                                element_text = element.text.strip()
                                logger.info(f"    Element {idx+1} text: '{element_text}'")
                                
                                # Extract number from text (e.g., "149円" -> 149)
                                price_match = re.search(r'(\d+)', element_text)
                                if price_match:
                                    price = int(price_match.group(1))
                                    logger.info(f"✅ Found purchase price from buy-price element: {price} yen")
                                    break
                        
                        if price > 0:
                            break
                            
                    except Exception as e:
                        logger.debug(f"  Selector '{selector}' failed: {e}")
                        continue
                
            except Exception as e:
                logger.warning(f"[EXTRACT] Error while searching for buy-price: {e}")
            
            if price == 0:
                logger.warning("[EXTRACT] ⚠️ Purchase price not found (buy-price element not found), defaulting to 0 yen")
            
            # Get book information
            book_info = {
                'isbn': isbn,
                'title': title or f'Title not found (ISBN: {isbn})',
                'author': '',
                'publisher': '',
                'price': price,
                'price_date': self._get_jst_now().strftime('%Y/%m/%d %H:%M:%S')
            }
            
            logger.info(f"[EXTRACT] Retrieved book information: '{book_info['title']}' - {book_info['price']} yen")
            return book_info
            
        except Exception as e:
            logger.error(f"[EXTRACT] Error: {str(e)}")
            logger.error(f"[EXTRACT] Error details:", exc_info=True)
            return None
    
    def update_spreadsheet(self, spreadsheet_id):
        """
        Update spreadsheet with purchase prices
        
        Args:
            spreadsheet_id: Google Spreadsheet ID
        """
        try:
            logger.info("============================================================")
            logger.info("PRICE UPDATE PROCESS START")
            logger.info("============================================================")
            logger.info(f"Spreadsheet ID: {spreadsheet_id}")
            
            # Open spreadsheet
            logger.info(f"Opening spreadsheet: {spreadsheet_id}")
            spreadsheet = self.sheet_client.open_by_key(spreadsheet_id)
            sheet = spreadsheet.worksheet('ISBNリスト')
            
            # Get all data
            records = sheet.get_all_records()
            logger.info(f"Total records retrieved: {len(records)}")
            
            # Get today's date (date part only, Japan time)
            today_date = self._get_jst_now().strftime('%Y/%m/%d')
            logger.info(f"Today's date (JST): {today_date}")
            
            # Filter records: Only process records NOT updated today
            MAX_PROCESS_COUNT = 10
            records_to_process = []
            already_updated_count = 0
            
            logger.info("Filtering records to process...")
            for idx, record in enumerate(records, start=2):  # start=2 for row number
                isbn = str(record.get('ISBN', '')).strip()
                
                if not isbn:
                    continue
                
                update_date_time = record.get('価格更新日時', '')
                
                # Check if already updated today (compare date part only)
                if update_date_time:
                    update_date_str = str(update_date_time).strip()
                    if update_date_str:
                        # Extract date part (format: "2025/12/05 12:34:56" or "2025/12/05")
                        update_date = update_date_str.split(' ')[0]
                        
                        if update_date == today_date:
                            already_updated_count += 1
                            logger.debug(f"  Row {idx} (ISBN {isbn}): Already updated today, skipping")
                            continue
                
                # Add to processing list with row number
                records_to_process.append({
                    'row': idx,
                    'record': record,
                    'isbn': isbn
                })
                
                # Limit to MAX_PROCESS_COUNT
                if len(records_to_process) >= MAX_PROCESS_COUNT:
                    logger.info(f"  Reached maximum process count ({MAX_PROCESS_COUNT}), stopping filter")
                    break
            
            # Log filtering results
            logger.info("============================================================")
            logger.info("FILTERING RESULTS")
            logger.info("============================================================")
            logger.info(f"Total records in sheet: {len(records)}")
            logger.info(f"Records already updated today: {already_updated_count}")
            logger.info(f"Records to process this run: {len(records_to_process)} (max: {MAX_PROCESS_COUNT})")
            logger.info("============================================================")
            
            # Check if all records already updated today
            if len(records_to_process) == 0:
                logger.info("============================================================")
                logger.info("✅ ALL RECORDS ALREADY UPDATED TODAY")
                logger.info("============================================================")
                logger.info(f"Today's date: {today_date}")
                logger.info(f"Total records: {len(records)}")
                logger.info(f"Already updated: {already_updated_count}")
                logger.info("All ISBNs have been updated today. Exiting early.")
                logger.info("No processing needed. Process completed successfully.")
                logger.info("============================================================")
                return  # Exit early - no processing needed
            
            update_count = 0
            error_count = 0
            failed_isbns = []  # Record failed ISBNs
            
            # Process filtered records
            for item in records_to_process:
                i = item['row']
                record = item['record']
                isbn = item['isbn']
                
                logger.info(f"Processing ({update_count + error_count + 1}/{len(records_to_process)}): ISBN {isbn} (Row {i})")
                
                # Wrap individual ISBN processing in try-except to continue even if one fails
                try:
                    # Get purchase price
                    result = self.search_isbn_estimate(isbn)
                    
                    # Memory countermeasure: Clear page after each ISBN processing
                    try:
                        if self.driver:
                            logger.info("Clearing page (memory release)...")
                            self.driver.execute_script("window.localStorage.clear();")
                            self.driver.execute_script("window.sessionStorage.clear();")
                            self.driver.delete_all_cookies()
                            logger.info("✅ Page cleanup completed")
                    except Exception as cleanup_error:
                        logger.warning(f"⚠️ Page cleanup error: {cleanup_error}")
                    
                    if result:
                        current_price = record.get('最新見積価格')
                        new_price = result['price']
                        
                        # Convert current_price to number (None if empty string or None)
                        if current_price == '' or current_price is None:
                            previous_price = None
                        else:
                            try:
                                previous_price = int(current_price)
                            except (ValueError, TypeError):
                                logger.warning(f"  Failed to convert price: '{current_price}' → treating as None")
                                previous_price = None
                        
                        logger.info(f"  Current price: {current_price}")
                        logger.info(f"  New price: {new_price}")
                        logger.info(f"  previous_price (after conversion): {previous_price} (type: {type(previous_price)})")
                        
                        # Don't overwrite title if Google Books API info exists
                        if not record.get('書籍名'):
                            if result.get('title'):
                                sheet.update_cell(i, 2, result['title'])
                        
                        # Update price
                        sheet.update_cell(i, 5, new_price)  # Column E (最新見積価格)
                        
                        # Set update datetime (Japan time)
                        update_time = self._get_jst_now().strftime('%Y/%m/%d %H:%M:%S')
                        sheet.update_cell(i, 6, update_time)  # Column F (価格更新日時)
                        
                        # Calculate price change
                        if previous_price is not None:
                            change = new_price - previous_price
                            sheet.update_cell(i, 7, change)  # Column G (価格増減)
                            logger.info(f"  → Updated: {previous_price}円 → {new_price}円 (change: {change:+d}円)")
                        else:
                            logger.info(f"  → New entry: {new_price}円")
                        
                        # Record in price history
                        logger.info(f"  _add_price_history call started")
                        try:
                            self._add_price_history(spreadsheet, result, previous_price)
                            logger.info(f"  _add_price_history call completed")
                        except Exception as history_error:
                            logger.error(f"  _add_price_history call error: {history_error}")
                            logger.error(f"  Error details:", exc_info=True)
                        
                        update_count += 1
                        logger.info(f"✅ ISBN {isbn} processed successfully")
                    else:
                        logger.error(f"❌ Failed to fetch purchase price: {isbn}")
                        error_count += 1
                        failed_isbns.append(isbn)  # Record failed ISBN
                
                except Exception as process_error:
                    # Catch any unexpected error during individual ISBN processing
                    error_type = type(process_error).__name__
                    error_message = str(process_error)
                    
                    logger.error(f"❌ Unexpected error processing ISBN {isbn}")
                    logger.error(f"   Error type: {error_type}")
                    logger.error(f"   Error message: {error_message}")
                    logger.error(f"   Error details:", exc_info=True)
                    
                    # Classify error type for better debugging
                    if 'session' in error_message.lower() or 'driver' in error_message.lower() or 'chrome' in error_message.lower():
                        logger.error(f"   → Classified as: SELENIUM_ERROR")
                    elif 'timeout' in error_message.lower():
                        logger.error(f"   → Classified as: TIMEOUT_ERROR")
                    elif 'connection' in error_message.lower() or 'network' in error_message.lower():
                        logger.error(f"   → Classified as: NETWORK_ERROR")
                    elif 'memory' in error_message.lower():
                        logger.error(f"   → Classified as: MEMORY_ERROR")
                    else:
                        logger.error(f"   → Classified as: UNKNOWN_ERROR")
                    
                    error_count += 1
                    failed_isbns.append(isbn)
                    logger.warning(f"⏭️ Skipping ISBN {isbn} and continuing to next item")
                    # Continue to next ISBN despite error
            
            # Calculate statistics
            total_count = update_count + error_count
            success_rate = (update_count / total_count * 100) if total_count > 0 else 0
            
            logger.info("============================================================")
            logger.info("PROCESS COMPLETED")
            logger.info("============================================================")
            logger.info(f"Total processed: {total_count} items")
            logger.info(f"  ✅ Success: {update_count} items ({success_rate:.1f}%)")
            logger.info(f"  ❌ Failed: {error_count} items")
            if failed_isbns:
                logger.info(f"  Failed ISBNs: {', '.join(failed_isbns)}")
            logger.info("============================================================")
            
            # Write execution summary to spreadsheet
            self._write_execution_summary(spreadsheet, update_count, error_count, failed_isbns)
            
        finally:
            # Clean up resources
            self.close()
    
    def _add_price_history(self, spreadsheet, book_info, previous_price):
        """
        Add price to price history sheet
        
        Args:
            spreadsheet: Spreadsheet object
            book_info: Book information dictionary
            previous_price: Previous price (None for first registration)
        """
        try:
            logger.info("    [PRICE HISTORY] Record processing started")
            logger.info(f"    [PRICE HISTORY] previous_price: {previous_price} (type: {type(previous_price)})")
            logger.info(f"    [PRICE HISTORY] book_info['price']: {book_info['price']} (type: {type(book_info['price'])})")
            logger.info(f"    [PRICE HISTORY] book_info['title']: {book_info['title']}")
            
            history_sheet = spreadsheet.worksheet('価格履歴')
            
            if previous_price is None:
                # First registration → Always record
                logger.info(f"    [PRICE HISTORY] First registration pattern")
                change = 0
                should_record = True
            else:
                # 2nd+ registration → Record only if price changed
                logger.info(f"    [PRICE HISTORY] 2nd+ registration pattern")
                change = book_info['price'] - previous_price
                should_record = (change != 0)
            
            logger.info(f"    [PRICE HISTORY] Price change amount: {change}円")
            logger.info(f"    [PRICE HISTORY] Should record: {should_record}")
            
            if should_record:
                row = [
                    book_info['isbn'],
                    book_info['title'],
                    book_info['price_date'],
                    book_info['price'],
                    change
                ]
                logger.info(f"    [PRICE HISTORY] Record data: {row}")
                history_sheet.append_row(row)
                logger.info(f"    [PRICE HISTORY] ✅ Record success (change: {change:+d}円)")
            else:
                logger.info(f"    [PRICE HISTORY] ⏭️ No price change (skip recording)")
                
        except Exception as e:
            logger.error(f"    [PRICE HISTORY] ❌ Error: {str(e)}")
            logger.error(f"    [PRICE HISTORY] Error details:", exc_info=True)
    
    def _write_execution_summary(self, spreadsheet, success_count, error_count, failed_isbns):
        """
        Write execution summary to Error Log sheet
        
        Args:
            spreadsheet: Spreadsheet object
            success_count: Number of successful processes
            error_count: Number of failed processes
            failed_isbns: List of failed ISBNs
        """
        try:
            logger.info("[SUMMARY] Writing execution summary to spreadsheet...")
            
            error_log_sheet = spreadsheet.worksheet('エラーログ')
            
            # Get current Japan time
            execution_time = self._get_jst_now().strftime('%Y/%m/%d %H:%M:%S')
            
            # Calculate totals
            total_count = success_count + error_count
            success_rate = (success_count / total_count * 100) if total_count > 0 else 0
            
            # Create summary text
            summary = f"Processed: {success_count}/{total_count} ({success_rate:.1f}%)"
            
            # Create failed ISBNs text (comma-separated)
            failed_isbns_text = ', '.join(failed_isbns) if failed_isbns else 'None'
            
            # Append summary row to Error Log sheet
            # Format: [実行日時, 処理件数, 成功件数, 失敗件数, 成功率, 失敗ISBN]
            summary_row = [
                execution_time,
                total_count,
                success_count,
                error_count,
                f"{success_rate:.1f}%",
                failed_isbns_text
            ]
            
            logger.info(f"[SUMMARY] Summary data: {summary_row}")
            error_log_sheet.append_row(summary_row)
            logger.info(f"[SUMMARY] ✅ Execution summary written successfully")
            
            # Log summary to console
            logger.info("============================================================")
            logger.info("EXECUTION SUMMARY")
            logger.info("============================================================")
            logger.info(f"Execution time: {execution_time}")
            logger.info(f"Total processed: {total_count}")
            logger.info(f"Success: {success_count}")
            logger.info(f"Failed: {error_count}")
            logger.info(f"Success rate: {success_rate:.1f}%")
            if failed_isbns:
                logger.info(f"Failed ISBNs: {failed_isbns_text}")
            logger.info("============================================================")
            
        except Exception as e:
            logger.error(f"[SUMMARY] ❌ Failed to write execution summary: {str(e)}")
            logger.error(f"[SUMMARY] Error details:", exc_info=True)
    
    def close(self):
        """Clean up resources"""
        if self.driver:
            try:
                logger.info("Closing browser...")
                self.driver.quit()
                logger.info("Browser closed successfully")
            except Exception as e:
                logger.error(f"Error while closing browser: {e}")
                # Try force close
                try:
                    self.driver.close()
                except:
                    pass


def main():
    """Main processing"""
    # Configuration
    CREDENTIALS_FILE = 'credentials.json'
    SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'  # Replace with actual spreadsheet ID
    HEADLESS = True
    
    scraper = None
    try:
        # Initialize scraper
        logger.info("Initializing scraper...")
        scraper = ValueBooksScraper(
            credentials_file=CREDENTIALS_FILE,
            headless=HEADLESS
        )
        
        # Update spreadsheet
        scraper.update_spreadsheet(SPREADSHEET_ID)
        
        logger.info("Processing completed")
        
    except Exception as e:
        logger.error(f"Error occurred: {str(e)}")
        logger.error(f"Error details:", exc_info=True)
    finally:
        if scraper:
            scraper.close()


def update_prices(request=None):
    """
    Cloud Functions entry point
    
    Args:
        request: HTTP request (not used)
        
    Returns:
        tuple: (message, status_code)
    """
    import os
    
    CREDENTIALS_FILE = 'credentials.json'
    SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
    
    if not SPREADSHEET_ID:
        logger.error("SPREADSHEET_ID environment variable not set")
        return ('Error: SPREADSHEET_ID not configured', 500)
    
    scraper = None
    try:
        logger.info("Cloud Function execution started")
        scraper = ValueBooksScraper(
            credentials_file=CREDENTIALS_FILE,
            headless=True
        )
        
        scraper.update_spreadsheet(SPREADSHEET_ID)
        logger.info("Cloud Function execution completed successfully")
        return ('Success: Prices updated successfully', 200)
        
    except Exception as e:
        logger.error(f"Cloud Function execution error: {str(e)}")
        logger.error(f"Error details:", exc_info=True)
        return (f'Error: {str(e)}', 500)
    finally:
        if scraper:
            scraper.close()


if __name__ == '__main__':
    main()
