"""
Amazon Card モード（クレジットカード登録）
- Playwrightを使用
- クッキーが保存されていればそれを読み込んでログイン状態を復元
- ログインされていなければログイン処理を実行
- 既存のカードを削除して新しいカードを登録
"""

import argparse
import json
import sys
import time
import re
import random
import imaplib
import email
from pathlib import Path
from playwright.sync_api import sync_playwright
import os


class AmazonCard:
    """Amazonクレジットカード登録クラス（Playwright版）"""
    
    BASE_URL = "https://www.amazon.co.jp"
    LOGIN_URL = "https://www.amazon.co.jp/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.co.jp%2F%3Fref_%3Dnav_signin&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=jpflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0"
    WALLET_URL = "https://www.amazon.co.jp/cpe/yourpayments/wallet"
    
    def __init__(self, task_data, settings=None):
        self.task_data = task_data
        self.settings = settings or {}
        
        # 基本情報
        self.profile = task_data.get("Profile", "")
        self.site = task_data.get("Site", "")
        self.mode = task_data.get("Mode", "")
        self.proxy = task_data.get("Proxy", "")
        self.headless = task_data.get("Headless", False)
        
        # ログイン情報
        self.login_id = task_data.get("Loginid", "")
        self.login_pass = task_data.get("Loginpass", "")
        
        # カード情報
        self.card_number = task_data.get("Cardnumber", "")
        self.card_month = task_data.get("Cardmonth", "")
        self.card_year = task_data.get("Cardyear", "")
        self.security_code = task_data.get("Securitycode", "")
        self.card_first_name = task_data.get("Cardfirstname", "")
        self.card_last_name = task_data.get("Cardlastname", "")
        self.card_holder_name = f"{self.card_first_name} {self.card_last_name}".strip()
        
        # 設定ディレクトリを取得（exe化対応）
        self._settings_dir = self._get_settings_dir()
        
        # IMAP設定
        self.imap_settings = self._load_imap_settings()
        
        # General設定（リトライ回数など）
        self.general_settings = self._load_general_settings()
        self.max_retries = self.general_settings.get("retry_count", 3)
        
        # ブラウザ関連
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        
        # ストップ制御（GUIから設定される）
        self._worker = None  # GUIのワーカー参照
        self._stop_requested = False  # ストップフラグ
        self._browser_closed = False  # ブラウザ閉じ検知
    
    def _get_settings_dir(self):
        """設定ディレクトリを取得（exe化対応）"""
        # グローバル変数APP_DIRを確認（GUI.pyから渡される）
        global APP_DIR
        if 'APP_DIR' in globals() and APP_DIR:
            settings_dir = APP_DIR / "_internal" / "settings"
            if settings_dir.exists():
                return settings_dir
            settings_dir = APP_DIR / "settings"
            if settings_dir.exists():
                return settings_dir
        
        if getattr(sys, 'frozen', False):
            base_dir = Path(sys.executable).parent
            settings_dir = base_dir / "_internal" / "settings"
            if settings_dir.exists():
                return settings_dir
            return base_dir / "settings"
        else:
            return Path(__file__).parent.parent.parent / "settings"
    
    def _load_general_settings(self):
        """General設定を読み込む"""
        settings_file = self._settings_dir / "general_settings.json"
        
        if settings_file.exists():
            try:
                with open(settings_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
        
        return {"retry_count": 3}
    
    def _get_cookies_dir(self):
        """クッキー保存ディレクトリを取得（exe化対応）"""
        # グローバル変数APP_DIRを確認（GUI.pyから渡される）
        global APP_DIR
        if 'APP_DIR' in globals() and APP_DIR:
            cookies_dir = APP_DIR / "_internal" / "cookies" / "Amazon"
            cookies_dir.mkdir(parents=True, exist_ok=True)
            return cookies_dir
        
        if getattr(sys, 'frozen', False):
            root_dir = Path(sys.executable).parent
        else:
            # Nuitka standalone対応
            exe_path = sys.executable.lower()
            if '.dist' in exe_path or 'projectwin' in exe_path:
                exe_dir = Path(sys.executable).parent
                for i in range(5):
                    check_dir = exe_dir
                    for _ in range(i):
                        check_dir = check_dir.parent
                    if (check_dir / "ProjectWIN.exe").exists():
                        root_dir = check_dir
                        break
                else:
                    root_dir = exe_dir
            else:
                root_dir = Path(__file__).resolve().parent.parent.parent
        
        cookies_dir = root_dir / "_internal" / "cookies" / "Amazon"
        cookies_dir.mkdir(parents=True, exist_ok=True)
        return cookies_dir
    
    def _get_cookie_file(self):
        cookies_dir = self._get_cookies_dir()
        safe_profile = "".join(c if c.isalnum() else "_" for c in self.profile)
        if not safe_profile:
            safe_profile = "default"
        return cookies_dir / f"{safe_profile}_cookies.json"
    
    def _load_imap_settings(self):
        try:
            settings_file = self._settings_dir / "fetch_settings.json"
            
            if settings_file.exists():
                with open(settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                accounts = data.get("accounts", [])
                for acc in accounts:
                    if acc.get("selected"):
                        return acc
                if accounts:
                    return accounts[0]
        except Exception as e:
            print(f"Failed to load IMAP settings: {e}")
        return {}
    
    def _parse_proxy(self):
        if not self.proxy:
            return None
        parts = self.proxy.split(":")
        if len(parts) == 2:
            return {"server": f"http://{parts[0]}:{parts[1]}"}
        elif len(parts) == 4:
            return {"server": f"http://{parts[0]}:{parts[1]}", "username": parts[2], "password": parts[3]}
        return None
    
    # ========== ブラウザ操作 ==========
    
    def _find_chrome_path(self):
        """システムにインストールされているChromeのパスを探す"""
        import platform
        
        if platform.system() == "Windows":
            possible_paths = [
                os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
            ]
        elif platform.system() == "Darwin":
            possible_paths = [
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            ]
        else:
            possible_paths = [
                "/usr/bin/google-chrome",
                "/usr/bin/chromium-browser",
                "/usr/bin/chromium",
            ]
        
        for path in possible_paths:
            if os.path.exists(path):
                return path
        
        return None
    
    def start_browser(self, headless=False):
        self.playwright = sync_playwright().start()
        proxy_config = self._parse_proxy()
        launch_options = {"headless": headless}
        if proxy_config:
            launch_options["proxy"] = proxy_config
            print(f"Using proxy: {proxy_config['server']}")
        
        # システムのChromeを探す
        chrome_path = self._find_chrome_path()
        if chrome_path:
            launch_options["executable_path"] = chrome_path
            print(f"Using system Chrome: {chrome_path}")
        else:
            print("System Chrome not found, using Playwright's browser")
        
        # channel="chrome"でシステムのChromeを使う（executable_pathが見つからない場合のフォールバック）
        try:
            self.browser = self.playwright.chromium.launch(**launch_options)
        except Exception as e:
            print(f"Failed to launch with options, trying channel='chrome': {e}")
            launch_options.pop("executable_path", None)
            launch_options["channel"] = "chrome"
            self.browser = self.playwright.chromium.launch(**launch_options)
        
        self.context = self.browser.new_context(
            viewport={"width": 1280, "height": 1280},
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        self.page = self.context.new_page()
        
        # ブラウザが閉じられた時の検知
        self._browser_closed = False
        self.page.on("close", self._on_page_closed)
        self.browser.on("disconnected", self._on_browser_disconnected)
        
        print("Browser started")
    
    def _on_page_closed(self):
        """ページが閉じられた時のコールバック"""
        print("Browser window closed by user")
        self._browser_closed = True
        self._stop_requested = True
    
    def _on_browser_disconnected(self):
        """ブラウザが切断された時のコールバック"""
        print("Browser disconnected")
        self._browser_closed = True
        self._stop_requested = True
    
    def close_browser(self):
        try:
            if self.context:
                try:
                    self.context.close()
                except:
                    pass
            if self.browser:
                try:
                    self.browser.close()
                except:
                    pass
            if self.playwright:
                try:
                    self.playwright.stop()
                except:
                    pass
            print("Browser closed")
        except:
            pass
    
    # ========== クッキー管理 ==========
    
    def save_cookies(self):
        try:
            cookies = self.context.cookies()
            cookie_file = self._get_cookie_file()
            with open(cookie_file, 'w', encoding='utf-8') as f:
                json.dump(cookies, f, ensure_ascii=False, indent=2)
            print(f"Cookies saved: {cookie_file}")
            return True
        except Exception as e:
            print(f"Failed to save cookies: {e}")
            return False
    
    def load_cookies(self):
        try:
            cookie_file = self._get_cookie_file()
            if cookie_file.exists():
                with open(cookie_file, 'r', encoding='utf-8') as f:
                    cookies = json.load(f)
                self.context.add_cookies(cookies)
                print(f"Cookies loaded: {cookie_file}")
                return True
        except Exception as e:
            print(f"Failed to load cookies: {e}")
        return False
    
    # ========== ユーティリティ ==========
    
    def random_sleep(self, min_sec=1, max_sec=3):
        """ランダムな待機時間（ストップ対応）"""
        total_sleep = random.uniform(min_sec, max_sec)
        interval = 0.2
        elapsed = 0
        while elapsed < total_sleep:
            if self._stop_requested:
                return False
            time.sleep(min(interval, total_sleep - elapsed))
            elapsed += interval
        return True
    
    def human_click(self, selector):
        """人間らしいクリック"""
        time.sleep(random.uniform(0.3, 0.8))
        self.page.click(selector)
        time.sleep(random.uniform(0.5, 1.0))
    
    def human_type(self, selector, text, min_delay=50, max_delay=150):
        element = self.page.locator(selector)
        element.click()
        time.sleep(random.uniform(0.2, 0.5))
        for char in text:
            element.type(char, delay=random.randint(min_delay, max_delay))
            if random.random() < 0.1:
                time.sleep(random.uniform(0.1, 0.3))
        time.sleep(random.uniform(0.3, 0.6))
    
    def human_type_element(self, element, text, min_delay=50, max_delay=150):
        element.click()
        time.sleep(random.uniform(0.2, 0.5))
        for char in text:
            element.type(char, delay=random.randint(min_delay, max_delay))
            if random.random() < 0.1:
                time.sleep(random.uniform(0.1, 0.3))
        time.sleep(random.uniform(0.3, 0.6))
    
    # ========== OTP取得 ==========
    
    def _get_email_body(self, msg):
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain" or content_type == "text/html":
                    try:
                        payload = part.get_payload(decode=True)
                        charset = part.get_content_charset() or 'utf-8'
                        body += payload.decode(charset, errors='ignore')
                    except:
                        pass
        else:
            try:
                payload = msg.get_payload(decode=True)
                charset = msg.get_content_charset() or 'utf-8'
                body = payload.decode(charset, errors='ignore')
            except:
                pass
        return body
    
    def fetch_otp_from_email(self, max_wait=120):
        if not self.imap_settings:
            print("IMAP settings not configured")
            return None
        
        imap_server = self.imap_settings.get("imap_server", "")
        imap_port = self.imap_settings.get("imap_port", 993)
        email_address = self.imap_settings.get("email", "")
        email_password = self.imap_settings.get("password", "")
        
        if not all([imap_server, email_address, email_password]):
            print("IMAP settings incomplete")
            return None
        
        print(f"Fetching OTP for account: {self.login_id}")
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                mail = imaplib.IMAP4_SSL(imap_server, imap_port)
                mail.login(email_address, email_password)
                mail.select("INBOX")
                _, message_numbers = mail.search(None, '(UNSEEN FROM "amazon")')
                
                if message_numbers[0]:
                    msg_nums = message_numbers[0].split()
                    msg_nums = msg_nums[-20:] if len(msg_nums) > 20 else msg_nums
                    
                    for msg_num in reversed(msg_nums):
                        _, msg_data = mail.fetch(msg_num, "(BODY.PEEK[])")
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                body = self._get_email_body(msg)
                                otp_match = re.search(r'\b(\d{6})\b', body)
                                if otp_match:
                                    otp = otp_match.group(1)
                                    print(f"Found OTP: {otp}")
                                    mail.store(msg_num, '+FLAGS', '\\Seen')
                                    mail.logout()
                                    return otp
                mail.logout()
            except Exception as e:
                print(f"Email error: {e}")
            
            print("Waiting for OTP email...")
            time.sleep(5)
        
        print("OTP not found within timeout")
        return None
    
    # ========== ログイン処理 ==========
    
    ORDER_HISTORY_URL = "https://www.amazon.co.jp/gp/css/order-history?ref_=nav_orders_first"
    
    def is_logged_in(self):
        """注文履歴ページにアクセスしてログイン状態を確認"""
        for attempt in range(self.max_retries + 1):
            try:
                self.page.goto(self.ORDER_HISTORY_URL, wait_until="domcontentloaded", timeout=120000)
                self.random_sleep(2, 3)
                
                # パスワード入力欄があればログインが必要
                password_input = self.page.query_selector("#ap_password")
                if password_input:
                    return False
                
                # メールアドレス入力欄があればログインが必要
                email_input = self.page.query_selector("#ap_email")
                if email_input:
                    return False
                
                # 注文履歴ページが表示されていればログイン済み
                # URLに「order-history」が含まれているか、注文履歴の要素があるか確認
                current_url = self.page.url
                if "order-history" in current_url:
                    return True
                
                # サインインページにリダイレクトされた場合
                if "signin" in current_url or "ap/signin" in current_url:
                    return False
                
                return True
            except Exception as e:
                if attempt < self.max_retries:
                    print(f"Retry {attempt + 1}/{self.max_retries}...")
                    time.sleep(3)
                else:
                    print(f"Failed to check login status after {self.max_retries} retries")
                    return None  # 接続失敗
        return None
    
    def login(self):
        """ログイン処理
        
        再ログインの種類:
        1. メールアドレス + パスワード (ログインしたままにする + 後で)
        2. パスワードのみ (ログインしたままにする + 後で)
        3. アカウント選択 + パスワード (後で のみ)
        
        Returns:
            True: ログイン成功
            False: ログイン失敗
            "locked": アカウントロック
        """
        try:
            print("Starting login process...")
            
            # 初期ロックチェック
            if self._check_account_locked():
                print("Account is locked!")
                return "locked"
            
            # 現在の画面状態を判定（最大60秒待機、Resi対応）
            login_type = self._detect_login_type(max_wait=60)
            print(f"Login type: {login_type}")
            
            if login_type == "account_switcher":
                # パターン3: アカウント選択画面
                if not self._handle_account_switcher():
                    print("Failed to select account")
                    return False
                
                # アカウント選択後、ページ遷移を待機
                self.random_sleep(3, 5)
                
                # ロックチェック
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
                # ページが安定するまで待機してからパスワード入力画面を探す
                password_input = self._wait_for_password_field(max_wait=60)
                
                if password_input:
                    print("Entering password...")
                    self.random_sleep(1, 2)
                    self.human_type("#ap_password", self.login_pass)
                    self.random_sleep(0.5, 1)
                else:
                    print("Password field not found after account selection")
                    return False
                
                # ログインボタンクリック
                if not self._click_login_button():
                    return False
                
                self.random_sleep(3, 5)
                
                # ログイン後ロックチェック
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
            elif login_type == "password_only":
                # パターン2: パスワードのみ
                print("Entering password...")
                self.random_sleep(1, 2)
                self.human_type("#ap_password", self.login_pass)
                self.random_sleep(0.5, 1)
                
                # 「ログインしたままにする」にチェック
                self._check_remember_me()
                
                # ログインボタンクリック
                if not self._click_login_button():
                    return False
                
                self.random_sleep(3, 5)
                
                # ログイン後ロックチェック
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
            elif login_type == "email_password":
                # パターン1: メールアドレス + パスワード
                # ログインページにアクセス（必要な場合）
                current_url = self.page.url
                if "/ap/signin" not in current_url:
                    for attempt in range(self.max_retries + 1):
                        try:
                            self.page.goto(self.LOGIN_URL, wait_until="domcontentloaded", timeout=120000)
                            self.random_sleep(2, 3)
                            break
                        except Exception as e:
                            if attempt < self.max_retries:
                                print(f"Retry {attempt + 1}/{self.max_retries}...")
                                time.sleep(3)
                            else:
                                print(f"Failed to access login page")
                                return False
                    
                    # ロックチェック
                    if self._check_account_locked():
                        print("Account is locked!")
                        return "locked"
                    
                    # ページ読み込み後、再度画面判定
                    login_type = self._detect_login_type(max_wait=30)
                    if login_type == "account_switcher":
                        return self.login()  # 再帰呼び出し
                    elif login_type == "password_only":
                        return self.login()  # 再帰呼び出し
                
                # メールアドレス入力
                email_selector = self._wait_for_email_field(max_wait=60)
                if not email_selector:
                    # パスワードフィールドが先に現れた場合
                    if self.page.query_selector("#ap_password"):
                        return self.login()  # 再帰呼び出し
                    print("Email field not found")
                    return False
                
                print("Entering email...")
                self.random_sleep(1, 2)
                self.human_type(email_selector, self.login_id)
                self.random_sleep(0.5, 1)
                
                # 次へボタンをクリック
                continue_btn = self.page.query_selector("#continue")
                if continue_btn:
                    self.human_click("#continue")
                self.random_sleep(2, 3)
                
                # ロックチェック
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
                # パスワード入力フィールドを待機
                for _ in range(30):
                    if self.page.query_selector("#ap_password"):
                        break
                    # ロックチェック
                    if self._check_account_locked():
                        print("Account is locked!")
                        return "locked"
                    time.sleep(2)
                
                password_input = self.page.query_selector("#ap_password")
                if not password_input:
                    print("Password field not found")
                    return False
                
                # パスワード入力
                print("Entering password...")
                self.random_sleep(1, 2)
                self.human_type("#ap_password", self.login_pass)
                self.random_sleep(0.5, 1)
                
                # 「ログインしたままにする」にチェック
                self._check_remember_me()
                
                # ログインボタンクリック
                if not self._click_login_button():
                    return False
                
                self.random_sleep(3, 5)
                
                # ログイン後ロックチェック
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
            else:
                print(f"Unknown login type: {login_type}")
                return False
            
            # OTPチェック
            if self._check_otp_required():
                print("OTP verification required...")
                
                if self.imap_settings:
                    otp = self.fetch_otp_from_email()
                    if otp:
                        print("Entering OTP...")
                        otp_input = self.page.query_selector("#auth-mfa-otpcode")
                        if not otp_input:
                            otp_input = self.page.query_selector("input[name='otpCode']")
                        if not otp_input:
                            otp_input = self.page.query_selector("#cvf-input-code")
                        
                        if otp_input:
                            self.human_type_element(otp_input, otp)
                            self.random_sleep(0.5, 1)
                            submit_btn = self.page.query_selector("#auth-signin-button")
                            if not submit_btn:
                                submit_btn = self.page.query_selector("#cvf-submit-otp-button")
                            if submit_btn:
                                submit_btn.click()
                                self.random_sleep(3, 5)
                        
                        # OTP後ロックチェック
                        if self._check_account_locked():
                            print("Account is locked!")
                            return "locked"
                    else:
                        print("Failed to get OTP from email")
                        return False
                else:
                    print("No IMAP settings configured, cannot get OTP")
                    return False
            
            # 「後で」ボタンがあればクリック
            self._click_skip_button()
            
            # 最終ロックチェック
            if self._check_account_locked():
                print("Account is locked!")
                return "locked"
            
            # ログイン確認
            self.random_sleep(2, 3)
            if self.is_logged_in():
                print("Login successful!")
                self.save_cookies()
                return True
            else:
                print("Login may have failed")
                return False
                
        except Exception as e:
            print(f"Login error: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _detect_login_type(self, max_wait=60):
        """ログイン画面のタイプを判定（Resi対応で待機）"""
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                # アカウント切り替え画面チェック
                account_switcher = self.page.query_selector(".cvf-account-switcher-profile-details-after-account-removed")
                if account_switcher:
                    return "account_switcher"
                
                # ページ内容でアカウント切り替えを判定
                try:
                    content = self.page.content()
                    if "アカウントの切り替え" in content:
                        return "account_switcher"
                except:
                    pass
                
                # パスワードのみの画面チェック
                password_input = self.page.query_selector("#ap_password")
                email_input = self.page.query_selector("#ap_email_login") or self.page.query_selector("#ap_email")
                
                if password_input and not email_input:
                    return "password_only"
                
                if email_input:
                    return "email_password"
                
                # ログインページのURLチェック
                current_url = self.page.url
                if "/ap/signin" in current_url or "/ap/register" in current_url:
                    # ページはあるがフィールドがまだ表示されていない
                    time.sleep(2)
                    continue
                
                # 既にログイン済みの可能性
                if self.is_logged_in():
                    return "already_logged_in"
                
            except Exception as e:
                pass
            
            time.sleep(2)
        
        return "unknown"
    
    def _wait_for_email_field(self, max_wait=60):
        """メールアドレス入力フィールドを待機"""
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            for selector in ["#ap_email_login", "#ap_email"]:
                elem = self.page.query_selector(selector)
                if elem:
                    return selector
            
            # パスワードフィールドが先に現れた場合は終了
            if self.page.query_selector("#ap_password"):
                return None
            
            time.sleep(2)
        
        return None
    
    def _wait_for_password_field(self, max_wait=60):
        """パスワード入力フィールドを待機（ページ遷移対応）"""
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                elem = self.page.query_selector("#ap_password")
                if elem:
                    return elem
            except Exception as e:
                # ページ遷移中のエラーは無視して再試行
                pass
            
            time.sleep(2)
        
        return None
    
    def _check_remember_me(self):
        """「ログインしたままにする」にチェック"""
        try:
            remember_me = self.page.query_selector("input[name='rememberMe']")
            if remember_me and not remember_me.is_checked():
                remember_me.click()
                print("Checked 'Keep me signed in'")
        except:
            pass
    
    def _check_account_locked(self):
        """アカウントがロックされているかチェック"""
        try:
            alert = self.page.query_selector("#alert-0")
            if alert:
                return True
            return False
        except:
            return False
    
    def _click_login_button(self):
        """ログインボタンをクリック"""
        # ログインボタンを待機（最大60秒）
        for _ in range(30):
            login_btn = self.page.query_selector("#signInSubmit")
            if login_btn:
                self.human_click("#signInSubmit")
                print("Clicked login button")
                return True
            time.sleep(2)
        
        print("Login button not found")
        return False
    
    def _handle_account_switcher(self):
        """アカウント切り替え画面を処理"""
        try:
            # 方法1: ログインIDに一致するアカウントをクリック
            if self.login_id:
                # メールアドレスを含むリンクを探す
                account_link = self.page.query_selector(f"a:has-text('{self.login_id}')")
                if account_link:
                    print("Found account link")
                    account_link.click()
                    self.random_sleep(2, 3)
                    return True
                
                # divやspan内のテキストも試す
                account_elem = self.page.query_selector(f"div:has-text('{self.login_id}')")
                if account_elem:
                    # クリック可能な親要素を探す
                    parent = account_elem.query_selector("xpath=ancestor::a")
                    if parent:
                        parent.click()
                        self.random_sleep(2, 3)
                        return True
                    else:
                        account_elem.click()
                        self.random_sleep(2, 3)
                        return True
            
            # 方法2: 最初のアカウント（既存アカウント）をクリック
            first_account = self.page.query_selector(".cvf-account-switcher-profile-details-after-account-removed")
            if first_account:
                parent_link = self.page.query_selector("a:has(.cvf-account-switcher-profile-details-after-account-removed)")
                if parent_link:
                    parent_link.click()
                    self.random_sleep(2, 3)
                    return True
            
            # 方法3: a-row内の最初のリンクをクリック
            row_link = self.page.query_selector(".a-fixed-left-grid a")
            if row_link:
                row_link.click()
                self.random_sleep(2, 3)
                return True
            
            # 方法4: 「アカウントの追加」以外のクリック可能な要素を探す
            all_links = self.page.query_selector_all("a")
            for link in all_links:
                try:
                    text = link.inner_text()
                    if text and "アカウントの追加" not in text and "ログアウト" not in text and "@" in text:
                        link.click()
                        self.random_sleep(2, 3)
                        return True
                except:
                    continue
            
            print("Could not find account to select")
            return False
            
        except Exception as e:
            print(f"Account switcher error: {e}")
            return False
    
    def _check_otp_required(self):
        """OTP入力が必要か確認"""
        try:
            otp_input = self.page.query_selector("#auth-mfa-otpcode")
            if otp_input:
                return True
            otp_input = self.page.query_selector("input[name='otpCode']")
            if otp_input:
                return True
            otp_input = self.page.query_selector("#cvf-input-code")
            if otp_input:
                return True
            return False
        except:
            return False
    
    def _click_skip_button(self):
        """「後で」ボタンがあればクリック"""
        try:
            skip_selectors = [
                "#ap-account-fixup-phone-skip-link",
                "#skip-link",
                "#cvf-skip-link",
                "input[aria-labelledby='cvf-skip-link']",
                "a:has-text('後で')",
                "a:has-text('今はしない')",
                "a:has-text('スキップ')",
                "button:has-text('後で')",
                "button:has-text('今はしない')",
                "button:has-text('スキップ')",
                "a.a-link-normal:has-text('後で')",
                "span:has-text('後で')",
            ]
            
            for selector in skip_selectors:
                try:
                    skip_btn = self.page.query_selector(selector)
                    if skip_btn:
                        self.random_sleep(1, 2)
                        skip_btn.click()
                        self.random_sleep(2, 3)
                        print("Clicked skip button")
                        return True
                except:
                    continue
            
            return False
        except Exception as e:
            return False
    
    # ========== カード登録処理 ==========
    
    def find_apx_security_frame(self, timeout=60):
        """apx-security iframeを探す"""
        for _ in range(timeout // 2):
            self.random_sleep(2, 2)
            for frame in self.page.frames:
                frame_url = frame.url or ""
                frame_name = frame.name or ""
                if "apx-security" in frame_url or "cpe/pm" in frame_url:
                    return frame
                if "ApxSecureIframe" in frame_name:
                    try:
                        inputs = frame.query_selector_all("input")
                        if len(inputs) > 3:
                            return frame
                    except:
                        pass
        return None
    
    def delete_existing_card(self):
        """既存のカードを削除"""
        
        # Resi対応: ページ読み込み待機
        self._wait_for_loading_complete()
        
        try:
            # まず「支払い方法を追加」ボタンを探す（見つかれば既存カードなし）
            add_payment_btn = self.page.query_selector(".a-button-input")
            if not add_payment_btn:
                add_payment_btn = self.page.query_selector("span:has-text('支払い方法を追加')")
            
            # 編集ボタンがあるか即座にチェック
            edit_btn = self.page.query_selector("a.apx-edit-link-button")
            if not edit_btn:
                edit_btn = self.page.query_selector("a:has-text('編集')")
            
            # 編集ボタンがなければカードなし → 即終了
            if add_payment_btn and not edit_btn:
                print("No existing card found")
                return True
            
            # 編集ボタンも追加ボタンも見つからない場合、Resi対応でリトライ
            if not edit_btn and not add_payment_btn:
                for attempt in range(15):  # 最大30秒待機
                    # 追加ボタンを探す
                    add_payment_btn = self.page.query_selector(".a-button-input")
                    if add_payment_btn:
                        # 追加ボタンが見つかったら編集ボタンも確認
                        edit_btn = self.page.query_selector("a.apx-edit-link-button")
                        if not edit_btn:
                            edit_btn = self.page.query_selector("a:has-text('編集')")
                        
                        if not edit_btn:
                            print("No existing card found")
                            return True
                        else:
                            break  # 編集ボタンが見つかったのでカード削除処理へ
                    
                    self.random_sleep(2, 2)
            
            # 編集ボタンを再取得
            if not edit_btn:
                edit_btn = self.page.query_selector("a.apx-edit-link-button")
            if not edit_btn:
                edit_btn = self.page.query_selector("xpath=/html/body/div[1]/div[3]/div/div[2]/div/div[2]/div[2]/div[3]/div/div[2]/div/div/div/div/div[2]/div[3]/div/div[1]/div/a")
            if not edit_btn:
                edit_btn = self.page.query_selector("a:has-text('編集')")
            
            if edit_btn:
                print("\x00STATUS:Existing Card Found", flush=True)
                print("Found existing card, deleting...")
                edit_btn.click()
                self.random_sleep(3, 5)
                self._wait_for_loading_complete()
                
                print("\x00STATUS:Clicking Delete", flush=True)
                # Resi対応: リトライしながら削除リンクを探す
                remove_link = None
                for attempt in range(15):  # 最大30秒待機
                    remove_link = self.page.query_selector(".apx-remove-link-button")
                    if not remove_link:
                        remove_link = self.page.query_selector("a:has-text('ウォレットから削除')")
                    if not remove_link:
                        remove_link = self.page.query_selector("a:has-text('Amazonウォレットから削除')")
                    if not remove_link:
                        remove_link = self.page.query_selector("a:has-text('削除')")
                    
                    if remove_link:
                        break
                    
                    self.random_sleep(2, 2)
                
                if remove_link:
                    remove_link.click()
                    print("Clicked 'Remove from Amazon Wallet'")
                    self.random_sleep(3, 5)
                    self._wait_for_loading_complete()
                    
                    # Resi対応: リトライしながら削除ボタンを探す
                    delete_btn = None
                    for attempt in range(15):  # 最大30秒待機
                        delete_btn = self.page.query_selector(".pmts-delete-instrument input[type='submit']")
                        if not delete_btn:
                            delete_btn = self.page.query_selector(".apx-remove-button-desktop input[type='submit']")
                        if not delete_btn:
                            delete_btn = self.page.query_selector("input.a-button-input[type='submit']")
                        if not delete_btn:
                            delete_btn = self.page.query_selector("span.a-button-primary input[type='submit']")
                        
                        if delete_btn:
                            break
                        
                        self.random_sleep(2, 2)
                    
                    if delete_btn:
                        delete_btn.click()
                        print("Clicked delete button")
                        self.random_sleep(3, 5)
                        self._wait_for_loading_complete()
                        print("Existing card deleted successfully")
                        return True
                    else:
                        print("Could not find delete button")
                else:
                    print("Could not find 'Remove from Amazon Wallet' link")
            else:
                print("No existing card found")
                return True
                
        except Exception as e:
            print(f"Error deleting card: {e}")
        
        return False
    
    def _wait_for_loading_complete(self, timeout=30):
        """ローディングスピナーが消えるまで待機"""
        try:
            # スピナーが消えるまで待機
            self.page.wait_for_selector(
                ".pmts-loading-async-widget-spinner-overlay",
                state="hidden",
                timeout=timeout * 1000
            )
        except:
            pass
        
        # 追加で少し待機
        self.random_sleep(1, 2)
    
    def add_card(self):
        """カードを追加"""
        print("Adding new card...")
        
        # ローディングスピナーが消えるまで待機
        self._wait_for_loading_complete()
        
        # お支払方法を追加をクリック
        print("\x00STATUS:Clicked Add Payment", flush=True)
        try:
            add_btn = self.page.wait_for_selector(".a-button-input", timeout=60000)
            if add_btn:
                add_btn.click()
                print("Clicked 'Add payment method' button")
            else:
                print("Could not find 'Add payment method' button")
                return False
        except Exception as e:
            print(f"Error clicking add payment button: {e}")
            return False
        
        self.random_sleep(2, 3)
        
        # ローディングスピナーが消えるまで待機
        self._wait_for_loading_complete()
        
        # クレジットまたはデビットカードを追加をクリック
        print("\x00STATUS:Clicked Add Card", flush=True)
        try:
            card_btns = self.page.query_selector_all(".a-button-input, .a-button-text")
            clicked = False
            for btn in card_btns:
                try:
                    parent = btn.evaluate("el => el.closest('.a-button') ? el.closest('.a-button').innerText : ''")
                    if "クレジット" in parent or "デビット" in parent or "カードを追加" in parent:
                        btn.click()
                        print("Clicked 'Add credit/debit card' button")
                        clicked = True
                        break
                except:
                    continue
            
            if not clicked:
                card_btn = self.page.wait_for_selector(".a-button-input", timeout=60000)
                if card_btn:
                    card_btn.click()
                else:
                    print("Could not find 'Add credit/debit card' button")
                    return False
        except Exception as e:
            print(f"Error clicking add card button: {e}")
            return False
        
        # iframeが読み込まれるまで待機（最大120秒）
        print("\x00STATUS:Waiting Card Form", flush=True)
        print("Waiting for card form iframe...")
        target_frame = None
        
        for _ in range(60):  # 60回 × 2秒 = 最大120秒
            self.random_sleep(2, 2)
            for frame in self.page.frames:
                frame_url = frame.url or ""
                frame_name = frame.name or ""
                if "apx-security" in frame_url or "cpe/pm/register" in frame_url:
                    target_frame = frame
                    break
                if "ApxSecureIframe" in frame_name:
                    try:
                        inputs = frame.query_selector_all("input")
                        if len(inputs) > 3:
                            for inp in inputs:
                                name = inp.get_attribute("name") or ""
                                placeholder = inp.get_attribute("placeholder") or ""
                                if "card" in name.lower() or "カード" in placeholder:
                                    target_frame = frame
                                    break
                    except:
                        pass
                if target_frame:
                    break
            if target_frame:
                break
        
        if not target_frame:
            print("Could not find card form iframe")
            return False
        
        print("Found card form iframe")
        self.random_sleep(5, 8)  # iframe読み込み待機をさらに延長
        
        # カード番号入力欄を探す（リトライあり、最大60秒、Frame detached対策）
        card_number_input = None
        for retry in range(30):  # 30回 × 2秒 = 最大60秒待機
            try:
                # Frame detachedの場合は再取得
                try:
                    target_frame.url  # フレームが有効か確認
                except:
                    print("Frame detached, re-acquiring...")
                    self.random_sleep(2, 3)
                    target_frame = self.find_apx_security_frame(timeout=30)
                    if not target_frame:
                        continue
                
                card_number_input = target_frame.query_selector(
                    "input[name='addCreditCardNumber'], input[placeholder*='カード番号'], input[autocomplete='cc-number'], input[data-testid='cardNumber']"
                )
                
                if not card_number_input:
                    inputs = target_frame.query_selector_all("input[type='text'], input[type='tel'], input[type='number']")
                    for inp in inputs:
                        placeholder = inp.get_attribute("placeholder") or ""
                        name = inp.get_attribute("name") or ""
                        if "card" in name.lower() or "カード" in placeholder or "number" in name.lower():
                            card_number_input = inp
                            break
                
                if card_number_input:
                    break
            except Exception as e:
                if "detached" in str(e).lower():
                    print("Frame detached, re-acquiring...")
                    self.random_sleep(2, 3)
                    target_frame = self.find_apx_security_frame(timeout=30)
            
            self.random_sleep(2, 2)
        
        if not card_number_input:
            print("Could not find card number input")
            return False
        
        # カード番号入力
        print("\x00STATUS:Entering Card Number", flush=True)
        print("Entering card number...")
        try:
            self.human_type_element(card_number_input, self.card_number)
        except Exception as e:
            if "detached" in str(e).lower():
                print("Frame detached during input, retrying...")
                return False
        self.random_sleep(0.3, 0.6)
        
        # 有効期限入力
        print("\x00STATUS:Entering Expiry", flush=True)
        try:
            expiry_input = target_frame.query_selector(
                "input[name='addCreditCardExpirationDate'], input[placeholder*='MM'], input[autocomplete='cc-exp']"
            )
            if expiry_input:
                expiry_value = f"{self.card_month}/{self.card_year[-2:]}" if len(self.card_year) == 4 else f"{self.card_month}/{self.card_year}"
                print("Entering expiry date...")
                self.human_type_element(expiry_input, expiry_value)
                self.random_sleep(0.3, 0.6)
        except Exception as e:
            if "detached" not in str(e).lower():
                print(f"Expiry input error: {e}")
        
        # セキュリティコード入力
        try:
            cvv_input = target_frame.query_selector(
                "input[name='addCreditCardVerificationNumber'], input[placeholder*='CVV'], input[placeholder*='セキュリティ'], input[autocomplete='cc-csc']"
            )
            if cvv_input:
                print("\x00STATUS:Entering Security Code", flush=True)
                print("Entering security code")
                self.human_type_element(cvv_input, self.security_code)
                self.random_sleep(0.3, 0.6)
        except Exception as e:
            if "detached" not in str(e).lower():
                print(f"CVV input error: {e}")
        
        # カード名義入力
        print("\x00STATUS:Entering Card Name", flush=True)
        try:
            name_input = target_frame.query_selector(
                "input[name='ppw-accountHolderName'], input[placeholder*='名義'], input[autocomplete='cc-name']"
            )
            if name_input:
                print("Entering card holder name...")
                self.human_type_element(name_input, self.card_holder_name)
                self.random_sleep(0.3, 0.6)
        except Exception as e:
            print(f"Name input error: {e}")
        
        self.random_sleep(1, 2)
        
        # 次へ進むボタンをクリック
        print("\x00STATUS:Clicked Next", flush=True)
        is_pattern1 = False
        try:
            # navigation-linkのaタグがあるかでパターン判定
            navigation_link = target_frame.query_selector("a[data-testid='navigation-link']")
            
            if navigation_link:
                # パターン1: 住所も同じフォームにある場合（navigation-linkあり）
                is_pattern1 = True
                next_btn = target_frame.query_selector(
                    "xpath=/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[3]/div/div/div/div/div[13]/div"
                )
            else:
                # パターン2: 住所選択が別画面の場合（navigation-linkなし）
                is_pattern1 = False
                next_btn = target_frame.query_selector(
                    "xpath=/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[3]/div/div/div/div/div[12]/div"
                )
            
            if not next_btn:
                # フォールバック: 様々なセレクタで探す
                selectors = [
                    "button[type='submit']",
                    "input[type='submit']",
                    ".a-button-primary",
                    "button:has-text('次へ')",
                    "button:has-text('追加')",
                    "span:has-text('次へ')",
                    "span:has-text('カードを追加')",
                    "div[data-testid='add-credit-card-submit-button']",
                    "[data-action='add-credit-card']",
                ]
                for selector in selectors:
                    try:
                        btn = target_frame.query_selector(selector)
                        if btn:
                            next_btn = btn
                            break
                    except:
                        continue
            
            # それでも見つからない場合、全ボタンをスキャンしてテキストで探す
            if not next_btn:
                buttons = target_frame.query_selector_all("button, span.a-button-text, div[role='button']")
                for btn in buttons:
                    try:
                        text = btn.inner_text() or ""
                        if "次へ" in text or "追加" in text or "カードを追加" in text:
                            next_btn = btn
                            break
                    except:
                        continue
            
            if next_btn:
                try:
                    next_btn.click()
                except:
                    target_frame.evaluate("btn => btn.click()", next_btn)
                print("Clicked 'Next' button")
                self.random_sleep(3, 5)
            else:
                print("Could not find 'Next' button")
                return False, False
                
        except Exception as e:
            print(f"Next button error: {e}")
            return False, False
        
        return True, is_pattern1
    
    def select_address_and_complete(self, is_pattern1=False):
        """住所を選択して完了"""
        print("Selecting address and completing...")
        self.random_sleep(3, 5)
        
        # ローディング待機
        self._wait_for_loading_complete()
        
        # パターン1の場合は住所選択をスキップしてデフォルト設定へ
        if is_pattern1:
            print("Skipping address selection...")
        else:
            # パターン2: 「請求先住所を選択」画面が表示された場合
            address_selected = False
            max_iframe_retries = 10  # iframeリロードの最大試行回数
            
            for iframe_retry in range(max_iframe_retries):
                if address_selected:
                    break
                
                if iframe_retry > 0:
                    print(f"Retrying address selection ({iframe_retry + 1}/{max_iframe_retries})...")
                
                # Resi対応: 最大60秒待機してボタンを探す
                for attempt in range(30):
                    if address_selected:
                        break
                    
                    # 各フレームで「この住所を使用」ボタンを探す
                    for frame in self.page.frames:
                        if address_selected:
                            break
                        
                        try:
                            # 方法1: name属性で検索（最も確実）
                            use_address_btn = frame.query_selector("input[name='ppw-widgetEvent:SelectAddressEvent']")
                            if use_address_btn:
                                try:
                                    print("\x00STATUS:Clicked Use Address", flush=True)
                                    use_address_btn.click()
                                    print("Clicked 'Use this address' button")
                                    self.random_sleep(3, 4)  # 画面表示を待つ
                                    
                                    # 成功判定: 「追加されました」テキストがあるかチェック
                                    # 正常: 「追加されました」テキストがある → デフォルト設定画面
                                    # エラー: 「追加されました」テキストがない → エラー画面
                                    is_success = False
                                    
                                    # 最大5秒間、1秒ごとにチェック
                                    for check_attempt in range(5):
                                        try:
                                            # 「追加されました」テキストを探す
                                            success_text = self.page.query_selector("text=追加されました")
                                            if success_text:
                                                is_success = True
                                                break
                                            
                                            # 各フレーム内でも確認
                                            for check_frame in self.page.frames:
                                                try:
                                                    success_elem = check_frame.query_selector("span:has-text('追加されました')")
                                                    if not success_elem:
                                                        success_elem = check_frame.query_selector("h1:has-text('追加されました')")
                                                    if not success_elem:
                                                        success_elem = check_frame.query_selector("text=追加されました")
                                                    if success_elem:
                                                        is_success = True
                                                        break
                                                except:
                                                    continue
                                            
                                            if is_success:
                                                break
                                                
                                            # まだ見つからない場合は少し待つ
                                            time.sleep(1)
                                        except:
                                            time.sleep(1)
                                    
                                    if is_success:
                                        # 正常に追加完了
                                        print("Card added successfully, proceeding to default setting...")
                                        address_selected = True
                                        break
                                    else:
                                        # エラーの可能性、iframeをリロードして再試行
                                        print("\x00STATUS:Widget Error Retry", flush=True)
                                        print("Widget error, reloading iframe...")
                                        
                                        try:
                                            frame_url = frame.url
                                            if frame_url and frame_url != "about:blank":
                                                frame.goto(frame_url, wait_until="domcontentloaded", timeout=30000)
                                                self.random_sleep(3, 5)
                                                self._wait_for_loading_complete()
                                        except:
                                            try:
                                                self.page.keyboard.press("F5")
                                                self.random_sleep(3, 5)
                                            except:
                                                pass
                                        
                                        break  # 内側のforループを抜けて再試行
                                except:
                                    # JavaScriptでクリック
                                    try:
                                        frame.evaluate("btn => btn.click()", use_address_btn)
                                        print("Clicked 'Use this address' button")
                                        address_selected = True
                                        break
                                    except:
                                        pass
                            
                            # 方法2: pmts-use-selected-addressクラスで検索
                            if not address_selected:
                                use_address_btn = frame.query_selector(".pmts-use-selected-address input")
                                if use_address_btn:
                                    try:
                                        use_address_btn.click()
                                        print("Clicked 'Use this address' button")
                                        address_selected = True
                                        break
                                    except:
                                        pass
                            
                            # 方法3: JavaScriptでテキスト検索してクリック
                            if not address_selected:
                                clicked = frame.evaluate("""
                                    () => {
                                        var inputs = document.querySelectorAll('input[type="submit"]');
                                        for (var i = 0; i < inputs.length; i++) {
                                            var inp = inputs[i];
                                            var labelId = inp.getAttribute('aria-labelledby');
                                            if (labelId) {
                                                var label = document.getElementById(labelId);
                                                if (label && label.textContent.includes('この住所を使用')) {
                                                    inp.click();
                                                    return true;
                                                }
                                            }
                                        }
                                        return false;
                                    }
                                """)
                                if clicked:
                                    print("Clicked 'Use this address' button")
                                    address_selected = True
                                    break
                                    
                        except:
                            continue
                    
                    if not address_selected:
                        self.random_sleep(2, 2)
            
            if address_selected:
                self.random_sleep(3, 5)
                self._wait_for_loading_complete()
            else:
                print("Address selection button not found")
        
        # デフォルトのお支払方法に設定（トグルスイッチ）
        print("\x00STATUS:Enable Default Payment", flush=True)
        try:
            target_frame = self.find_apx_security_frame(timeout=30)
            if not target_frame:
                target_frame = self.page
            
            toggle_switch = target_frame.query_selector("div[role='switch']")
            if not toggle_switch:
                toggle_switch = target_frame.query_selector("[role='switch']")
            if not toggle_switch:
                toggle_switch = target_frame.query_selector("div[aria-label*='デフォルト']")
            if not toggle_switch:
                toggle_switch = target_frame.query_selector("[data-testid='switch-knob-wrapper']")
            if not toggle_switch:
                toggle_switch = self.page.query_selector("div[role='switch']")
            
            if toggle_switch:
                is_checked = toggle_switch.get_attribute("aria-checked")
                if is_checked == "false":
                    toggle_switch.click()
                    print("Enabled default payment method toggle")
                    self.random_sleep(1, 2)
        except Exception as e:
            print(f"Default toggle error: {e}")
        
        self.random_sleep(1, 2)
        
        # 次へ進む/完了ボタンをクリック
        print("\x00STATUS:Clicked Complete", flush=True)
        try:
            complete_btn = target_frame.query_selector(
                "xpath=/html/body/div[3]/div/div[2]/div/div/div[2]/div/div[3]/div/div/div[8]/div/div[3]"
            )
            
            if not complete_btn:
                # フォールバック: 様々なセレクタで探す
                selectors = [
                    "button:has-text('次へ進む')",
                    "button:has-text('続行')",
                    "button:has-text('完了')",
                    "span:has-text('次へ進む')",
                    "span:has-text('続行')",
                    "span:has-text('完了')",
                    ".a-button-primary",
                    "button[type='submit']",
                    "input[type='submit']",
                ]
                for selector in selectors:
                    try:
                        btn = target_frame.query_selector(selector)
                        if btn:
                            complete_btn = btn
                            break
                    except:
                        continue
            
            # それでも見つからない場合、全ボタンをスキャンしてテキストで探す
            if not complete_btn:
                buttons = target_frame.query_selector_all("button, span.a-button-text, div[role='button']")
                for btn in buttons:
                    try:
                        text = btn.inner_text() or ""
                        if "次へ進む" in text or "続行" in text or "完了" in text:
                            complete_btn = btn
                            break
                    except:
                        continue
            
            if complete_btn:
                try:
                    complete_btn.click()
                except:
                    target_frame.evaluate("btn => btn.click()", complete_btn)
                print("Clicked next/complete button")
                self.random_sleep(3, 5)
        except Exception as e:
            print(f"Complete button error: {e}")
        
        # 成功確認
        print("\x00STATUS:Card Registration", flush=True)
        try:
            success_element = self.page.wait_for_selector(".a-color-success, .a-alert-success", timeout=60000)
            if success_element:
                print("Card registration successful!")
                return True
        except:
            pass
        
        page_content = self.page.content()
        if "正常に追加されました" in page_content or "successfully added" in page_content.lower():
            print("Card registration successful!")
            return True
        
        print("Could not confirm card registration success")
        return False
    
    # ========== メイン処理 ==========
    
    def run(self):
        """メイン実行
        
        Returns:
            tuple: (success: bool, error_status: str or None)
            error_status: エラー時のステータス文字列（成功時はNone）
        """
        try:
            print("=" * 50)
            print(f"Site: {self.site} Mode: {self.mode}")
            print("=" * 50)
            
            if not self.card_number:
                print("Card number is required")
                return False, "Failed No Card"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            # ブラウザ起動
            print("\x00STATUS:Starting Task", flush=True)
            try:
                self.start_browser(headless=self.headless)
            except Exception as e:
                print(f"Browser start failed: {e}")
                return False, "Failed Browser"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            self.load_cookies()
            
            # Step 1: ログイン確認
            print("\x00STATUS:Checking Login", flush=True)
            print("Step 1: Checking login status...")
            try:
                login_status = self.is_logged_in()
            except Exception as e:
                print(f"Login check failed: {e}")
                return False, "Failed Login Check"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            if login_status is None:
                print("Failed to connect")
                return False, "Failed Connection"
            elif not login_status:
                print("\x00STATUS:Logging In", flush=True)
                print("Step 2: Logging in...")
                try:
                    login_result = self.login()
                    if login_result == "locked":
                        return False, "Failed Locked"
                    elif not login_result:
                        print("Login failed")
                        return False, "Failed Login"
                except Exception as e:
                    print(f"Login error: {e}")
                    return False, "Failed Login"
            else:
                print("Already logged in")
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            # Step 3: ウォレットページにアクセス
            print("\x00STATUS:Opening Page", flush=True)
            print("Step 3: Opening wallet page...")
            try:
                self.page.goto(self.WALLET_URL, wait_until="domcontentloaded", timeout=120000)
                self.random_sleep(3, 5)
            except Exception as e:
                print(f"Failed to open wallet page: {e}")
                return False, "Failed Wallet"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            # Step 4: 既存のカードを確認・削除
            print("\x00STATUS:Checking Existing Card", flush=True)
            print("Step 4: Checking existing cards...")
            try:
                delete_result = self.delete_existing_card()
            except Exception as e:
                print(f"Error checking existing cards: {e}")
                # 続行可能なエラーなので継続
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            # Step 5: カードを追加
            print("\x00STATUS:Adding New Card", flush=True)
            print("Step 5: Adding card...")
            try:
                result = self.add_card()
                if isinstance(result, tuple):
                    card_added, is_pattern1 = result
                else:
                    card_added = result
                    is_pattern1 = False
                
                if not card_added:
                    print("Failed to add card")
                    return False, "Failed Add Card"
            except Exception as e:
                print(f"Error adding card: {e}")
                return False, "Failed Add Card"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            # Step 6: 住所選択して完了
            print("Step 6: Completing registration...")
            try:
                if not self.select_address_and_complete(is_pattern1=is_pattern1):
                    print("Failed to complete registration")
                    return False, "Failed Address"
            except Exception as e:
                print(f"Error completing registration: {e}")
                return False, "Failed Address"
            
            print("\x00STATUS:Cookies Saved", flush=True)
            self.save_cookies()
            print("NAVIGATION_COMPLETE")
            print("Card registration completed successfully")
            return True, None
            
        except Exception as e:
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
            return False, "Failed Unknown"
        finally:
            self.close_browser()


def main():
    parser = argparse.ArgumentParser(description="Amazon Card Registration Bot")
    parser.add_argument("--task-data", required=True, help="Path to task data JSON file")
    parser.add_argument("--settings", help="Path to settings JSON file")
    
    args = parser.parse_args()
    print(f"Task data file: {args.task_data}")
    
    with open(args.task_data, 'r', encoding='utf-8') as f:
        task_data = json.load(f)
    
    print("Task data loaded successfully")
    
    settings = {}
    if args.settings:
        with open(args.settings, 'r', encoding='utf-8') as f:
            settings = json.load(f)
    
    bot = AmazonCard(task_data, settings)
    success, error_status = bot.run()
    
    # エラーステータスがある場合は出力（GUIが読み取る）
    if error_status:
        print(f"CARD_ERROR:{error_status}")
    
    print(f"Bot finished with success={success}")
    # 正常終了（GUIはCARD_ERROR/NAVIGATION_COMPLETEで判断する）


if __name__ == "__main__":
    main()