"""
Amazon 住所追加モード
- ログイン状態をチェック（未ログインならログイン処理）
- 既存の住所があれば削除
- 新しい住所を追加
"""

import argparse
import json
import os
import sys
import time
import re
import random
import imaplib
import email
from email.header import decode_header
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


class AmazonAddy:
    """Amazon住所追加クラス"""
    
    # URL
    BASE_URL = "https://www.amazon.co.jp"
    LOGIN_URL = "https://www.amazon.co.jp/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.co.jp%2F%3Fref_%3Dnav_signin&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=jpflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0"
    ORDER_HISTORY_URL = "https://www.amazon.co.jp/gp/css/order-history?ref_=nav_orders_first"
    ADDRESS_URL = "https://www.amazon.co.jp/a/addresses"
    
    def __init__(self, task_data, settings=None):
        """
        初期化
        
        Args:
            task_data: Excelから読み込んだ全カラムのデータ（辞書）
            settings: GUIから渡される設定（辞書）
        """
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
        
        # 住所情報
        self.last_name = task_data.get("LastName", "")
        self.first_name = task_data.get("FirstName", "")
        self.full_name = f"{self.last_name} {self.first_name}".strip()
        self.tell = task_data.get("Tell", "")
        self.zipcode = task_data.get("Zipcode", "")
        self.state = task_data.get("State", "")  # 都道府県
        self.city = task_data.get("City", "")  # 市区町村
        self.address1 = task_data.get("Address1", "")  # 住所1
        self.address2 = task_data.get("Address2", "")  # 住所2（丁目・番地）
        
        # 設定ディレクトリを取得（exe化対応）
        self._settings_dir = self._get_settings_dir()
        
        # IMAP設定（メールOTP取得用）
        self.imap_settings = self._load_imap_settings()
        
        # General設定（リトライ回数など）
        self.general_settings = self._load_general_settings()
        self.max_retries = self.general_settings.get("retry_count", 3)
        
        # ブラウザ関連
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        
        # クッキー保存先
        self.cookies_dir = self._get_cookies_dir()
        
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
    
    def _load_imap_settings(self):
        """IMAP設定を読み込む"""
        settings_file = self._settings_dir / "fetch_settings.json"
        
        if settings_file.exists():
            try:
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
        """プロファイル別のクッキーファイルパスを取得"""
        safe_profile = "".join(c if c.isalnum() else "_" for c in self.profile)
        if not safe_profile:
            safe_profile = "default"
        return self.cookies_dir / f"{safe_profile}_cookies.json"
    
    def _parse_proxy(self):
        """プロキシ文字列をPlaywright形式に変換"""
        if not self.proxy:
            return None
        
        parts = self.proxy.split(":")
        
        if len(parts) == 2:
            return {"server": f"http://{parts[0]}:{parts[1]}"}
        elif len(parts) == 4:
            return {
                "server": f"http://{parts[0]}:{parts[1]}",
                "username": parts[2],
                "password": parts[3]
            }
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
        """ブラウザを起動"""
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
        """ブラウザを閉じる"""
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
        """クッキーを保存"""
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
        """クッキーを読み込む"""
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
    
    # ========== 人間らしい入力 ==========
    
    def human_type(self, selector, text, min_delay=50, max_delay=150):
        """人間らしいタイピングでテキストを入力"""
        element = self.page.locator(selector)
        element.click()
        time.sleep(random.uniform(0.2, 0.5))
        
        # クリア
        element.fill("")
        time.sleep(random.uniform(0.1, 0.3))
        
        for char in text:
            element.type(char, delay=random.randint(min_delay, max_delay))
            if random.random() < 0.1:
                time.sleep(random.uniform(0.1, 0.3))
        
        time.sleep(random.uniform(0.3, 0.6))
    
    def human_click(self, selector):
        """人間らしいクリック"""
        time.sleep(random.uniform(0.3, 0.8))
        self.page.click(selector)
        time.sleep(random.uniform(0.5, 1.0))
    
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
    
    # ========== ログイン関連 ==========
    
    def is_logged_in(self):
        """ログイン状態を確認"""
        try:
            current_url = self.page.url
            
            # ログインページにリダイレクトされた場合
            if "/ap/signin" in current_url or "/ap/register" in current_url:
                return False
            
            # パスワード入力欄があればログインが必要
            password_input = self.page.query_selector("#ap_password")
            if password_input:
                return False
            
            # メールアドレス入力欄があればログインが必要
            email_input = self.page.query_selector("#ap_email")
            if email_input:
                return False
            
            # 注文履歴ページが表示されていればログイン済み
            if "order-history" in current_url:
                return True
            
            # ナビゲーションバーでログイン状態を確認
            account_elem = self.page.query_selector("#nav-link-accountList")
            if account_elem:
                inner_text = account_elem.inner_text()
                if "ログイン" in inner_text or "Sign in" in inner_text:
                    return False
                return True
            
            return True
            
        except Exception as e:
            return False
    
    def check_login_with_navigation(self):
        """注文履歴ページにアクセスしてログイン状態を確認"""
        for attempt in range(self.max_retries + 1):
            try:
                self.page.goto(self.ORDER_HISTORY_URL, wait_until="domcontentloaded", timeout=30000)
                self.random_sleep(2, 3)
                return self.is_logged_in()
            except Exception as e:
                if attempt < self.max_retries:
                    print(f"Retry {attempt + 1}/{self.max_retries}...")
                    time.sleep(3)
                else:
                    print(f"Failed to access Amazon after {self.max_retries} retries")
                    return None
        return None
    
    def do_login(self):
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
        if not self.login_id or not self.login_pass:
            print("Login credentials not provided")
            return False
        
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
                        return self.do_login()
                    elif login_type == "password_only":
                        return self.do_login()
                
                # メールアドレス入力
                email_selector = self._wait_for_email_field(max_wait=60)
                if not email_selector:
                    if self.page.query_selector("#ap_password"):
                        return self.do_login()
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
                    if self._check_account_locked():
                        print("Account is locked!")
                        return "locked"
                    time.sleep(2)
                
                password_input = self.page.query_selector("#ap_password")
                if not password_input:
                    print("Password field not found")
                    return False
                
                print("Entering password...")
                self.random_sleep(1, 2)
                self.human_type("#ap_password", self.login_pass)
                self.random_sleep(0.5, 1)
                
                self._check_remember_me()
                
                if not self._click_login_button():
                    return False
                
                self.random_sleep(3, 5)
                
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
            else:
                print(f"Unknown login type: {login_type}")
                return False
            
            # OTP処理
            if self._handle_otp_verification():
                print("OTP verification completed")
                self.random_sleep(1, 2)
                
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
            
            self._skip_phone_number_prompt()
            
            if self._check_account_locked():
                print("Account is locked!")
                return "locked"
            
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
        """ログイン画面のタイプを判定"""
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                account_switcher = self.page.query_selector(".cvf-account-switcher-profile-details-after-account-removed")
                if account_switcher:
                    return "account_switcher"
                
                try:
                    content = self.page.content()
                    if "アカウントの切り替え" in content:
                        return "account_switcher"
                except:
                    pass
                
                password_input = self.page.query_selector("#ap_password")
                email_input = self.page.query_selector("#ap_email_login") or self.page.query_selector("#ap_email")
                
                if password_input and not email_input:
                    return "password_only"
                
                if email_input:
                    return "email_password"
                
                current_url = self.page.url
                if "/ap/signin" in current_url or "/ap/register" in current_url:
                    time.sleep(2)
                    continue
                
                if self.is_logged_in():
                    return "already_logged_in"
                
            except:
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
            
            if self.page.query_selector("#ap_password"):
                return None
            
            time.sleep(2)
        
        return None
    
    def _wait_for_password_field(self, max_wait=60):
        """パスワード入力フィールドを待機"""
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                elem = self.page.query_selector("#ap_password")
                if elem:
                    return elem
            except:
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
            return alert is not None
        except:
            return False
    
    def _click_login_button(self):
        """ログインボタンをクリック"""
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
            if self.login_id:
                account_link = self.page.query_selector(f"a:has-text('{self.login_id}')")
                if account_link:
                    print("Found account link")
                    account_link.click()
                    self.random_sleep(2, 3)
                    return True
                
                account_elem = self.page.query_selector(f"div:has-text('{self.login_id}')")
                if account_elem:
                    parent = account_elem.query_selector("xpath=ancestor::a")
                    if parent:
                        parent.click()
                    else:
                        account_elem.click()
                    self.random_sleep(2, 3)
                    return True
            
            first_account = self.page.query_selector(".cvf-account-switcher-profile-details-after-account-removed")
            if first_account:
                parent_link = self.page.query_selector("a:has(.cvf-account-switcher-profile-details-after-account-removed)")
                if parent_link:
                    parent_link.click()
                    self.random_sleep(2, 3)
                    return True
            
            row_link = self.page.query_selector(".a-fixed-left-grid a")
            if row_link:
                row_link.click()
                self.random_sleep(2, 3)
                return True
            
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
    
    def _skip_phone_number_prompt(self):
        """電話番号追加をスキップ"""
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
        except:
            return False
    
    def _handle_password_verification(self):
        """パスワード再確認が求められた場合に対処"""
        try:
            # パスワード入力フォームがあるか確認
            password_input = self.page.query_selector("#ap_password")
            if not password_input:
                return True  # パスワード確認不要
            
            print("Password verification required...")
            self.human_type("#ap_password", self.login_pass)
            self.random_sleep(0.5, 1)
            
            # 「ログインしたままにする」チェックボックスをクリック
            try:
                remember_checkbox = self.page.query_selector("input[name='rememberMe']")
                if remember_checkbox and not remember_checkbox.is_checked():
                    self.page.click("input[name='rememberMe']")
                    self.random_sleep(0.3, 0.5)
            except:
                pass
            
            # サインインボタンをクリック
            signin_btn = self.page.query_selector("#signInSubmit")
            if signin_btn:
                self.human_click("#signInSubmit")
                self.random_sleep(2, 3)
            
            # 電話番号登録を求められた場合「後で」をクリック
            try:
                later_btn = self.page.query_selector("a:has-text('後で'), button:has-text('後で'), #ap-account-fixup-phone-skip-link")
                if later_btn:
                    later_btn.click()
                    self.random_sleep(1, 2)
            except:
                pass
            
            return True
        except Exception as e:
            print(f"Password verification error: {e}")
            return False
    
    def _handle_otp_verification(self):
        """OTP認証処理"""
        try:
            otp_selectors = [
                "#cvf-input-code",
                "#input-box-otp",
                "input[name='otpCode']",
                "#auth-mfa-otpcode",
                "input[name='code']"
            ]
            
            otp_input = None
            for selector in otp_selectors:
                try:
                    otp_input = self.page.wait_for_selector(selector, timeout=5000)
                    if otp_input:
                        break
                except:
                    continue
            
            if not otp_input:
                return True
            
            print("OTP verification required...")
            otp_code = self.fetch_otp_from_email(max_wait=120)
            
            if otp_code:
                print(f"Entering OTP: {otp_code}")
                self.human_type("#cvf-input-code", otp_code)
                self.random_sleep(1, 2)
                
                # 送信ボタン
                submit_selectors = [
                    "#cvf-submit-otp-button",
                    "input[type='submit']",
                    "button[type='submit']"
                ]
                
                for selector in submit_selectors:
                    try:
                        if self.page.query_selector(selector):
                            self.human_click(selector)
                            self.random_sleep(2, 4)
                            break
                    except:
                        continue
                
                return True
            
            return False
            
        except Exception as e:
            print(f"OTP error: {e}")
            return False
    
    # ========== メールOTP取得 ==========
    
    def fetch_otp_from_email(self, max_wait=120):
        """メールからOTPを取得（未読メールのみ、最新優先）"""
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
        
        target_email = self.login_id.lower().strip()
        print(f"Fetching OTP for: {target_email}")
        print(f"Using IMAP: {email_address}")
        
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                mail = imaplib.IMAP4_SSL(imap_server, imap_port)
                mail.login(email_address, email_password)
                mail.select("INBOX")
                
                _, message_numbers = mail.search(None, '(UNSEEN FROM "amazon")')
                
                if message_numbers[0]:
                    msg_nums = message_numbers[0].split()
                    total_count = len(msg_nums)
                    msg_nums = msg_nums[-20:] if total_count > 20 else msg_nums
                    print(f"Found {total_count} unread Amazon emails, checking latest {len(msg_nums)}")
                    
                    # 日付でソート
                    emails_with_date = []
                    for msg_num in msg_nums:
                        try:
                            _, header_data = mail.fetch(msg_num, "(BODY.PEEK[HEADER.FIELDS (DATE)])")
                            if header_data and header_data[0]:
                                header_bytes = header_data[0][1] if isinstance(header_data[0], tuple) else b''
                                header_str = header_bytes.decode('utf-8', errors='ignore')
                                date_match = re.search(r'Date:\s*(.+)', header_str, re.IGNORECASE)
                                date_str = date_match.group(1).strip() if date_match else ""
                                try:
                                    from email.utils import parsedate_to_datetime
                                    mail_date = parsedate_to_datetime(date_str)
                                except:
                                    mail_date = None
                                emails_with_date.append((msg_num, mail_date))
                        except:
                            emails_with_date.append((msg_num, None))
                    
                    emails_with_date.sort(key=lambda x: (x[1] is None, x[1] if x[1] else ""), reverse=True)
                    
                    for msg_num, mail_date in emails_with_date:
                        _, msg_data = mail.fetch(msg_num, "(BODY.PEEK[])")
                        
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                
                                to_addresses = []
                                for header in ["To", "Delivered-To", "X-Original-To", "Envelope-To"]:
                                    val = msg.get(header, "")
                                    if val:
                                        to_addresses.append(val.lower())
                                
                                body = self._get_email_body(msg)
                                body_lower = body.lower()
                                
                                is_target = any(target_email in addr for addr in to_addresses)
                                if not is_target and target_email in body_lower:
                                    is_target = True
                                
                                if is_target:
                                    otp_match = re.search(r'\b(\d{6})\b', body)
                                    if otp_match:
                                        otp = otp_match.group(1)
                                        date_info = mail_date.strftime('%Y-%m-%d %H:%M:%S') if mail_date else 'unknown'
                                        print(f"Found OTP: {otp} (mail date: {date_info})")
                                        mail.store(msg_num, '+FLAGS', '\\Seen')
                                        mail.logout()
                                        return otp
                                else:
                                    print(f"Skipping email - not for {target_email}")
                else:
                    print("No unread Amazon emails found")
                
                mail.logout()
                
            except Exception as e:
                print(f"Email error: {e}")
            
            print("Waiting for OTP email...")
            time.sleep(5)
        
        print("OTP fetch timeout")
        return None
    
    def _get_email_body(self, msg):
        """メール本文を取得"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    charset = part.get_content_charset() or 'utf-8'
                    try:
                        body = part.get_payload(decode=True).decode(charset, errors='ignore')
                    except:
                        body = str(part.get_payload(decode=True))
                    break
                elif content_type == "text/html" and not body:
                    charset = part.get_content_charset() or 'utf-8'
                    try:
                        body = part.get_payload(decode=True).decode(charset, errors='ignore')
                    except:
                        body = str(part.get_payload(decode=True))
        else:
            charset = msg.get_content_charset() or 'utf-8'
            try:
                body = msg.get_payload(decode=True).decode(charset, errors='ignore')
            except:
                body = str(msg.get_payload(decode=True))
        return body
    
    # ========== 住所操作 ==========
    
    def delete_existing_address(self):
        """既存の住所を削除（または置き換えフローを開始）"""
        try:
            # 削除ボタンがあるか確認
            delete_btn = self.page.query_selector("#ya-myab-address-delete-btn-0")
            
            if delete_btn:
                print("\x00STATUS:Address Found", flush=True)
                print("Existing address found, clicking delete...")
                print("\x00STATUS:Clicking Delete", flush=True)
                self.human_click("#ya-myab-address-delete-btn-0")
                self.random_sleep(1, 2)
                
                # ポップアップヘッダーがあるか確認（既定の住所の場合）
                popover_header = self.page.query_selector("#a-popover-header-4")
                
                if popover_header:
                    print("Default address detected - need to choose new address first")
                    # 「新しい住所を選んでください」ボタンをクリック
                    try:
                        self.page.wait_for_selector("#deleteAddressModal-0-choose-new-address-btn", timeout=5000)
                        self.human_click("#deleteAddressModal-0-choose-new-address-btn")
                        self.random_sleep(2, 3)
                        return "replace"  # 置き換えフローを示す
                    except:
                        print("Choose new address button not found")
                        return False
                else:
                    # 通常の削除確認
                    try:
                        self.page.wait_for_selector("#deleteAddressModal-0-submit-btn", timeout=5000)
                        self.human_click("#deleteAddressModal-0-submit-btn")
                        print("Address deleted")
                        self.random_sleep(2, 3)
                        return "deleted"
                    except:
                        print("Delete confirmation not found")
                        return False
            else:
                print("No existing address to delete")
                return "none"
                
        except Exception as e:
            print(f"Delete address error: {e}")
            return "none"
    
    def _fill_input(self, selector, value):
        """Angular対応の入力ヘルパー関数（人間らしいタイピング）"""
        try:
            # フォーカスを当ててクリア
            element = self.page.locator(selector)
            element.click()
            time.sleep(random.uniform(0.2, 0.4))
            
            # 既存の値をクリア
            element.fill("")
            time.sleep(random.uniform(0.1, 0.2))
            
            # 人間らしいタイピングで入力
            for char in value:
                element.type(char, delay=random.randint(50, 150))
                if random.random() < 0.1:
                    time.sleep(random.uniform(0.1, 0.3))
            
            time.sleep(random.uniform(0.2, 0.4))
            
            # Angularに変更を通知
            self.page.evaluate(f'''
                const input = document.querySelector("{selector}");
                if (input) {{
                    input.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    input.dispatchEvent(new Event('change', {{ bubbles: true }}));
                }}
            ''')
            return True
        except Exception as e:
            print(f"Input error for {selector}: {e}")
            return False
    
    def add_address_replacement_flow(self):
        """既定住所の置き換えフロー（別フォーマット）"""
        try:
            # フォームが完全に読み込まれるまで待機
            self.page.wait_for_selector("input#adr_FullName", state="visible", timeout=10000)
            self.random_sleep(1, 2)
            
            # 氏名入力
            print("\x00STATUS:Entering Name", flush=True)
            print("Entering name...")
            self._fill_input("input#adr_FullName", self.full_name)
            self.random_sleep(0.3, 0.6)
            
            # 郵便番号を分割
            zipcode_clean = self.zipcode.replace("-", "").replace("ー", "")
            zip1 = zipcode_clean[:3]
            zip2 = zipcode_clean[3:7] if len(zipcode_clean) >= 7 else zipcode_clean[3:]
            
            # 郵便番号入力
            print("\x00STATUS:Entering Zipcode", flush=True)
            print("Entering zipcode...")
            self._fill_input("input#adr_PostalCode1", zip1)
            self.random_sleep(0.2, 0.4)
            self._fill_input("input#adr_PostalCode2", zip2)
            
            self.random_sleep(1, 2)
            
            # 都道府県選択
            if self.state:
                print("\x00STATUS:Selecting State", flush=True)
                print("Selecting state...")
                try:
                    # 日本語・英語両方の都道府県名リスト
                    prefectures_jp = ["北海道", "青森県", "岩手県", "宮城県", "秋田県", "山形県", "福島県",
                                      "茨城県", "栃木県", "群馬県", "埼玉県", "千葉県", "東京都", "神奈川県",
                                      "新潟県", "富山県", "石川県", "福井県", "山梨県", "長野県", "岐阜県",
                                      "静岡県", "愛知県", "三重県", "滋賀県", "京都府", "大阪府", "兵庫県",
                                      "奈良県", "和歌山県", "鳥取県", "島根県", "岡山県", "広島県", "山口県",
                                      "徳島県", "香川県", "愛媛県", "高知県", "福岡県", "佐賀県", "長崎県",
                                      "熊本県", "大分県", "宮崎県", "鹿児島県", "沖縄県"]
                    
                    prefectures_en = ["Hokkaido", "Aomori-ken", "Iwate-ken", "Miyagi-ken", "Akita-ken", "Yamagata-ken", "Fukushima-ken",
                                      "Ibaraki-ken", "Tochigi-ken", "Gunma-ken", "Saitama-ken", "Chiba-ken", "Tokyo-to", "Kanagawa-ken",
                                      "Niigata-ken", "Toyama-ken", "Ishikawa-ken", "Fukui-ken", "Yamanashi-ken", "Nagano-ken", "Gifu-ken",
                                      "Shizuoka-ken", "Aichi-ken", "Mie-ken", "Shiga-ken", "Kyoto-fu", "Osaka-fu", "Hyogo-ken",
                                      "Nara-ken", "Wakayama-ken", "Tottori-ken", "Shimane-ken", "Okayama-ken", "Hiroshima-ken", "Yamaguchi-ken",
                                      "Tokushima-ken", "Kagawa-ken", "Ehime-ken", "Kochi-ken", "Fukuoka-ken", "Saga-ken", "Nagasaki-ken",
                                      "Kumamoto-ken", "Oita-ken", "Miyazaki-ken", "Kagoshima-ken", "Okinawa-ken"]
                    
                    # 日英マッピング
                    jp_to_en = dict(zip(prefectures_jp, prefectures_en))
                    
                    # 全ての都道府県名（日本語・英語）をリスト化
                    all_prefectures = prefectures_jp + prefectures_en
                    
                    dropdown_clicked = False
                    for pref in all_prefectures:
                        try:
                            btn = self.page.locator(f"button.myx-button-text:has-text('{pref}')").first
                            if btn.count() > 0 and btn.is_visible():
                                btn.click()
                                dropdown_clicked = True
                                self.random_sleep(0.5, 1)
                                break
                        except:
                            continue
                    
                    if dropdown_clicked:
                        # ドロップダウンが開いたら目的の都道府県を選択
                        self.random_sleep(0.3, 0.5)
                        
                        # まず日本語で試す
                        state_option = self.page.get_by_text(self.state, exact=True).first
                        if state_option.count() > 0:
                            state_option.click()
                            self.random_sleep(0.5, 1)
                        else:
                            # 英語名で試す
                            state_en = jp_to_en.get(self.state, self.state)
                            state_option = self.page.get_by_text(state_en, exact=True).first
                            if state_option.count() > 0:
                                state_option.click()
                                self.random_sleep(0.5, 1)
                except:
                    pass
            
            self.random_sleep(0.3, 0.6)
            
            # 住所入力
            if self.city:
                print("\x00STATUS:Entering City", flush=True)
                print("Entering city...")
                self._fill_input("input#adr_AddressLine1", self.city)
                self.random_sleep(0.3, 0.6)
            
            address_line2 = f"{self.address1} {self.address2}".strip()
            if address_line2:
                print("\x00STATUS:Entering Address", flush=True)
                print("Entering address...")
                self._fill_input("input#adr_AddressLine2", address_line2)
                self.random_sleep(0.3, 0.6)
            
            # 電話番号入力
            print("\x00STATUS:Entering Phone", flush=True)
            print("Entering phone...")
            self._fill_input("input#adr_PhoneNumber", self.tell)
            
            self.random_sleep(1, 2)
            
            # 更新ボタンクリック
            print("\x00STATUS:Clicking Submit", flush=True)
            print("Submitting...")
            try:
                update_btn = self.page.locator("a[id^='dialogButton_ok_myx']").first
                if update_btn.count() > 0 and update_btn.is_visible():
                    update_btn.click()
                else:
                    for text in ["更新", "Update", "Save", "OK"]:
                        update_btn = self.page.locator(f"a:has-text('{text}'), button:has-text('{text}')").first
                        if update_btn.count() > 0 and update_btn.is_visible():
                            update_btn.click()
                            break
                    else:
                        update_btn = self.page.locator("a.myx-button-primary").first
                        if update_btn.count() > 0 and update_btn.is_visible():
                            update_btn.click()
            except:
                pass
            
            self.random_sleep(3, 5)
            
            # 古い住所を削除
            print("\x00STATUS:Deleting Old", flush=True)
            print("Deleting old address...")
            try:
                self.page.wait_for_selector("#ya-myab-address-delete-btn-0", timeout=10000)
                self.human_click("#ya-myab-address-delete-btn-0")
                self.random_sleep(1, 2)
                
                self.page.wait_for_selector("#deleteAddressModal-0-submit-btn", timeout=5000)
                self.human_click("#deleteAddressModal-0-submit-btn")
                self.random_sleep(2, 3)
            except:
                pass
            
            # 追加した住所を既定に設定
            print("\x00STATUS:Setting Default", flush=True)
            print("Setting as default...")
            try:
                self.page.wait_for_selector("#ya-myab-set-default-shipping-btn-0", timeout=10000)
                self.human_click("#ya-myab-set-default-shipping-btn-0")
                self.random_sleep(2, 3)
                
                # アラートボックス内のテキストで成功判定
                try:
                    alert_box = self.page.wait_for_selector("#yaab-alert-box", timeout=10000)
                    if alert_box:
                        alert_text = alert_box.inner_text()
                        # 「デフォルト」「Default」「default」が含まれていれば成功
                        if "デフォルト" in alert_text or "default" in alert_text.lower():
                            print("\x00STATUS:Address Replaced", flush=True)
                            print("Address replaced!")
                            return True
                except:
                    pass
            except:
                pass
            
            # フォールバック: ページ内容で確認
            page_content = self.page.content()
            if "既定の住所" in page_content or "default" in page_content.lower():
                print("\x00STATUS:Address Replaced", flush=True)
                print("Address replaced!")
                return True
            
            print("\x00STATUS:Address Replaced", flush=True)
            return True
            
        except Exception as e:
            print(f"Address replacement error: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def add_new_address(self):
        """新しい住所を追加（通常フロー）"""
        try:
            # 「新しい住所を追加」をクリック
            self.page.wait_for_selector("#ya-myab-address-add-link", timeout=10000)
            self.human_click("#ya-myab-address-add-link")
            self.random_sleep(2, 3)
            
            # 氏名入力
            print("\x00STATUS:Entering Name", flush=True)
            print("Entering name...")
            self.page.wait_for_selector("#address-ui-widgets-enterAddressFullName", timeout=10000)
            self.human_type("#address-ui-widgets-enterAddressFullName", self.full_name)
            
            # 電話番号入力
            print("\x00STATUS:Entering Phone", flush=True)
            print("Entering phone...")
            self.human_type("#address-ui-widgets-enterAddressPhoneNumber", self.tell)
            
            # 郵便番号を分割
            zipcode_clean = self.zipcode.replace("-", "").replace("ー", "")
            if len(zipcode_clean) >= 7:
                zip1 = zipcode_clean[:3]
                zip2 = zipcode_clean[3:7]
            else:
                zip1 = zipcode_clean[:3] if len(zipcode_clean) >= 3 else zipcode_clean
                zip2 = zipcode_clean[3:] if len(zipcode_clean) > 3 else ""
            
            # 郵便番号入力
            print("\x00STATUS:Entering Zipcode", flush=True)
            print("Entering zipcode...")
            self.human_type("#address-ui-widgets-enterAddressPostalCodeOne", zip1)
            self.human_type("#address-ui-widgets-enterAddressPostalCodeTwo", zip2)
            
            self.random_sleep(1, 2)
            
            # 丁目、番地を入力
            print("\x00STATUS:Entering Address", flush=True)
            print("Entering address...")
            self.human_type("#address-ui-widgets-enterAddressLine2", self.address2)
            
            # 「いつもこの住所に届ける」をチェック
            try:
                default_checkbox = self.page.query_selector("#address-ui-widgets-use-as-my-default")
                if default_checkbox:
                    is_checked = default_checkbox.is_checked()
                    if not is_checked:
                        self.human_click("#address-ui-widgets-use-as-my-default")
            except:
                pass
            
            self.random_sleep(1, 2)
            
            # 「住所を追加」をクリック
            print("\x00STATUS:Clicking Submit", flush=True)
            print("Submitting...")
            self.human_click("#address-ui-widgets-form-submit-button")
            self.random_sleep(3, 5)
            
            # 成功確認
            try:
                alert_box = self.page.wait_for_selector("#yaab-alert-box", timeout=10000)
                if alert_box:
                    print("\x00STATUS:Address Added", flush=True)
                    print("Address added!")
                    return True
            except:
                pass
            
            # 別の成功パターンも確認
            page_content = self.page.content()
            if "住所が追加されました" in page_content or "address has been added" in page_content.lower():
                print("\x00STATUS:Address Added", flush=True)
                print("Address added!")
                return True
            
            return False
            
        except Exception as e:
            print(f"Add address error: {e}")
            return False
    
    # ========== メイン処理 ==========
    
    def run(self):
        """メイン実行
        
        Returns:
            tuple: (success: bool, error_status: str or None)
        """
        try:
            print("=" * 50)
            print(f"Site: {self.site} Mode: {self.mode}")
            print("=" * 50)
            
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
            
            # クッキーを読み込み
            self.load_cookies()
            
            # Step 1: ログイン状態を確認
            print("\x00STATUS:Checking Login", flush=True)
            print("Step 1: Checking login status...")
            try:
                login_status = self.check_login_with_navigation()
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
                    login_result = self.do_login()
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
            
            # Step 3: 住所ページにアクセス
            print("\x00STATUS:Opening Page", flush=True)
            print("Step 3: Opening address page...")
            try:
                self.page.goto(self.ADDRESS_URL, wait_until="domcontentloaded", timeout=30000)
                self.random_sleep(2, 3)
            except Exception as e:
                print(f"Failed to open address page: {e}")
                return False, "Failed Address Page"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            # パスワード再確認が求められた場合に対処
            self._handle_password_verification()
            
            # Step 4: 既存の住所を確認
            print("\x00STATUS:Checking Addresses", flush=True)
            print("Step 4: Checking existing addresses...")
            try:
                delete_result = self.delete_existing_address()
            except Exception as e:
                print(f"Error checking addresses: {e}")
                delete_result = None
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            if delete_result == "replace":
                # Step 5: 住所を置き換え
                print("Step 5: Replacing address...")
                try:
                    if self.add_address_replacement_flow():
                        print("\x00STATUS:Saving Cookies", flush=True)
                        print("NAVIGATION_COMPLETE")
                        print("Address replaced successfully")
                        self.save_cookies()
                        return True, None
                    else:
                        print("Failed to replace address")
                        return False, "Failed Replace"
                except Exception as e:
                    print(f"Error replacing address: {e}")
                    return False, "Failed Replace"
            else:
                # Step 5: 新しい住所を追加
                print("\x00STATUS:Adding New", flush=True)
                print("Step 5: Adding new address...")
                try:
                    if self.add_new_address():
                        print("\x00STATUS:Saving Cookies", flush=True)
                        print("NAVIGATION_COMPLETE")
                        print("Address added successfully")
                        self.save_cookies()
                        return True, None
                    else:
                        print("Failed to add address")
                        return False, "Failed Add Address"
                except Exception as e:
                    print(f"Error adding address: {e}")
                    return False, "Failed Add Address"
            
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
    """エントリーポイント"""
    parser = argparse.ArgumentParser()
    parser.add_argument("--task-data", required=True, help="Path to task data JSON file")
    args = parser.parse_args()
    
    try:
        with open(args.task_data, 'r', encoding='utf-8') as f:
            task_data = json.load(f)
    except Exception as e:
        print(f"Failed to load task data: {e}", file=sys.stderr)
        return
    
    bot = AmazonAddy(task_data)
    success, error_status = bot.run()
    
    # エラーステータスがある場合は出力（GUIが読み取る）
    if error_status:
        print(f"ADDY_ERROR:{error_status}")
    
    print(f"Bot finished with success={success}")


if __name__ == "__main__":
    main()