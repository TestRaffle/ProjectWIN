"""
Amazon 抽選(Raffle)モード
- ログイン状態をチェック（未ログインならログイン処理）
- URLにアクセスして抽選ボタンをクリック
- 複数URL対応（カンマ区切り）
- ハイブリッド型: requestsで抽選参加（データ転送量削減）
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
import requests
from bs4 import BeautifulSoup


def get_app_dir():
    """アプリのルートディレクトリを取得（exe化・import両対応）"""
    # GUIからimportされた場合、APP_DIRがグローバルに設定される
    if 'APP_DIR' in globals() and globals()['APP_DIR']:
        return globals()['APP_DIR']
    # 直接実行された場合
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    else:
        # bots/amazon/amazon_raffle.py -> 3階層上がルート
        return Path(__file__).resolve().parent.parent.parent


# アプリのルートディレクトリ
APP_DIR = get_app_dir()


class AmazonRaffle:
    """Amazon抽選クラス"""
    
    # URL
    BASE_URL = "https://www.amazon.co.jp"
    LOGIN_URL = "https://www.amazon.co.jp/ap/signin?openid.pape.max_auth_age=0&openid.return_to=https%3A%2F%2Fwww.amazon.co.jp%2F%3Fref_%3Dnav_signin&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.assoc_handle=jpflex&openid.mode=checkid_setup&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0"
    ORDER_HISTORY_URL = "https://www.amazon.co.jp/gp/css/order-history?ref_=nav_orders_first"
    
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
        
        # URL（カンマ区切りで複数対応）
        self.urls = self._parse_urls(task_data.get("URL", ""))
        
        # IMAP設定（メールOTP取得用）
        self.imap_settings = self._load_imap_settings()
        
        # general_settings読み込み（リトライ回数）
        self.general_settings = self._load_general_settings()
        self.max_retries = self.general_settings.get("retry_count", 3)
        
        # ブラウザ関連
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        
        # クッキー保存先
        self.cookies_dir = self._get_cookies_dir()
        
        # 結果格納
        self.results = []  # [{url, status, title}, ...]
        
        # ストップ制御（GUIから設定される）
        self._worker = None  # GUIのワーカー参照
        self._stop_requested = False  # ストップフラグ
        self._browser_closed = False  # ブラウザ閉じ検知
    
    def _load_general_settings(self):
        """general_settings.jsonを読み込む"""
        try:
            settings_path = APP_DIR / "settings" / "general_settings.json"
            if settings_path.exists():
                with open(settings_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Failed to load general settings: {e}")
        return {}
    
    def _parse_urls(self, url_string):
        """URLをパース（カンマ区切り対応）"""
        if not url_string:
            return []
        
        # カンマで分割してトリム
        urls = [u.strip() for u in url_string.split(",")]
        # 空でないURLのみ
        return [u for u in urls if u]
    
    def _load_imap_settings(self):
        """IMAP設定を読み込む"""
        settings_file = APP_DIR / "settings" / "fetch_settings.json"
        
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
        """クッキー保存ディレクトリを取得"""
        # グローバル変数APP_DIRを確認（GUI.pyから渡される）
        global APP_DIR
        if 'APP_DIR' in globals() and APP_DIR:
            cookies_dir = APP_DIR / "_internal" / "cookies" / "Amazon"
            cookies_dir.mkdir(parents=True, exist_ok=True)
            return cookies_dir
        
        # フォールバック
        cookies_dir = get_app_dir() / "_internal" / "cookies" / "Amazon"
        cookies_dir.mkdir(parents=True, exist_ok=True)
        return cookies_dir
    
    def _get_cookie_file(self):
        """クッキーファイルパスを取得"""
        return self.cookies_dir / f"{self.profile}_cookies.json"
    
    def _parse_proxy(self):
        """プロキシ設定をパース"""
        if not self.proxy:
            return None
        
        proxy = self.proxy.strip()
        
        # 形式: host:port:user:pass または host:port
        if proxy.count(":") >= 3:
            parts = proxy.split(":")
            host = parts[0]
            port = parts[1]
            user = parts[2]
            password = ":".join(parts[3:])
            
            return {
                "server": f"http://{host}:{port}",
                "username": user,
                "password": password
            }
        elif ":" in proxy:
            return {"server": f"http://{proxy}"}
        
        return None
    
    def _find_chrome_path(self):
        """システムにインストールされているChromeのパスを探す"""
        import platform
        
        if platform.system() == "Windows":
            # Windowsの一般的なChromeのパス
            possible_paths = [
                os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
            ]
        elif platform.system() == "Darwin":
            # macOS
            possible_paths = [
                "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            ]
        else:
            # Linux
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
        self.context.set_default_timeout(120000)  # 2分
        self.context.set_default_navigation_timeout(180000)  # 3分
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
        
        element.fill("")
        time.sleep(random.uniform(0.1, 0.3))
        
        for char in text:
            element.type(char, delay=random.randint(min_delay, max_delay))
            if random.random() < 0.1:
                time.sleep(random.uniform(0.1, 0.3))
        
        time.sleep(random.uniform(0.2, 0.5))
    
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
            
            nav_link = self.page.query_selector("#nav-link-accountList")
            if nav_link:
                text = nav_link.inner_text()
                if "ログイン" in text or "サインイン" in text or "Sign in" in text:
                    return False
                return True
            return False
        except:
            return False
    
    def check_login_with_navigation(self):
        """注文履歴ページにアクセスしてログイン状態を確認"""
        for attempt in range(self.max_retries + 1):
            try:
                self.page.goto(self.ORDER_HISTORY_URL, wait_until="domcontentloaded", timeout=120000)
                self.random_sleep(2, 3)
                return self.is_logged_in()
            except Exception as e:
                if attempt < self.max_retries:
                    print(f"Retry {attempt + 1}/{self.max_retries}...")
                    time.sleep(3)
                else:
                    print(f"Failed to check login status after {self.max_retries} retries")
                    return None
        return None
    
    def check_login_on_page(self):
        """現在のページでログイン状態を確認（ページ遷移なし）"""
        try:
            current_url = self.page.url
            
            # ログインページにリダイレクトされた場合
            if "/ap/signin" in current_url or "/ap/register" in current_url:
                print("Redirected to login page")
                return False
            
            return self.is_logged_in()
        except:
            return False
    
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
                        return self.do_login()  # 再帰呼び出し
                    elif login_type == "password_only":
                        return self.do_login()  # 再帰呼び出し
                
                # メールアドレス入力
                email_selector = self._wait_for_email_field(max_wait=60)
                if not email_selector:
                    # パスワードフィールドが先に現れた場合
                    if self.page.query_selector("#ap_password"):
                        return self.do_login()  # 再帰呼び出し
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
                    otp = self._fetch_otp_from_email()
                    if otp:
                        print("Entering OTP...")
                        self.human_type("#auth-mfa-otpcode", otp)
                        self.random_sleep(0.5, 1)
                        self.human_click("#auth-signin-button")
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
            # a-fixed-left-grid内のアカウント情報
            first_account = self.page.query_selector(".cvf-account-switcher-profile-details-after-account-removed")
            if first_account:
                # 親のリンクを探す
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
                    # 「アカウントの追加」や「ログアウト」以外のリンク
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
            return otp_input is not None
        except:
            return False
    
    def _click_skip_button(self):
        """「後で」ボタンがあればクリック"""
        try:
            # 様々な「後で」「スキップ」ボタンのセレクタ
            skip_selectors = [
                "#ap-account-fixup-phone-skip-link",  # 電話番号登録スキップ
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
    
    def _fetch_otp_from_email(self, max_wait=120):
        """メールからOTPを取得"""
        if not self.imap_settings:
            print("No IMAP settings configured")
            return None
        
        server = self.imap_settings.get("server", "")
        port = int(self.imap_settings.get("port", 993))
        username = self.imap_settings.get("username", "")
        password = self.imap_settings.get("password", "")
        
        if not all([server, username, password]):
            print("Incomplete IMAP settings")
            return None
        
        print(f"Fetching OTP from email: {username}")
        
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                mail = imaplib.IMAP4_SSL(server, port)
                mail.login(username, password)
                mail.select("INBOX")
                
                _, messages = mail.search(None, "UNSEEN")
                message_ids = messages[0].split()
                
                message_ids = sorted(message_ids, key=lambda x: int(x), reverse=True)[:20]
                
                for msg_id in message_ids:
                    _, msg_data = mail.fetch(msg_id, "(RFC822)")
                    email_body = msg_data[0][1]
                    msg = email.message_from_bytes(email_body)
                    
                    from_header = msg.get("From", "")
                    if "amazon" not in from_header.lower():
                        continue
                    
                    body = self._get_email_body(msg)
                    
                    otp_match = re.search(r'\b(\d{6})\b', body)
                    if otp_match:
                        otp = otp_match.group(1)
                        print(f"Found OTP: {otp}")
                        mail.logout()
                        return otp
                
                mail.logout()
                
            except Exception as e:
                print(f"Email check error: {e}")
            
            print("OTP not found, waiting...")
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
                    try:
                        charset = part.get_content_charset() or 'utf-8'
                        body = part.get_payload(decode=True).decode(charset, errors='replace')
                        break
                    except:
                        pass
                elif content_type == "text/html":
                    try:
                        charset = part.get_content_charset() or 'utf-8'
                        body = part.get_payload(decode=True).decode(charset, errors='replace')
                    except:
                        pass
        else:
            try:
                charset = msg.get_content_charset() or 'utf-8'
                body = msg.get_payload(decode=True).decode(charset, errors='replace')
            except:
                pass
        
        return body
    
    # ========== リクエストベース抽選処理 ==========
    
    def _get_cookies_for_requests(self):
        """Playwrightのクッキーをrequestsセッション用に変換"""
        cookies = {}
        try:
            pw_cookies = self.context.cookies()
            for cookie in pw_cookies:
                cookies[cookie['name']] = cookie['value']
        except:
            pass
        return cookies
    
    def _format_proxy_for_requests(self, proxy_str):
        """プロキシ文字列をrequests用の形式に変換
        
        対応形式:
        - http://host:port
        - http://user:pass@host:port
        - host:port:user:pass (Donut形式)
        - host:port
        """
        try:
            if not proxy_str:
                return None
            
            # 既にhttp://で始まっている場合
            if proxy_str.startswith('http://') or proxy_str.startswith('https://'):
                return proxy_str
            
            parts = proxy_str.split(':')
            
            if len(parts) == 2:
                # host:port
                host, port = parts
                return f'http://{host}:{port}'
            elif len(parts) == 4:
                # host:port:user:pass (Donut形式)
                host, port, user, passwd = parts
                return f'http://{user}:{passwd}@{host}:{port}'
            elif len(parts) > 4:
                # host:port:user:pass (passにコロンが含まれる場合)
                host = parts[0]
                port = parts[1]
                user = parts[2]
                passwd = ':'.join(parts[3:])  # 残りは全てパスワード
                return f'http://{user}:{passwd}@{host}:{port}'
            else:
                # その他の形式はそのまま
                return f'http://{proxy_str}'
                
        except Exception as e:
            print(f"Failed to parse proxy: {proxy_str}")
            return None
    
    def _extract_raffle_info(self, html_content):
        """HTMLから抽選に必要な情報を抽出"""
        try:
            soup = BeautifulSoup(html_content, 'html.parser')
            
            csrf_token = None
            endpoint = None
            signed_in = False
            
            # パターン1: hdp-request-* (サイドバー用)
            csrf_elem = soup.find('input', {'id': 'hdp-request-csrf-token'})
            if csrf_elem:
                csrf_token = csrf_elem.get('value')
                endpoint_elem = soup.find('input', {'id': 'hdp-request-ajax-endpoint'})
                endpoint = endpoint_elem.get('value') if endpoint_elem else None
                signed_in_elem = soup.find('input', {'id': 'hdp-request-signedIn'})
                signed_in = signed_in_elem.get('value') == 'true' if signed_in_elem else False
            
            # パターン2: hdp-ib-* (メインページ用)
            if not csrf_token:
                csrf_elem = soup.find('input', {'id': 'hdp-ib-csrf-token'})
                if csrf_elem:
                    csrf_token = csrf_elem.get('value')
                    endpoint_elem = soup.find('input', {'id': 'hdp-ib-ajax-endpoint'})
                    endpoint = endpoint_elem.get('value') if endpoint_elem else None
                    signed_in_elem = soup.find('input', {'id': 'hdp-ib-signedIn'})
                    signed_in = signed_in_elem.get('value') == 'true' if signed_in_elem else False
            
            # エンドポイントが抽選用かどうか検証（request-inviteを含む必要がある）
            if endpoint and 'request-invite' not in endpoint:
                endpoint = None  # 抽選用ではない
            
            # 既にリクエスト済みかチェック
            already_requested = False
            
            # メインページ用: hdp-detail-requested-idが存在し、かつaok-hiddenクラスを持っていない場合
            requested_elem = soup.find('div', {'id': 'hdp-detail-requested-id'})
            if requested_elem:
                classes = requested_elem.get('class', [])
                if 'aok-hidden' not in classes:
                    already_requested = True
            
            # サイドバー用: aod-hdp-offer-request-invited-* がaok-hiddenを持っていない場合
            if not already_requested:
                for i in range(6):
                    requested_elem = soup.find('div', {'id': f'aod-hdp-offer-request-invited-{i}'})
                    if requested_elem:
                        classes = requested_elem.get('class', [])
                        # aok-hiddenがなければ参加済み（aok-inline-blockがある）
                        if 'aok-hidden' not in classes:
                            already_requested = True
                            break
            
            # 商品タイトル
            title_elem = soup.find('span', {'id': 'productTitle'})
            title = title_elem.get_text().strip() if title_elem else ""
            
            # 商品画像URL
            image_url = ""
            # パターン1: landingImage (メイン商品画像)
            img_elem = soup.find('img', {'id': 'landingImage'})
            if img_elem:
                image_url = img_elem.get('src', '') or img_elem.get('data-old-hires', '')
            # パターン2: imgBlkFront
            if not image_url:
                img_elem = soup.find('img', {'id': 'imgBlkFront'})
                if img_elem:
                    image_url = img_elem.get('src', '')
            # パターン3: data-a-dynamic-imageから取得
            if not image_url:
                img_elem = soup.find('img', {'data-a-dynamic-image': True})
                if img_elem:
                    dynamic_data = img_elem.get('data-a-dynamic-image', '')
                    if dynamic_data:
                        try:
                            import json as json_mod
                            urls = json_mod.loads(dynamic_data)
                            if urls:
                                image_url = list(urls.keys())[0]
                        except:
                            pass
            
            return {
                'csrf_token': csrf_token,
                'endpoint': endpoint,
                'signed_in': signed_in,
                'already_requested': already_requested,
                'title': title,
                'image_url': image_url
            }
        except Exception as e:
            print(f"Error extracting raffle info: {e}")
            return None
    
    def _send_raffle_request(self, session, endpoint, csrf_token):
        """リクエストベースで抽選に参加"""
        try:
            url = f"https://{endpoint}"
            
            headers = {
                'Accept': 'application/vnd.com.amazon.api+json; type="aapi.highdemandproductcontracts.request-invite/v1"',
                'Accept-Language': 'ja-JP',
                'Content-Type': 'application/vnd.com.amazon.api+json; type="aapi.highdemandproductcontracts.request-invite.request/v1"',
                'Origin': 'https://www.amazon.co.jp',
                'Referer': 'https://www.amazon.co.jp/',
                'X-Api-Csrf-Token': csrf_token,
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36'
            }
            
            response = session.post(url, headers=headers, json={}, timeout=30)
            
            if response.status_code == 200:
                return True, "Success"
            elif response.status_code == 409:
                # 既にリクエスト済み
                return False, "Already Raffled"
            else:
                print(f"Raffle request failed: {response.status_code}")
                return False, f"Error {response.status_code}"
                
        except Exception as e:
            print(f"Error sending raffle request: {e}")
            return False, "Request Failed"
    
    def _extract_asin_from_url(self, url):
        """URLからASINを抽出"""
        # パターン1: /dp/ASIN
        match = re.search(r'/dp/([A-Z0-9]{10})', url)
        if match:
            return match.group(1)
        
        # パターン2: /gp/product/ASIN
        match = re.search(r'/gp/product/([A-Z0-9]{10})', url)
        if match:
            return match.group(1)
        
        # パターン3: asin=ASIN
        match = re.search(r'asin=([A-Z0-9]{10})', url)
        if match:
            return match.group(1)
        
        return None
    
    def _fetch_sidebar_html(self, session, asin):
        """サイドバーのHTMLを取得"""
        try:
            sidebar_url = f"https://www.amazon.co.jp/gp/product/ajax/aodAjaxMain/ref=dp_aod_ALL_mbc?asin={asin}&m=&qid=&smid=&sourcecustomerorglistid=&sourcecustomerorglistitemid=&sr=&pc=dp"
            
            headers = {
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'ja-JP,ja;q=0.9',
                'Referer': f'https://www.amazon.co.jp/dp/{asin}',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36'
            }
            
            response = session.get(sidebar_url, headers=headers, timeout=30)
            
            if response.status_code == 200:
                return response.text
            else:
                print(f"Failed to fetch sidebar: {response.status_code}")
                return None
                
        except Exception as e:
            print(f"Error fetching sidebar: {e}")
            return None
    
    def process_raffle_request_based(self, url):
        """リクエストベースで1つのURLの抽選処理（データ転送量削減）"""
        result = {
            "url": url,
            "status": "Unknown",
            "title": "",
            "image_url": ""
        }
        
        try:
            # PlaywrightのクッキーでrequestsセッションをCookieを設定
            cookies = self._get_cookies_for_requests()
            
            session = requests.Session()
            session.cookies.update(cookies)
            
            # プロキシ設定
            if self.proxy:
                proxy_url = self._format_proxy_for_requests(self.proxy)
                if proxy_url:
                    session.proxies = {
                        'http': proxy_url,
                        'https': proxy_url
                    }
            
            # HTMLを取得（画像・CSS・JSは読み込まない）
            headers = {
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
                'Accept-Language': 'ja-JP,ja;q=0.9',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36'
            }
            
            response = session.get(url, headers=headers, timeout=30)
            
            if response.status_code != 200:
                print(f"Connection failed: {response.status_code}")
                result["status"] = "Connection Failed"
                return result
            
            # 抽選情報を抽出（メインページ）
            raffle_info = self._extract_raffle_info(response.text)
            
            if not raffle_info:
                print("Parse error")
                result["status"] = "Parse Error"
                return result
            
            result["title"] = raffle_info['title']
            result["image_url"] = raffle_info.get('image_url', '')
            if result["title"]:
                print(f"Product: {result['title'][:50]}...")
            
            # 既にリクエスト済みかチェック
            if raffle_info['already_requested']:
                print("Already raffled!")
                result["status"] = "Already Raffled"
                return result
            
            # メインページでCSRFトークンが見つからない場合、サイドバーを取得
            if not raffle_info['csrf_token'] or not raffle_info['endpoint']:
                asin = self._extract_asin_from_url(url)
                if asin:
                    sidebar_html = self._fetch_sidebar_html(session, asin)
                    if sidebar_html:
                        sidebar_info = self._extract_raffle_info(sidebar_html)
                        if sidebar_info:
                            # サイドバーで既にリクエスト済みかチェック
                            if sidebar_info['already_requested']:
                                print("Already raffled!")
                                result["status"] = "Already Raffled"
                                return result
                            
                            # CSRFトークンとエンドポイントを取得
                            if sidebar_info['csrf_token'] and sidebar_info['endpoint']:
                                raffle_info['csrf_token'] = sidebar_info['csrf_token']
                                raffle_info['endpoint'] = sidebar_info['endpoint']
            
            # CSRFトークンとエンドポイントがない場合
            if not raffle_info['csrf_token'] or not raffle_info['endpoint']:
                print("Invite button not found")
                result["status"] = "Not Find"
                return result
            
            # 抽選リクエストを送信（ログイン状態はAPIレスポンスで判断）
            success, status = self._send_raffle_request(
                session, 
                raffle_info['endpoint'], 
                raffle_info['csrf_token']
            )
            
            if success:
                print("Raffle entry successful!")
                result["status"] = "Success"
            else:
                if status == "Already Raffled":
                    print("Already raffled!")
                else:
                    print(f"Error: {status}")
                result["status"] = status
            
            return result
            
        except Exception as e:
            print(f"Error: {e}")
            result["status"] = "Error"
            return result
    
    # ========== ブラウザベース抽選処理 ==========
    
    def process_raffle(self, url):
        """1つのURLの抽選処理"""
        result = {
            "url": url,
            "status": "Unknown",
            "title": "",
            "image_url": ""
        }
        
        try:
            # Resi対応: リトライしながらURLにアクセス
            for attempt in range(self.max_retries + 1):
                try:
                    self.page.goto(url, wait_until="domcontentloaded", timeout=120000)
                    self.random_sleep(3, 5)
                    break
                except Exception as e:
                    if attempt < self.max_retries:
                        print(f"Retry {attempt + 1}/{self.max_retries}...")
                        time.sleep(3)
                    else:
                        print(f"Failed to access URL after {self.max_retries} retries")
                        result["status"] = "Connection Failed"
                        return result
            
            # ログインページにリダイレクトされたかチェック
            current_url = self.page.url
            if "/ap/signin" in current_url or "/ap/register" in current_url:
                print("Redirected to login page, logging in...")
                login_result = self.do_login()
                if login_result == "locked":
                    result["status"] = "Locked"
                    return result
                elif not login_result:
                    print("Login failed")
                    result["status"] = "Login Failed"
                    return result
                
                # ログイン後、元のURLに再アクセス
                for attempt in range(self.max_retries + 1):
                    try:
                        self.page.goto(url, wait_until="domcontentloaded", timeout=120000)
                        self.random_sleep(3, 5)
                        break
                    except Exception as e:
                        if attempt < self.max_retries:
                            print(f"Retry {attempt + 1}/{self.max_retries}...")
                            time.sleep(3)
                        else:
                            result["status"] = "Connection Failed"
                            return result
            
            # 商品タイトルを取得
            title_elem = self.page.query_selector("#productTitle")
            if title_elem:
                result["title"] = title_elem.inner_text().strip()
                print(f"Product: {result['title'][:50]}...")
            
            # 商品画像URLを取得
            try:
                img_elem = self.page.query_selector("#landingImage")
                if img_elem:
                    result["image_url"] = img_elem.get_attribute("src") or img_elem.get_attribute("data-old-hires") or ""
                if not result["image_url"]:
                    img_elem = self.page.query_selector("#imgBlkFront")
                    if img_elem:
                        result["image_url"] = img_elem.get_attribute("src") or ""
            except:
                pass
            
            # 招待ボタンのセレクタ（メインページ用）
            main_invite_selectors = [
                "#hdp-invite-button",
                "#hdp-invite-button input.a-button-input",
                "#invite-button input.a-button-input",
                "#invite-button-announce",
                "input[name='submit.invite-button']",
            ]
            
            # 招待ボタンのセレクタ（サイドバー用）
            sidebar_invite_selectors = [
                "#hdp-request-invitation-0",
                "#hdp-request-invitation-1",
                "#hdp-request-invitation-2",
                "#hdp-request-invitation-3",
                "#hdp-request-invitation-4",
                "#hdp-request-invitation-5",
                "#hdp-request-invitation-0 input.a-button-input",
                "#hdp-request-invitation-1 input.a-button-input",
                "#hdp-request-invitation-2 input.a-button-input",
                "#hdp-request-invitation-3 input.a-button-input",
                "#hdp-request-invitation-4 input.a-button-input",
                "#hdp-request-invitation-5 input.a-button-input",
                "input[name='submit.inviteButton']",
            ]
            
            # リクエスト済みのセレクタ（メインページ用）
            main_already_selectors = [
                "#hdp-detail-requested-id",
            ]
            
            # リクエスト済みのセレクタ（サイドバー用）- aok-hiddenクラスがないもののみ
            sidebar_already_selectors = [
                "#aod-hdp-offer-request-invited-0:not(.aok-hidden)",
                "#aod-hdp-offer-request-invited-1:not(.aok-hidden)",
                "#aod-hdp-offer-request-invited-2:not(.aok-hidden)",
                "#aod-hdp-offer-request-invited-3:not(.aok-hidden)",
                "#aod-hdp-offer-request-invited-4:not(.aok-hidden)",
                "#aod-hdp-offer-request-invited-5:not(.aok-hidden)",
            ]
            
            # ステップ1: メインページでボタンを探す
            invite_btn = None
            found_selector = None
            is_main_page_button = False
            
            for selector in main_invite_selectors:
                try:
                    invite_btn = self.page.query_selector(selector)
                    if invite_btn:
                        found_selector = selector
                        is_main_page_button = True
                        break
                except:
                    continue
            
            # メインページにボタンがある場合 → エントリー処理へ
            if invite_btn:
                pass  # 後でクリック処理
            else:
                # メインページでリクエスト済みIDがあるかチェック
                for selector in main_already_selectors:
                    try:
                        already_elem = self.page.query_selector(selector)
                        if already_elem:
                            print("Already raffled!")
                            result["status"] = "Already Raffled"
                            return result
                    except:
                        continue
                
                # メインページにボタンも完了IDもない → サイドバーを開く
                print("Opening sidebar...")
                sidebar_trigger = self.page.query_selector("#aod-ingress-link")
                
                if sidebar_trigger:
                    try:
                        sidebar_trigger.click()
                        self.random_sleep(3, 5)
                    except:
                        pass
                
                # サイドバー内でリクエスト済みIDがあるかチェック
                for selector in sidebar_already_selectors:
                    try:
                        already_elem = self.page.query_selector(selector)
                        if already_elem:
                            print("Already raffled!")
                            result["status"] = "Already Raffled"
                            return result
                    except:
                        continue
                
                # サイドバー内でボタンを探す（最大15秒）
                for attempt in range(15):
                    # まずリクエスト済みIDをチェック
                    for selector in sidebar_already_selectors:
                        try:
                            already_elem = self.page.query_selector(selector)
                            if already_elem:
                                print("Already raffled!")
                                result["status"] = "Already Raffled"
                                return result
                        except:
                            continue
                    
                    # ボタンを探す
                    for selector in sidebar_invite_selectors:
                        try:
                            invite_btn = self.page.query_selector(selector)
                            if invite_btn:
                                found_selector = selector
                                is_main_page_button = False
                                break
                        except:
                            continue
                    
                    if invite_btn:
                        break
                    
                    if attempt < 14:
                        time.sleep(1)
                
                # サイドバーにボタンがない場合
                if not invite_btn:
                    print("Invite button not found")
                    result["status"] = "Not Find"
                    return result
            
            # ボタンをクリック（複数の方法を試す）
            self.random_sleep(1, 2)
            
            click_success = False
            
            # 方法1: スクロールしてから直接クリック
            try:
                invite_btn.scroll_into_view_if_needed()
                self.random_sleep(0.3, 0.5)
                invite_btn.click()
                print("Clicked invite button")
                click_success = True
            except:
                pass
            
            # 方法2: JavaScriptでスクロール＆クリック
            if not click_success:
                try:
                    self.page.evaluate("""
                        (element) => {
                            element.scrollIntoView({behavior: 'smooth', block: 'center'});
                            setTimeout(() => element.click(), 300);
                        }
                    """, invite_btn)
                    self.random_sleep(0.5, 1)
                    print("Clicked invite button")
                    click_success = True
                except:
                    pass
            
            # 方法3: dispatchEventでクリック
            if not click_success:
                try:
                    self.page.evaluate("""
                        (element) => {
                            element.dispatchEvent(new MouseEvent('click', {
                                bubbles: true,
                                cancelable: true,
                                view: window
                            }));
                        }
                    """, invite_btn)
                    print("Clicked invite button")
                    click_success = True
                except:
                    pass
            
            # 方法4: フォーム送信
            if not click_success and found_selector:
                try:
                    self.page.evaluate("""
                        (selector) => {
                            var btn = document.querySelector(selector);
                            if (btn && btn.form) {
                                btn.form.submit();
                                return true;
                            }
                            return false;
                        }
                    """, found_selector)
                    print("Submitted form")
                    click_success = True
                except:
                    pass
            
            # 方法5: セレクタでクリック
            if not click_success and found_selector:
                try:
                    self.page.click(found_selector)
                    print("Clicked invite button")
                    click_success = True
                except:
                    pass
            
            if not click_success:
                print("All click methods failed")
                result["status"] = "Click Failed"
                return result
            
            # ステップ4: 成功確認（ID確認）
            self.random_sleep(3, 5)
            
            # メインページかサイドバーかで期待するIDが異なる
            if is_main_page_button:
                success_selectors = main_already_selectors
            else:
                success_selectors = sidebar_already_selectors
            
            # 最大10秒待機（成功IDが表示されるまで）
            for attempt in range(10):
                # 成功IDが表示されているか確認
                for selector in success_selectors:
                    try:
                        success_elem = self.page.query_selector(selector)
                        if success_elem:
                            print("Raffle success!")
                            result["status"] = "Success"
                            return result
                    except:
                        continue
                
                time.sleep(1)
            
            # 10秒待っても成功IDが表示されなかった場合 → リロードしてリトライ
            print("Retrying with reload...")
            
            for reload_attempt in range(3):
                try:
                    # ページをリロード
                    self.page.reload(wait_until="domcontentloaded", timeout=120000)
                    self.random_sleep(3, 5)
                    
                    # メインページの場合
                    if is_main_page_button:
                        # 成功IDがあるか確認
                        for selector in main_already_selectors:
                            try:
                                success_elem = self.page.query_selector(selector)
                                if success_elem:
                                    print("Raffle success!")
                                    result["status"] = "Success"
                                    return result
                            except:
                                continue
                        
                        # ボタンを探してクリック
                        for selector in main_invite_selectors:
                            try:
                                btn = self.page.query_selector(selector)
                                if btn:
                                    btn.scroll_into_view_if_needed()
                                    self.random_sleep(0.3, 0.5)
                                    btn.click()
                                    print("Clicked invite button")
                                    break
                            except:
                                continue
                    else:
                        # サイドバーの場合
                        # 成功IDがあるか確認
                        for selector in sidebar_already_selectors:
                            try:
                                success_elem = self.page.query_selector(selector)
                                if success_elem:
                                    print("Raffle success!")
                                    result["status"] = "Success"
                                    return result
                            except:
                                continue
                        
                        # サイドバーを開く
                        sidebar_trigger = self.page.query_selector("#aod-ingress-link")
                        if sidebar_trigger:
                            try:
                                sidebar_trigger.click()
                                self.random_sleep(3, 5)
                            except:
                                pass
                        
                        # 成功IDがあるか再確認
                        for selector in sidebar_already_selectors:
                            try:
                                success_elem = self.page.query_selector(selector)
                                if success_elem:
                                    print("Raffle success!")
                                    result["status"] = "Success"
                                    return result
                            except:
                                continue
                        
                        # ボタンを探してクリック
                        for selector in sidebar_invite_selectors:
                            try:
                                btn = self.page.query_selector(selector)
                                if btn:
                                    btn.scroll_into_view_if_needed()
                                    self.random_sleep(0.3, 0.5)
                                    btn.click()
                                    print("Clicked invite button")
                                    break
                            except:
                                continue
                    
                    # クリック後、成功確認（10秒）
                    self.random_sleep(3, 5)
                    for attempt in range(10):
                        for selector in success_selectors:
                            try:
                                success_elem = self.page.query_selector(selector)
                                if success_elem:
                                    print("Raffle success!")
                                    result["status"] = "Success"
                                    return result
                            except:
                                continue
                        time.sleep(1)
                    
                except Exception as e:
                    pass
            
            # リトライしても成功しなかった場合
            print("Could not confirm success after retries")
            result["status"] = "Timeout"
            return result
            
        except Exception as e:
            print(f"Raffle error: {e}")
            import traceback
            traceback.print_exc()
            result["status"] = "Error"
            return result
    
    # ========== メイン処理 ==========
    
    def run(self):
        """メイン実行
        
        Returns:
            tuple: (success: bool, results: list, error_status: str or None)
            error_status: エラー時のステータス文字列（成功時はNone）
        """
        try:
            print("=" * 50)
            print(f"Site: {self.site} Mode: {self.mode}")
            print(f"URLs: {len(self.urls)}")
            print("=" * 50)
            
            if not self.urls:
                print("No URLs provided")
                return False, [], "Failed No URLs"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, [], "Stopped"
            
            # ブラウザ起動
            print("\x00STATUS:Starting Task", flush=True)
            try:
                self.start_browser(headless=self.headless)
            except Exception as e:
                print(f"Browser start failed: {e}")
                return False, [], "Failed Browser"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, [], "Stopped"
            
            # クッキー読み込み
            self.load_cookies()
            
            # ログイン確認
            print("\x00STATUS:Checking Login", flush=True)
            print("Step 1: Checking login status...")
            try:
                login_status = self.check_login_with_navigation()
            except Exception as e:
                print(f"Login check failed: {e}")
                return False, [], "Failed Login Check"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, [], "Stopped"
            
            if login_status is None:
                # 接続失敗
                print("Failed to connect")
                return False, [], "Failed Connection"
            elif not login_status:
                print("\x00STATUS:Logging In", flush=True)
                print("Step 2: Logging in...")
                try:
                    login_result = self.do_login()
                    if login_result == "locked":
                        return False, [], "Failed Locked"
                    elif not login_result:
                        print("Login failed")
                        return False, [], "Failed Login"
                except Exception as e:
                    print(f"Login error: {e}")
                    return False, [], "Failed Login"
            else:
                print("Already logged in")
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, [], "Stopped"
            
            # 各URLを処理
            success_count = 0
            
            try:
                for i, url in enumerate(self.urls, 1):
                    # ★ ストップチェック
                    if self._stop_requested:
                        print("Task stopped by user")
                        return False, self.results, "Stopped"
                    
                    print(f"STATUS:Raffle {i}/{len(self.urls)}", flush=True)
                    print(f"--------------------------------------------------")
                    print(f"Processing URL {i}/{len(self.urls)}")
                    
                    # 全てリクエストベースで処理（データ転送量削減）
                    result = self.process_raffle_request_based(url)
                    
                    # ★ ストップチェック
                    if self._stop_requested:
                        print("Task stopped by user")
                        return False, self.results, "Stopped"
                    
                    # リクエストベースで失敗した場合、ブラウザベースにフォールバック
                    # Not Findはフォールバックしない（抽選が開始していない）
                    if result["status"] in ["Parse Error", "Connection Failed", "Request Failed", "Error"]:
                        print(f"Fallback to browser ({result['status']})")
                        result = self.process_raffle(url)
                    
                    # ★ ストップチェック
                    if self._stop_requested:
                        print("Task stopped by user")
                        return False, self.results, "Stopped"
                    
                    print(f"Result: {result['status']}")
                    self.results.append(result)
                    
                    if result["status"] == "Success":
                        success_count += 1
                    
                    # URL間の待機
                    if i < len(self.urls):
                        self.random_sleep(0.5, 1.5)
            except Exception as e:
                print(f"Error: {e}")
                return False, self.results, "Failed Raffles"
            
            print(f"--------------------------------------------------")
            # 結果サマリー
            print(f"Total: {len(self.urls)}, Success: {success_count}")
            
            # クッキー保存
            print("\x00STATUS:Cookies Saved", flush=True)
            self.save_cookies()
            
            return success_count > 0, self.results, None
            
        except Exception as e:
            if self._stop_requested:
                print("Task stopped by user")
                return False, [], "Stopped"
            print(f"Error: {e}")
            import traceback
            traceback.print_exc()
            return False, [], "Failed Unknown"
        finally:
            self.close_browser()


def main():
    parser = argparse.ArgumentParser(description="Amazon Raffle Bot")
    parser.add_argument("--task-data", help="Path to task data JSON file")
    
    args = parser.parse_args()
    
    if args.task_data:
        with open(args.task_data, 'r', encoding='utf-8') as f:
            task_data = json.load(f)
    else:
        task_data = {
            "Profile": "test",
            "Site": "Amazon",
            "Mode": "Raffle",
            "URL": "",
            "Proxy": "",
            "Loginid": "",
            "Loginpass": ""
        }
    
    bot = AmazonRaffle(task_data)
    success, results, error_status = bot.run()
    
    # エラーステータスがある場合は出力（GUIが読み取る）
    if error_status:
        print(f"RAFFLE_ERROR:{error_status}")
    
    # 結果をJSONで出力（GUIが読み取れるように）
    print(f"\nRAFFLE_RESULTS:{json.dumps(results, ensure_ascii=False)}")
    
    # 正常終了（GUIはRAFFLE_ERROR/RAFFLE_RESULTSで判断する）


if __name__ == "__main__":
    main()