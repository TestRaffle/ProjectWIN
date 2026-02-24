"""
Amazon Browser モード
- クッキーが保存されていればそれを読み込んでログイン状態を復元
- ログインされていなければログイン処理を実行
- 確認コード（OTP）はメールから自動取得
- ログイン成功後、指定URLにアクセス
- ユーザーが自由に操作可能（Stopボタンが押されるまで待機）
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


class AmazonBrowser:
    """Amazon Browserモードクラス"""
    
    # Amazonのベース情報
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
        
        # よく使うデータを取り出し
        self.profile = task_data.get("Profile", "")
        self.site = task_data.get("Site", "")
        self.mode = task_data.get("Mode", "")
        self.url = task_data.get("URL", "") or self.BASE_URL
        self.proxy = task_data.get("Proxy", "")
        self.headless = task_data.get("Headless", False)
        
        # ログイン情報
        self.login_id = task_data.get("Loginid", "")
        self.login_pass = task_data.get("Loginpass", "")
        
        # 設定ディレクトリを取得（exe化対応）
        self._settings_dir = self._get_settings_dir()
        
        # IMAP設定（設定ファイルから読み込み）
        self.imap_settings = self._load_imap_settings()
        
        # general_settings読み込み（リトライ回数）
        self.general_settings = self._load_general_settings()
        self.max_retries = self.general_settings.get("retry_count", 3)
        
        # ブラウザ関連
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        
        # クッキー保存先ディレクトリ
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
        """general_settings.jsonを読み込む"""
        try:
            settings_path = self._settings_dir / "general_settings.json"
            if settings_path.exists():
                with open(settings_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Failed to load general settings: {e}")
        return {}
    
    def _load_imap_settings(self):
        """IMAP設定を読み込む（accountsリストから選択されたアカウントを取得）"""
        # 設定ファイルのパス
        settings_file = self._settings_dir / "fetch_settings.json"
        
        if settings_file.exists():
            try:
                with open(settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # accountsリストから選択されたアカウントを取得
                accounts = data.get("accounts", [])
                for acc in accounts:
                    if acc.get("selected"):
                        return acc
                # 選択されていなければ最初のアカウント
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
    
    def start_browser(self):
        """ブラウザを起動（通常のコンテキスト + JSONクッキー）"""
        self.playwright = sync_playwright().start()
        
        proxy_config = self._parse_proxy()
        
        launch_options = {
            "headless": self.headless,
            "args": [
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ]
        }
        
        # システムのChromeを探す
        chrome_path = self._find_chrome_path()
        if chrome_path:
            launch_options["executable_path"] = chrome_path
            print(f"Using system Chrome: {chrome_path}")
        else:
            print("System Chrome not found, trying channel='chrome'")
            launch_options["channel"] = "chrome"
        
        if self.headless:
            print("Running in headless mode")
        
        # ブラウザを起動
        try:
            self.browser = self.playwright.chromium.launch(**launch_options)
        except Exception as e:
            print(f"Failed to launch with options, trying channel='chrome': {e}")
            launch_options.pop("executable_path", None)
            launch_options["channel"] = "chrome"
            self.browser = self.playwright.chromium.launch(**launch_options)
        
        # コンテキストオプション
        context_options = {
            "viewport": {"width": 1280, "height": 1280},
            "user_agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "ignore_https_errors": True,
        }
        
        if proxy_config:
            context_options["proxy"] = proxy_config
            print(f"Using proxy: {proxy_config['server']}")
        
        # コンテキストを作成
        self.context = self.browser.new_context(**context_options)
        
        # タイムアウト設定
        self.context.set_default_timeout(60000)
        self.context.set_default_navigation_timeout(90000)
        
        # ページを作成
        self.page = self.context.new_page()
        
        # ブラウザが閉じられた時の検知
        self._browser_closed = False
        self.page.on("close", self._on_page_closed)
        self.browser.on("disconnected", self._on_browser_disconnected)
        
        print(f"Browser started")
    
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
        except:
            pass
    
    # ========== 人間らしい入力 ==========
    
    def human_type(self, element, text, min_delay=50, max_delay=150):
        """人間らしいタイピングでテキストを入力"""
        element.click()
        time.sleep(random.uniform(0.2, 0.5))
        element.fill("")  # クリア
        time.sleep(random.uniform(0.1, 0.3))
        
        for char in text:
            element.type(char, delay=random.randint(min_delay, max_delay))
            # たまにちょっと長めの間を入れる
            if random.random() < 0.1:
                time.sleep(random.uniform(0.1, 0.3))
        
        time.sleep(random.uniform(0.3, 0.6))
    
    def human_click(self, element):
        """人間らしいクリック（少し待ってからクリック）"""
        time.sleep(random.uniform(0.3, 0.8))
        element.click()
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
    
    def save_cookies(self):
        """クッキーをファイルに保存"""
        try:
            cookies = self.context.cookies()
            cookie_file = self._get_cookie_file()
            with open(cookie_file, 'w', encoding='utf-8') as f:
                json.dump(cookies, f, ensure_ascii=False, indent=2)
            print(f"Cookies saved: {cookie_file}")
        except Exception as e:
            print(f"Failed to save cookies: {e}", file=sys.stderr)
    
    def load_cookies(self):
        """保存されたクッキーを読み込む"""
        try:
            cookie_file = self._get_cookie_file()
            if cookie_file.exists():
                with open(cookie_file, 'r', encoding='utf-8') as f:
                    cookies = json.load(f)
                self.context.add_cookies(cookies)
                print(f"Cookies loaded: {cookie_file}")
                return True
        except Exception as e:
            print(f"Failed to load cookies: {e}", file=sys.stderr)
        return False
    
    def fetch_otp_from_email(self, max_wait=120):
        """
        メールからAmazonの確認コード（OTP）を取得
        ログインIDに対応する未読メールのみを対象とする（最新のものを優先）
        
        Args:
            max_wait: 最大待機時間（秒）
        
        Returns:
            str: OTPコード（見つからない場合はNone）
        """
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
        
        # ログインID（Amazonアカウントのメールアドレス）を取得
        target_email = self.login_id.lower().strip()
        print("Fetching OTP...")
        print(f"Using IMAP: {email_address}")
        
        start_time = time.time()
        
        while time.time() - start_time < max_wait:
            try:
                mail = imaplib.IMAP4_SSL(imap_server, imap_port)
                mail.login(email_address, email_password)
                mail.select("INBOX")
                
                # Amazonからの未読メールのみを検索
                _, message_numbers = mail.search(None, '(UNSEEN FROM "amazon")')
                
                if message_numbers[0]:
                    msg_nums = message_numbers[0].split()
                    total_count = len(msg_nums)
                    
                    # 最新20件のみを対象にする（パフォーマンス対策）
                    msg_nums = msg_nums[-20:] if total_count > 20 else msg_nums
                    print(f"Found {total_count} unread Amazon emails, checking latest {len(msg_nums)}")
                    
                    # メールを日付付きで取得してソート
                    emails_with_date = []
                    
                    for msg_num in msg_nums:
                        try:
                            # ヘッダーのみ取得して日付を確認
                            _, header_data = mail.fetch(msg_num, "(BODY.PEEK[HEADER.FIELDS (DATE)])")
                            
                            if header_data and header_data[0]:
                                header_bytes = header_data[0][1] if isinstance(header_data[0], tuple) else b''
                                header_str = header_bytes.decode('utf-8', errors='ignore')
                                
                                # 日付を抽出
                                date_match = re.search(r'Date:\s*(.+)', header_str, re.IGNORECASE)
                                date_str = date_match.group(1).strip() if date_match else ""
                                
                                # 日付をパース
                                try:
                                    from email.utils import parsedate_to_datetime
                                    mail_date = parsedate_to_datetime(date_str)
                                except:
                                    mail_date = None
                                
                                emails_with_date.append((msg_num, mail_date))
                        except Exception as e:
                            print(f"Error getting header for {msg_num}: {e}")
                            emails_with_date.append((msg_num, None))
                    
                    # 日付で降順ソート（最新が先頭）、日付がNoneのものは最後
                    emails_with_date.sort(key=lambda x: (x[1] is None, x[1] if x[1] else ""), reverse=True)
                    
                    print(f"Checking {len(emails_with_date)} emails (sorted by date, newest first)")
                    
                    # 最新のものから順にチェック
                    for msg_num, mail_date in emails_with_date:
                        # 本文を取得
                        _, msg_data = mail.fetch(msg_num, "(BODY.PEEK[])")
                        
                        for response_part in msg_data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                
                                # メールの宛先を確認（To, Delivered-To, X-Original-To など）
                                to_addresses = []
                                
                                # Toヘッダー
                                to_header = msg.get("To", "")
                                if to_header:
                                    to_addresses.append(to_header.lower())
                                
                                # Delivered-Toヘッダー（転送時に元の宛先が入る）
                                delivered_to = msg.get("Delivered-To", "")
                                if delivered_to:
                                    to_addresses.append(delivered_to.lower())
                                
                                # X-Original-Toヘッダー
                                original_to = msg.get("X-Original-To", "")
                                if original_to:
                                    to_addresses.append(original_to.lower())
                                
                                # Envelope-Toヘッダー
                                envelope_to = msg.get("Envelope-To", "")
                                if envelope_to:
                                    to_addresses.append(envelope_to.lower())
                                
                                # メール本文を取得
                                body = self._get_email_body(msg)
                                
                                # 本文内にもメールアドレスが含まれているか確認
                                body_lower = body.lower()
                                
                                # 対象のログインIDがメールに関連しているか確認
                                is_target_email = False
                                
                                # ヘッダーで確認
                                for addr in to_addresses:
                                    if target_email in addr:
                                        is_target_email = True
                                        print(f"Found matching email in header: {addr}")
                                        break
                                
                                # 本文で確認（ヘッダーになかった場合）
                                if not is_target_email and target_email in body_lower:
                                    is_target_email = True
                                    print(f"Found matching email in body")
                                
                                if is_target_email:
                                    # OTPコードを抽出（6桁の数字）
                                    otp_match = re.search(r'\b(\d{6})\b', body)
                                    if otp_match:
                                        otp = otp_match.group(1)
                                        date_info = mail_date.strftime('%Y-%m-%d %H:%M:%S') if mail_date else 'unknown'
                                        print(f"Found OTP for {target_email}: {otp} (mail date: {date_info})")
                                        
                                        # このメールを既読にする
                                        mail.store(msg_num, '+FLAGS', '\\Seen')
                                        
                                        mail.logout()
                                        return otp
                                else:
                                    print(f"Skipping email - not for {target_email}")
                else:
                    print("No unread Amazon emails found")
                
                mail.logout()
                
            except Exception as e:
                print(f"Error fetching email: {e}")
                import traceback
                traceback.print_exc()
            
            print(f"OTP for {target_email} not found yet, waiting...")
            time.sleep(5)
        
        print(f"Failed to fetch OTP for {target_email} within timeout")
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
            account_element = self.page.query_selector("#nav-link-accountList")
            if account_element:
                text = account_element.inner_text()
                if "Sign in" in text or "ログイン" in text:
                    return False
                return True
            
            return False
            
        except Exception as e:
            return False
    
    def check_login_with_navigation(self):
        """注文履歴ページにアクセスしてログイン状態を確認"""
        for attempt in range(self.max_retries + 1):
            try:
                self.page.goto(self.ORDER_HISTORY_URL, wait_until="commit", timeout=90000)
                time.sleep(2)
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
            print("Login credentials not provided - manual login required")
            try:
                self.page.goto(self.LOGIN_URL, wait_until="commit", timeout=90000)
                print("Please login manually...")
                return self._wait_for_manual_login()
            except Exception as e:
                print(f"Failed to open login page: {e}", file=sys.stderr)
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
                
                self.random_sleep(3, 5)
                
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
                password_input = self._wait_for_password_field(max_wait=60)
                
                if password_input:
                    print("Entering password...")
                    self.random_sleep(1, 2)
                    self.human_type(password_input, self.login_pass)
                    self.random_sleep(0.5, 1)
                else:
                    print("Password field not found after account selection")
                    return False
                
                if not self._click_login_button():
                    return False
                
                self.random_sleep(3, 5)
                
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
            elif login_type == "password_only":
                password_input = self.page.query_selector("#ap_password")
                print("Entering password...")
                self.random_sleep(1, 2)
                self.human_type(password_input, self.login_pass)
                self.random_sleep(0.5, 1)
                
                self._check_remember_me()
                
                if not self._click_login_button():
                    return False
                
                self.random_sleep(3, 5)
                
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
            elif login_type == "email_password":
                current_url = self.page.url
                if "/ap/signin" not in current_url:
                    for attempt in range(self.max_retries + 1):
                        try:
                            self.page.goto(self.LOGIN_URL, wait_until="commit", timeout=90000)
                            self.random_sleep(2, 3)
                            break
                        except Exception as e:
                            if attempt < self.max_retries:
                                print(f"Retry {attempt + 1}/{self.max_retries}...")
                                time.sleep(3)
                            else:
                                print(f"Failed to access login page")
                                return False
                    
                    if self._check_account_locked():
                        print("Account is locked!")
                        return "locked"
                    
                    login_type = self._detect_login_type(max_wait=30)
                    if login_type == "account_switcher":
                        return self.do_login()
                    elif login_type == "password_only":
                        return self.do_login()
                
                email_selector = self._wait_for_email_field(max_wait=60)
                if not email_selector:
                    if self.page.query_selector("#ap_password"):
                        return self.do_login()
                    print("Email field not found")
                    return False
                
                email_input = self.page.query_selector(email_selector)
                print("Entering email...")
                self.random_sleep(1, 2)
                self.human_type(email_input, self.login_id)
                self.random_sleep(0.5, 1)
                
                continue_btn = self.page.query_selector("#continue")
                if continue_btn:
                    self.human_click(continue_btn)
                self.random_sleep(2, 3)
                
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
                
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
                self.human_type(password_input, self.login_pass)
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
                self.save_cookies()
                
                if self._check_account_locked():
                    print("Account is locked!")
                    return "locked"
            
            self._skip_phone_number_prompt()
            
            if self._check_account_locked():
                print("Account is locked!")
                return "locked"
            
            login_result = self._wait_for_login_success()
            if login_result:
                print("Login successful!")
                self.save_cookies()
            return login_result
            
        except PlaywrightTimeout as e:
            print(f"Login timeout: {e}", file=sys.stderr)
            return False
        except Exception as e:
            print(f"Login error: {e}", file=sys.stderr)
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
                self.human_click(login_btn)
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
        """電話番号追加の画面が表示された場合はスキップする"""
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
                        skip_btn.click()
                        self.random_sleep(1, 2)
                        print("Clicked skip button")
                        return True
                except:
                    continue
            
            return False
            
        except:
            return False
    
    def _handle_otp_verification(self):
        """確認コード（OTP）の処理"""
        try:
            # OTP入力フィールドを探す（指定されたセレクタを優先）
            otp_selectors = [
                "#input-box-otp",
                "input[name='otpCode']",
                "#auth-mfa-otpcode",
                "#cvf-input-code",
                "input[name='code']"
            ]
            
            otp_input = None
            for selector in otp_selectors:
                try:
                    otp_input = self.page.wait_for_selector(selector, timeout=15000)
                    if otp_input:
                        print(f"Found OTP input field: {selector}")
                        break
                except:
                    continue
            
            if not otp_input:
                print("No OTP verification required or OTP field not found")
                return True
            
            print("OTP verification required")
            
            # IMAPからOTPを取得
            otp_code = self.fetch_otp_from_email(max_wait=120)
            
            if otp_code:
                print("Entering OTP...")
                # 人間らしいタイピングで入力
                self.human_type(otp_input, otp_code)
                print("OTP entered successfully")
                self.random_sleep(1, 2)
                
                # 送信ボタンを探してクリック（指定されたセレクタを優先）
                submit_selectors = [
                    "#cvf-submit-otp-button",
                    "#auth-signin-button",
                    "input[type='submit']",
                    "button[type='submit']"
                ]
                
                for selector in submit_selectors:
                    try:
                        submit_btn = self.page.query_selector(selector)
                        if submit_btn:
                            print(f"Clicking submit button: {selector}")
                            self.human_click(submit_btn)
                            print("Submit button clicked")
                            self.random_sleep(2, 4)
                            break
                    except Exception as e:
                        print(f"Error clicking {selector}: {e}")
                        continue
                
                return True
            else:
                print("Could not fetch OTP automatically - manual input required")
                return False
                
        except Exception as e:
            print(f"Error handling OTP: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _wait_for_login_success(self):
        """ログイン成功を確認（最大30秒）"""
        print("Checking login success...")
        
        for i in range(15):
            time.sleep(2)
            
            try:
                current_url = self.page.url
                print(f"Current URL: {current_url[:60]}...")
                
                # ログインページ以外にいるか確認
                if "ap/signin" not in current_url and "ap/mfa" not in current_url and "ap/cvf" not in current_url:
                    # ログイン状態を確認
                    if self.is_logged_in():
                        print("Login confirmed!")
                        return True
                    
                    # Amazonのメインサイトにいる場合
                    if "amazon.co.jp" in current_url and "/ap/" not in current_url:
                        time.sleep(1)
                        if self.is_logged_in():
                            print("Login confirmed on Amazon page!")
                            return True
                
            except Exception as e:
                print(f"Error checking login: {e}")
        
        # 最終確認
        print("Final login check...")
        if self.is_logged_in():
            return True
        
        print("Login could not be confirmed, but continuing...")
        return True  # タイムアウトしても続行（手動で操作している可能性）
    
    def navigate_to_target_url(self):
        """指定URLにアクセス"""
        try:
            target_url = self.url if self.url else self.BASE_URL
            print(f"Navigating to: {target_url}")
            # commit: サーバーからレスポンスを受け取った時点で続行（最速）
            self.page.goto(target_url, wait_until="commit", timeout=90000)
            # DOMContentLoadedを待つ（全リソースは待たない）
            try:
                self.page.wait_for_load_state("domcontentloaded", timeout=30000)
            except:
                pass  # タイムアウトしても続行
            print(f"Page loaded: {self.page.title()}")
            return True
        except Exception as e:
            print(f"Failed to navigate: {e}", file=sys.stderr)
            return False
    
    def wait_for_user(self):
        """ユーザー操作を待機（Stopボタンが押されるまで）"""
        print("=" * 50)
        print("Browser mode active - You can now operate freely")
        print("The browser will stay open until you stop the task")
        print("=" * 50)
        
        # ストップフラグをチェックしながら待機（0.2秒ごと）
        try:
            while not self._stop_requested:
                time.sleep(0.2)
            print("Stop requested - ending browser session")
        except KeyboardInterrupt:
            print("Interrupted by user")
        except:
            print("Wait loop ended")
    
    def run(self):
        """メイン実行
        
        Returns:
            tuple: (success: bool, error_status: str or None)
        """
        import traceback
        
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
                self.start_browser()
            except Exception as e:
                print(f"Browser start failed: {e}")
                return False, "Failed Browser"
            
            # ★ ストップチェック
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            
            self.load_cookies()
            
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
            
            # ログイン後、確実に目的URLに遷移
            print("\x00STATUS:Navigating", flush=True)
            print("=" * 50)
            print("Navigating to target URL...")
            print("=" * 50)
            
            time.sleep(2)
            
            try:
                if not self.navigate_to_target_url():
                    print("Failed to navigate to target URL")
                    return False, "Failed Navigation"
            except Exception as e:
                print(f"Navigation error: {e}")
                return False, "Failed Navigation"
            
            # ナビゲーション成功を示すマーカーを出力
            print("\x00STATUS:Browsing", flush=True)
            print("NAVIGATION_COMPLETE")
            print("Navigation successful!")
            self.save_cookies()
            
            self.wait_for_user()
            
            self.save_cookies()
            
            return True, None
            
        except Exception as e:
            if self._stop_requested:
                print("Task stopped by user")
                return False, "Stopped"
            print(f"Error in run(): {e}", file=sys.stderr)
            traceback.print_exc()
            return False, "Failed Unknown"
        finally:
            self.close_browser()


def main():
    import traceback
    
    parser = argparse.ArgumentParser(description="Amazon Browser Mode Bot")
    parser.add_argument("--task-data", required=True, help="Path to task data JSON file")
    
    args = parser.parse_args()
    
    print(f"Task data file: {args.task_data}")
    
    try:
        with open(args.task_data, 'r', encoding='utf-8') as f:
            task_data = json.load(f)
        print(f"Task data loaded successfully")
    except Exception as e:
        print(f"Failed to load task data: {e}")
        traceback.print_exc()
        return
    
    try:
        bot = AmazonBrowser(task_data)
        success, error_status = bot.run()
        
        # エラーステータスがある場合は出力（GUIが読み取る）
        if error_status:
            print(f"BROWSER_ERROR:{error_status}")
        
        print(f"Bot finished with success={success}")
    except Exception as e:
        print(f"Bot execution error: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main()