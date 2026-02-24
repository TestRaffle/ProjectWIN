"""
ライセンス認証マネージャー
- ハードウェアID生成（MAC + マシン名のハッシュ）
- サーバーへのライセンス認証リクエスト
- ローカルキャッシュによるオフライン猶予（3日間）
- 定期的なハートビート（同時起動検知）
"""

import json
import hashlib
import uuid
import platform
import time
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime, timedelta


class LicenseManager:
    """ライセンス認証を管理するクラス"""
    
    # ===== 設定 =====
    # サーバーURL（サーバー構築後に変更）
    API_BASE_URL = "https://web-production-eca53.up.railway.app/api"
    
    # オフライン猶予期間（日数）
    OFFLINE_GRACE_DAYS = 3
    
    # ハートビート間隔（秒） - 30分
    HEARTBEAT_INTERVAL = 1800
    
    # ライセンスキャッシュファイル（初期値はNone、__init__で設定）
    CACHE_FILE = None
    
    @staticmethod
    def _is_compiled():
        """exe化されているか判定（PyInstaller/Nuitka両対応）"""
        import sys
        import os
        # PyInstaller
        if getattr(sys, 'frozen', False):
            return True
        # Nuitka - __compiled__ 属性をチェック
        if "__compiled__" in globals():
            return True
        # モジュール属性としてAPP_DIRが設定されている場合（サーバーから実行時）
        current_module = sys.modules.get(__name__)
        if current_module and hasattr(current_module, 'APP_DIR') and current_module.APP_DIR:
            return True
        # Nuitka standalone - パスに.distまたはProjectWINが含まれている場合
        exe_path = sys.executable.lower()
        if '.dist' in exe_path or 'projectwin' in exe_path:
            return True
        # Nuitka onefile - Tempフォルダ内のONEFILで実行されている場合
        if 'temp' in exe_path and 'onefil' in exe_path:
            return True
        return False
    
    @staticmethod
    def _get_base_dir():
        """アプリのベースディレクトリを取得"""
        import sys
        # モジュール属性としてAPP_DIRが設定されていればそれを使用（サーバーから実行時）
        current_module = sys.modules.get(__name__)
        if current_module and hasattr(current_module, 'APP_DIR') and current_module.APP_DIR:
            return current_module.APP_DIR
        # APP_DIRがグローバルで設定されていればそれを使用
        if 'APP_DIR' in globals() and globals()['APP_DIR']:
            return globals()['APP_DIR']
        # exe化されている場合
        if LicenseManager._is_compiled():
            exe_dir = Path(sys.executable).parent
            
            # 複数の階層を探索してProjectWIN.exeを探す
            for i in range(5):
                check_dir = exe_dir
                for _ in range(i):
                    check_dir = check_dir.parent
                
                if (check_dir / "ProjectWIN.exe").exists():
                    return check_dir
            
            # 見つからない場合は、パスから推測
            parts = exe_dir.parts
            for i, part in enumerate(parts):
                if 'projectwin' in part.lower() or part.endswith('.dist'):
                    return Path(*parts[:i+1])
            
            return exe_dir
        # 通常の.pyファイルとして実行されている場合
        if '__file__' in globals() and not str(__file__).startswith('<'):
            return Path(__file__).parent
        # フォールバック: カレントディレクトリ
        return Path.cwd()
    
    def __init__(self):
        # CACHE_FILEを動的に設定（_internal/settings内に保存）
        base_dir = self._get_base_dir()
        self.CACHE_FILE = base_dir / "_internal" / "settings" / ".license_cache"
        
        self._license_key = ""
        self._hardware_id = self._generate_hardware_id()
        self._session_token = ""
        self._cache_data = {}
        self._load_cache()
    
    # ===== ハードウェアID =====
    
    @staticmethod
    def _generate_hardware_id():
        """PCを識別するハードウェアIDを生成
        
        MACアドレス + マシン名 + プロセッサのハッシュ
        → 同一PCでは常に同じ値が返る
        """
        try:
            raw = f"{uuid.getnode()}-{platform.node()}-{platform.machine()}"
            return hashlib.sha256(raw.encode()).hexdigest()[:32]
        except Exception:
            # フォールバック: MACアドレスのみ
            return hashlib.sha256(str(uuid.getnode()).encode()).hexdigest()[:32]
    
    @property
    def hardware_id(self):
        return self._hardware_id
    
    # ===== キャッシュ管理 =====
    
    def _load_cache(self):
        """ローカルキャッシュを読み込み"""
        try:
            if self.CACHE_FILE.exists():
                with open(self.CACHE_FILE, 'r', encoding='utf-8') as f:
                    self._cache_data = json.load(f)
                self._license_key = self._cache_data.get("license_key", "")
        except Exception:
            self._cache_data = {}
    
    def _save_cache(self, license_key, valid_until=""):
        """認証成功時にキャッシュを保存"""
        try:
            self.CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
            
            self._cache_data = {
                "license_key": license_key,
                "hardware_id": self._hardware_id,
                "last_verified": datetime.now().isoformat(),
                "valid_until": valid_until,
                "cached_at": datetime.now().isoformat()
            }
            
            with open(self.CACHE_FILE, 'w', encoding='utf-8') as f:
                json.dump(self._cache_data, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"Failed to save license cache: {e}")
    
    def _clear_cache(self):
        """キャッシュを削除（認証失敗時）"""
        try:
            if self.CACHE_FILE.exists():
                self.CACHE_FILE.unlink()
            self._cache_data = {}
            self._license_key = ""
        except Exception:
            pass
    
    @property
    def cached_license_key(self):
        """キャッシュされたライセンスキーを取得"""
        return self._cache_data.get("license_key", "")
    
    def has_valid_cache(self):
        """オフライン猶予期間内の有効なキャッシュがあるか"""
        if not self._cache_data:
            return False
        
        cached_at = self._cache_data.get("cached_at", "")
        if not cached_at:
            return False
        
        try:
            cached_time = datetime.fromisoformat(cached_at)
            grace_deadline = cached_time + timedelta(days=self.OFFLINE_GRACE_DAYS)
            return datetime.now() < grace_deadline
        except Exception:
            return False
    
    # ===== サーバー通信 =====
    
    def _api_request(self, endpoint, data):
        """APIリクエストを送信
        
        Returns:
            dict: レスポンスJSON
            None: 通信失敗
        """
        try:
            url = f"{self.API_BASE_URL}/{endpoint}"
            
            req = urllib.request.Request(
                url,
                data=json.dumps(data).encode('utf-8'),
                headers={
                    'Content-Type': 'application/json',
                    'User-Agent': 'BrowserAutomation/1.0'
                },
                method='POST'
            )
            
            with urllib.request.urlopen(req, timeout=10) as response:
                return json.loads(response.read().decode('utf-8'))
                
        except urllib.error.HTTPError as e:
            try:
                body = json.loads(e.read().decode('utf-8'))
                return body  # サーバーからのエラーレスポンス
            except Exception:
                return None
        except Exception:
            return None
    
    def activate(self, license_key):
        """ライセンスをアクティベート（初回 or 新規PC）
        
        Args:
            license_key: ユーザーが入力したライセンスキー
            
        Returns:
            tuple: (success: bool, message: str)
        """
        self._license_key = license_key
        
        response = self._api_request("license/activate", {
            "license_key": license_key,
            "hardware_id": self._hardware_id,
            "machine_name": platform.node()
        })
        
        # サーバー接続不可
        if response is None:
            # キャッシュがあればオフライン猶予
            if self.cached_license_key == license_key and self.has_valid_cache():
                return True, "Offline mode (cached license)"
            return False, "Cannot connect to server"
        
        # サーバーからの応答
        success = response.get("success", False)
        message = response.get("message", "Unknown error")
        
        if success:
            self._session_token = response.get("session_token", "")
            valid_until = response.get("valid_until", "")
            self._save_cache(license_key, valid_until)
            return True, message
        else:
            error_code = response.get("error_code", "")
            
            if error_code == "ALREADY_ACTIVATED":
                return False, "This key is already in use on another PC"
            elif error_code == "INVALID_KEY":
                self._clear_cache()
                return False, "Invalid license key"
            elif error_code == "EXPIRED":
                self._clear_cache()
                return False, "License expired - please renew your subscription"
            elif error_code == "SUSPENDED":
                self._clear_cache()
                return False, "License suspended - please contact support"
            else:
                return False, message
    
    def verify(self):
        """ライセンスを検証（起動時、キャッシュあり）
        
        Returns:
            tuple: (success: bool, message: str)
        """
        license_key = self.cached_license_key
        if not license_key:
            return False, "No license found"
        
        response = self._api_request("license/verify", {
            "license_key": license_key,
            "hardware_id": self._hardware_id
        })
        
        # サーバー接続不可 → オフライン猶予
        if response is None:
            if self.has_valid_cache():
                return True, "Offline mode"
            return False, "Cannot connect to server and cache expired"
        
        success = response.get("success", False)
        message = response.get("message", "")
        
        if success:
            self._license_key = license_key
            self._session_token = response.get("session_token", "")
            valid_until = response.get("valid_until", "")
            self._save_cache(license_key, valid_until)
            return True, message
        else:
            error_code = response.get("error_code", "")
            if error_code in ("EXPIRED", "SUSPENDED", "INVALID_KEY", "NOT_ACTIVATED"):
                self._clear_cache()
            return False, message
    
    def heartbeat(self):
        """ハートビート送信（同時起動検知用）
        
        Returns:
            bool: 継続利用可能か
        """
        if not self._license_key:
            return True  # キーなしの場合は無視
        
        response = self._api_request("license/heartbeat", {
            "license_key": self._license_key,
            "hardware_id": self._hardware_id,
            "session_token": self._session_token
        })
        
        if response is None:
            return True  # 通信失敗時は継続許可
        
        return response.get("success", True)
    
    def deactivate(self):
        """ライセンスのアクティベーションを解除（PC変更時）
        
        Returns:
            tuple: (success: bool, message: str)
        """
        license_key = self._license_key or self.cached_license_key
        if not license_key:
            return False, "No license to deactivate"
        
        response = self._api_request("license/deactivate", {
            "license_key": license_key,
            "hardware_id": self._hardware_id
        })
        
        self._clear_cache()
        
        if response and response.get("success"):
            return True, "License deactivated"
        return False, "Failed to deactivate"


# ===== オフラインテスト用のモックマネージャー =====

class OfflineLicenseManager(LicenseManager):
    """サーバーなしでテストできるモック版
    
    - サーバー構築前はこちらを使用
    - 任意のキーで認証成功する
    - 本番環境では LicenseManager に差し替える
    """
    
    # テスト用のキー（任意の文字列で認証通過）
    # 本番では削除してLicenseManagerを使用
    TEST_KEYS = {"TEST-1234-5678-ABCD", "DEV-0000-0000-0000"}
    
    def activate(self, license_key):
        # テストキーなら常に成功
        if license_key in self.TEST_KEYS:
            self._license_key = license_key
            self._save_cache(license_key, "2099-12-31")
            return True, "License activated (test mode)"
        
        # 空でなければ通常のサーバー認証を試行
        if license_key.strip():
            result = super().activate(license_key)
            # サーバー未構築でも、キーが入力されていればテスト通過
            if result[0] is False and "Cannot connect" in result[1]:
                self._license_key = license_key
                self._save_cache(license_key, "2099-12-31")
                return True, "License activated (offline test mode)"
            return result
        
        return False, "Please enter a license key"
    
    def verify(self):
        license_key = self.cached_license_key
        if license_key:
            self._license_key = license_key
            return True, "License valid (cached)"
        return False, "No license found"
    
    def heartbeat(self):
        return True