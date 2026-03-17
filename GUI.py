"""
ブラウザ自動操作 GUI アプリケーション
PySide6を使用した左サイドバー + 右メインコンテンツのレイアウト
"""

import sys
import os
import subprocess
import json
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QStackedWidget, QPushButton, QLabel, QLineEdit, QTextEdit,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QSpinBox, QCheckBox, QGroupBox, QFormLayout, QFrame,
    QSplitter, QListWidget, QListWidgetItem, QFileDialog, QMessageBox,
    QTabWidget, QGraphicsOpacityEffect, QProgressBar, QRadioButton,
    QMenu, QDialog, QPlainTextEdit
)
from PySide6.QtCore import Qt, QSize, QThread, Signal, QTimer, QPropertyAnimation, QEasingCurve, QPoint
from PySide6.QtGui import QFont, QIcon, QPalette, QColor, QPainter, QPen, QBrush, QPixmap
import base64

# ===== セキュリティ機能 =====
import hashlib as _hl
import zlib as _zl

def _xk(s):
    """キー復元（難読化）"""
    return bytes([c ^ 0x42 for c in s])

# 難読化された暗号化キー（実行時に復元）
_EK1 = _xk(b"\x120-('!6\x15\x0b\x0c\x1d\x00-6\x1d\t';\x1dprpv\x1d\x11'!70'cc")  # Bot用
_EK2 = _xk(b"\x120-('!6\x15\x0b\x0c\x1d\x01-0'\x1d\t';\x1dprpv\x1d\x11'!70'c")  # Core用

def _chk_dbg():
    """デバッガー検知"""
    try:
        import ctypes
        if hasattr(ctypes, 'windll'):
            return ctypes.windll.kernel32.IsDebuggerPresent() != 0
    except:
        pass
    return False

def _vfy(data: bytes, expected_crc: int = None) -> bool:
    """データ整合性チェック"""
    if expected_crc is None:
        return True
    return _zl.crc32(data) == expected_crc

# ===== アイコン管理（メモリのみ方式） =====
# アイコンのダウンロードURL（GitHub Raw）
ICONS_BASE_URL = "https://raw.githubusercontent.com/TestRaffle/projectwin-assets/main/"

# ===== Bot管理（メモリのみ方式） =====
# BotのダウンロードURL（GitHub Raw - 暗号化ファイル）
BOTS_BASE_URL = "https://raw.githubusercontent.com/TestRaffle/projectwin-assets/main/bots/"

# メモリ上のBotコードキャッシュ
_bot_code_cache = {}

# メモリ上のアイコンキャッシュ
_icon_cache = {}

def _download_icon(name: str) -> bytes:
    """サーバーからアイコンをダウンロード（メモリのみ、ファイル保存なし）"""
    import urllib.request
    import urllib.error
    
    url = f"{ICONS_BASE_URL}{name}.png"
    
    try:
        with urllib.request.urlopen(url, timeout=10) as response:
            return response.read()
    except Exception as e:
        print(f"アイコンダウンロードエラー: {name} - {e}")
        return None

def get_icon_from_base64(name):
    """アイコンを取得（メモリキャッシュ → ダウンロード）"""
    global _icon_cache
    
    # 1. メモリキャッシュをチェック
    if name in _icon_cache:
        return _icon_cache[name]
    
    # 2. サーバーからダウンロード（ファイルには保存しない）
    icon_data = _download_icon(name)
    
    # 3. アイコンを生成してメモリキャッシュに保存
    if icon_data:
        pixmap = QPixmap()
        pixmap.loadFromData(icon_data)
        icon = QIcon(pixmap)
        _icon_cache[name] = icon
        return icon
    
    # 失敗した場合は空のアイコン
    empty_icon = QIcon()
    _icon_cache[name] = empty_icon
    return empty_icon


# ===== Bot管理関数（メモリのみ方式） =====
def _dcd(data: bytes, key: bytes) -> bytes:
    """汎用XOR復号"""
    result = bytearray()
    for i, byte in enumerate(data):
        result.append(byte ^ key[i % len(key)])
    return bytes(result)

def download_bot_code(site: str, mode: str) -> str:
    """サーバーから暗号化されたBotコードをダウンロードしてメモリ上で復号
    ローカルに.pyファイルがあればそれを優先して使用（開発・テスト用）
    """
    global _bot_code_cache
    
    # デバッガー検知
    if _chk_dbg():
        return None
    
    cache_key = f"{site}_{mode}"
    
    # 1. メモリキャッシュをチェック
    if cache_key in _bot_code_cache:
        return _bot_code_cache[cache_key]
    
    # 2. ローカルの.pyファイルをチェック（_internal/bots/site/site_mode.py）
    local_bot_path = APP_DIR / "_internal" / "bots" / site / f"{site}_{mode}.py"
    print(f"[DEBUG] Looking for local bot: {local_bot_path}")
    print(f"[DEBUG] APP_DIR: {APP_DIR}")
    print(f"[DEBUG] File exists: {local_bot_path.exists()}")
    if local_bot_path.exists():
        try:
            with open(local_bot_path, 'r', encoding='utf-8') as f:
                bot_code = f.read()
            print(f"[DEV] Loaded local bot: {local_bot_path}")
            _bot_code_cache[cache_key] = bot_code
            return bot_code
        except Exception as e:
            print(f"[DEV] Failed to load local bot: {e}")
    
    # 3. サーバーからダウンロード
    import urllib.request
    import urllib.error
    
    # 暗号化ファイル名: amazon_signup.enc
    url = f"{BOTS_BASE_URL}{site}/{site}_{mode}.enc"
    
    try:
        with urllib.request.urlopen(url, timeout=30) as response:
            encrypted_data = response.read()
        
        # 4. メモリ上で復号（難読化されたキーを使用）
        decrypted_data = _dcd(encrypted_data, _EK1)
        bot_code = decrypted_data.decode('utf-8')
        
        # 5. メモリキャッシュに保存（ファイルには保存しない）
        _bot_code_cache[cache_key] = bot_code
        
        return bot_code
    except Exception as e:
        print(f"Botダウンロードエラー: {cache_key} - {e}")
        return None

def get_bot_module(site: str, mode: str, task_data: dict):
    """Botコードをダウンロードしてモジュールとして返す"""
    import types
    
    bot_code = download_bot_code(site, mode)
    if bot_code is None:
        return None
    
    try:
        # コードをコンパイルして実行
        bot_module = types.ModuleType(f"bot_{site}_{mode}")
        bot_module.__file__ = f"<server>/{site}_{mode}.py"
        bot_module.__builtins__ = __builtins__
        
        # APP_DIRをexec前にモジュール辞書に追加（botコード内で使えるように）
        bot_module.__dict__['APP_DIR'] = APP_DIR
        
        exec(bot_code, bot_module.__dict__)
        
        return bot_module
    except Exception as e:
        print(f"Bot実行エラー: {site}_{mode} - {e}")
        import traceback
        traceback.print_exc()
        return None


# ===== コアモジュール管理（メモリのみ方式） =====
# コアモジュールのダウンロードURL
CORE_MODULES_BASE_URL = "https://raw.githubusercontent.com/TestRaffle/projectwin-assets/main/core/"

# メモリ上のコアモジュールキャッシュ
_core_module_cache = {}

def download_core_module(module_name: str):
    """サーバーからコアモジュールをダウンロードしてメモリ上で復号・実行"""
    global _core_module_cache
    
    # デバッガー検知
    if _chk_dbg():
        return None
    
    # 1. メモリキャッシュをチェック
    if module_name in _core_module_cache:
        return _core_module_cache[module_name]
    
    # 2. サーバーからダウンロード
    import urllib.request
    import urllib.error
    import types
    
    url = f"{CORE_MODULES_BASE_URL}{module_name}.enc"
    
    try:
        print(f"Loading {module_name}...")
        with urllib.request.urlopen(url, timeout=30) as response:
            encrypted_data = response.read()
        
        # 3. メモリ上で復号（難読化されたキーを使用）
        decrypted_data = _dcd(encrypted_data, _EK2)
        module_code = decrypted_data.decode('utf-8')
        
        # 4. モジュールとして実行（APP_DIRを渡す）
        module = types.ModuleType(module_name)
        module.__file__ = f"<server>/{module_name}.py"
        module.APP_DIR = APP_DIR  # APP_DIRを渡す
        exec(module_code, module.__dict__)
        
        # 5. メモリキャッシュに保存
        _core_module_cache[module_name] = module
        
        return module
    except Exception as e:
        print(f"コアモジュールダウンロードエラー: {module_name} - {e}")
        return None

def load_license_manager():
    """LicenseManagerをサーバーから取得"""
    module = download_core_module("license_manager")
    if module and hasattr(module, 'LicenseManager'):
        return module.LicenseManager
    return None

def load_updater_functions():
    """updaterの関数をサーバーから取得"""
    module = download_core_module("updater")
    if module:
        return {
            'check_for_update': getattr(module, 'check_for_update', None),
            'download_update': getattr(module, 'download_update', None),
            'apply_update': getattr(module, 'apply_update', None),
            'restart_app': getattr(module, 'restart_app', None),
            'save_version': getattr(module, 'save_version', None),
        }
    return None


try:
    import openpyxl
except ImportError:
    print("openpyxlがインストールされていません。pip install openpyxl を実行してください。")


def is_compiled():
    """exe化されているか判定（PyInstaller/Nuitka両対応）"""
    # PyInstaller
    if getattr(sys, 'frozen', False):
        return True
    # Nuitka - __compiled__ 属性をチェック
    if "__compiled__" in globals():
        return True
    main_module = sys.modules.get('__main__', None)
    if main_module and hasattr(main_module, '__compiled__'):
        return True
    # Nuitka standalone - パスに.distまたはProjectWINが含まれている場合
    exe_path = sys.executable.lower()
    if '.dist' in exe_path or 'projectwin' in exe_path:
        return True
    # Nuitka onefile - Tempフォルダ内のONEFILで実行されている場合
    if 'temp' in exe_path and 'onefil' in exe_path:
        return True
    return False


def get_app_dir():
    """アプリのルートディレクトリを取得（PyInstaller/Nuitka両対応）"""
    if is_compiled():
        # Nuitka standalone の場合
        exe_dir = Path(sys.executable).parent
        
        # 複数の階層を探索してProjectWIN.exeを探す
        for i in range(5):  # 最大5階層上まで探索
            check_dir = exe_dir
            for _ in range(i):
                check_dir = check_dir.parent
            
            if (check_dir / "ProjectWIN.exe").exists():
                return check_dir
        
        # 見つからない場合は、パスから推測
        # パスに"ProjectWIN"または".dist"が含まれるフォルダを探す
        parts = exe_dir.parts
        for i, part in enumerate(parts):
            if 'projectwin' in part.lower() or part.endswith('.dist'):
                return Path(*parts[:i+1])
        
        # それでも見つからない場合はexeの場所を返す
        return exe_dir
    else:
        # 開発中: スクリプトの場所
        return Path(__file__).parent


# アプリのベースディレクトリ（コアモジュール読み込み前に定義が必要）
APP_DIR = get_app_dir()

# 設定フォルダ（_internal内に配置）
SETTINGS_DIR = APP_DIR / "_internal" / "settings"
SETTINGS_DIR.mkdir(parents=True, exist_ok=True)


# ライセンス認証（サーバーからダウンロード、フォールバックとしてローカル）
LicenseManager = None
try:
    # まずサーバーから取得を試みる
    LicenseManager = load_license_manager()
except:
    pass

if LicenseManager is None:
    # フォールバック: ローカルファイル（開発用）
    try:
        from license_manager import LicenseManager
    except ImportError:
        LicenseManager = None

# 自動アップデーター（サーバーからダウンロード、フォールバックとしてローカル）
check_for_update = None
download_update = None
apply_update = None
restart_app = None
save_version = None

try:
    # まずサーバーから取得を試みる
    updater_funcs = load_updater_functions()
    if updater_funcs:
        check_for_update = updater_funcs['check_for_update']
        download_update = updater_funcs['download_update']
        apply_update = updater_funcs['apply_update']
        restart_app = updater_funcs['restart_app']
        save_version = updater_funcs['save_version']
except:
    pass

if check_for_update is None:
    # フォールバック: ローカルファイル（開発用）
    try:
        from updater import check_for_update, download_update, apply_update, restart_app, save_version
    except ImportError:
        check_for_update = None


def get_app_version():
    """アプリのバージョンを取得"""
    try:
        if is_compiled():
            # exe化されている場合: _internal内を参照
            version_path = APP_DIR / "_internal" / "version.json"
        else:
            # 開発中: APP_DIR内を参照
            version_path = APP_DIR / "version.json"
        
        if version_path.exists():
            content = version_path.read_text().strip()
            # JSON形式の場合
            if content.startswith('{'):
                data = json.loads(content)
                return data.get("version", "0.0.0")
            # プレーンテキストの場合（vプレフィックスを除去）
            return content.lstrip('v').strip()
    except:
        pass
    return "0.0.0"




class CheckmarkCheckBox(QCheckBox):
    """緑色のチェックマーク付きカスタムチェックボックス"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(22, 22)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # ボックスの描画
        rect = self.rect().adjusted(2, 2, -2, -2)
        
        if self.isChecked():
            # チェック時: 緑の枠線
            pen = QPen(QColor("#27ae60"))
            pen.setWidth(2)
            painter.setPen(pen)
            painter.setBrush(QBrush(Qt.GlobalColor.transparent))
            painter.drawRoundedRect(rect, 4, 4)
            
            # チェックマークを描画（細めでスタイリッシュに）
            pen = QPen(QColor("#27ae60"))
            pen.setWidth(2)
            pen.setCapStyle(Qt.PenCapStyle.RoundCap)
            pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen)
            
            # チェックマークの座標（少し小さめに）
            painter.drawLine(6, 11, 9, 14)
            painter.drawLine(9, 14, 16, 6)
        else:
            # 未チェック時: グレーの枠線
            if self.underMouse():
                pen = QPen(QColor("#27ae60"))
            else:
                pen = QPen(QColor("#505050"))
            pen.setWidth(2)
            painter.setPen(pen)
            painter.setBrush(QBrush(Qt.GlobalColor.transparent))
            painter.drawRoundedRect(rect, 4, 4)
        
        painter.end()


class SwitchButton(QPushButton):
    """カスタムスイッチボタン"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCheckable(True)
        self.setFixedSize(50, 26)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._update_style()
        self.toggled.connect(self._update_style)
    
    def _update_style(self):
        if self.isChecked():
            self.setStyleSheet("""
                QPushButton {
                    background-color: #27ae60;
                    border-radius: 13px;
                    border: none;
                    text-align: right;
                    padding-right: 5px;
                }
            """)
            self.setText("●")
        else:
            self.setStyleSheet("""
                QPushButton {
                    background-color: #404050;
                    border-radius: 13px;
                    border: none;
                    text-align: left;
                    padding-left: 5px;
                    color: #808080;
                }
            """)
            self.setText("●")


class ToastNotification(QWidget):
    """アニメーション付きトースト通知（親ウィジェット内に表示）"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        # 親ウィジェット内の子ウィジェットとして表示（別ウィンドウではなく）
        self.setFixedHeight(50)
        
        # メインレイアウト
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        # コンテナ
        self.container = QWidget()
        self.container.setObjectName("toast_container")
        container_layout = QHBoxLayout(self.container)
        container_layout.setContentsMargins(15, 8, 10, 8)
        container_layout.setSpacing(10)
        
        # アイコン
        self.icon_label = QLabel()
        self.icon_label.setFixedSize(20, 20)
        self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        container_layout.addWidget(self.icon_label)
        
        # メッセージ
        self.message_label = QLabel()
        self.message_label.setStyleSheet("color: #ffffff; font-size: 13px; background: transparent;")
        container_layout.addWidget(self.message_label)
        
        container_layout.addStretch()
        
        # プログレスバー（タイマー表示用）- 右から左に向かって減る
        self.progress = QProgressBar()
        self.progress.setFixedSize(60, 4)
        self.progress.setTextVisible(False)
        self.progress.setRange(0, 100)
        self.progress.setValue(100)
        self.progress.setLayoutDirection(Qt.LayoutDirection.RightToLeft)  # 右から左へ
        self.progress.setStyleSheet("""
            QProgressBar { background-color: rgba(255,255,255,0.3); border-radius: 2px; }
            QProgressBar::chunk { background-color: #ffffff; border-radius: 2px; }
        """)
        container_layout.addWidget(self.progress)
        
        # 閉じるボタン
        self.close_btn = QPushButton("×")
        self.close_btn.setFixedSize(20, 20)
        self.close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.close_btn.setStyleSheet("""
            QPushButton { background: transparent; color: #808080; border: none; font-size: 14px; font-weight: bold; }
            QPushButton:hover { color: #e74c3c; }
        """)
        self.close_btn.clicked.connect(self.hide_toast)
        container_layout.addWidget(self.close_btn)
        
        layout.addWidget(self.container)
        
        # タイマー
        self.timer = QTimer()
        self.timer.timeout.connect(self._update_progress)
        self.progress_value = 100
        self.duration = 3000
        
        self.hide()
    
    def show_toast(self, message, toast_type="success", duration=3000):
        """トーストを表示"""
        self.duration = duration
        self.progress_value = 100
        self.progress.setValue(100)
        
        # タイプに応じた文字色とアイコン
        if toast_type == "success":
            text_color = "#2ecc71"  # 緑
            icon = "✓"
        elif toast_type == "error":
            text_color = "#e74c3c"  # 赤
            icon = "✗"
        elif toast_type == "warning":
            text_color = "#f39c12"  # 黄
            icon = "⚠"
        else:
            text_color = "#3498db"  # 青
            icon = "ℹ"
        
        # GUIメイン画面と同じ背景色 + シャドウ効果
        self.container.setStyleSheet(f"""
            #toast_container {{
                background-color: #1e1e2e;
                border: 1px solid #404050;
                border-radius: 8px;
            }}
        """)
        
        # シャドウ効果を追加
        from PySide6.QtWidgets import QGraphicsDropShadowEffect
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setXOffset(0)
        shadow.setYOffset(4)
        shadow.setColor(QColor(0, 0, 0, 100))
        self.container.setGraphicsEffect(shadow)
        
        self.icon_label.setText(icon)
        self.icon_label.setStyleSheet(f"color: {text_color}; font-size: 16px; font-weight: bold; background: transparent;")
        self.message_label.setText(message)
        self.message_label.setStyleSheet(f"color: {text_color}; font-size: 13px; font-weight: bold; background: transparent;")
        
        # プログレスバーの色も文字色に合わせる
        self.progress.setStyleSheet(f"""
            QProgressBar {{ background-color: #404050; border-radius: 2px; }}
            QProgressBar::chunk {{ background-color: {text_color}; border-radius: 2px; }}
        """)
        
        # 親ウィジェットの下部（Start All/Stop Allと同じ高さ）に配置
        if self.parent():
            parent_rect = self.parent().rect()
            toast_width = min(350, parent_rect.width() - 40)
            self.setFixedWidth(toast_width)
            # 右寄せで下部に配置
            x = parent_rect.width() - toast_width - 30
            y = parent_rect.height() - 80  # 下から80pxの位置
            self.move(x, y)
        
        self.show()
        self.raise_()
        
        # プログレスバーのアニメーション開始
        self.timer.start(30)
    
    def _update_progress(self):
        """プログレスバーを更新（左から右に減る）"""
        decrement = 100 / (self.duration / 30)
        self.progress_value -= decrement
        
        if self.progress_value <= 0:
            self.hide_toast()
        else:
            # 左から減るように値を反転
            self.progress.setValue(int(self.progress_value))
    
    def hide_toast(self):
        """トーストを非表示"""
        self.timer.stop()
        self.hide()


class SidebarButton(QPushButton):
    """サイドバー用のカスタムボタン（画像アイコン対応）"""
    
    def __init__(self, text, icon_name="", parent=None):
        super().__init__(parent)
        self.full_text = text
        self.icon_name = icon_name
        self.setCheckable(True)
        self.setFixedHeight(50)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        
        # アイコンを設定（Base64から）
        if icon_name:
            icon = get_icon_from_base64(icon_name)
            if not icon.isNull():
                self.setIcon(icon)
                self.setIconSize(QSize(20, 20))
        
        self.setText(f"  {text}" if icon_name else text)
        self.setStyleSheet(self._get_style(expanded=True))
    
    def set_expanded(self, expanded):
        """展開/折りたたみ状態を設定"""
        if expanded:
            self.setText(f"  {self.full_text}" if self.icon_name else self.full_text)
            self.setStyleSheet(self._get_style(expanded=True))
        else:
            self.setText("")  # 折りたたみ時はテキストなし、アイコンのみ
            self.setStyleSheet(self._get_style(expanded=False))
    
    def _get_style(self, expanded=True):
        alignment = "left" if expanded else "center"
        padding = "12px 15px 12px 8px" if expanded else "12px 5px"
        return f"""
            QPushButton {{
                background-color: transparent;
                color: #808080;
                border: none;
                border-radius: 8px;
                padding: {padding};
                margin-right: 12px;
                text-align: {alignment};
                font-size: 14px;
                font-weight: 500;
            }}
            QPushButton:hover {{
                color: #b0b0b0;
            }}
            QPushButton:checked {{
                background-color: transparent;
                color: #ffffff;
            }}
        """
    
    def paintEvent(self, event):
        """選択時に短いボーダーラインを描画"""
        super().paintEvent(event)
        if self.isChecked():
            from PySide6.QtGui import QPainter, QColor
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing)
            painter.setBrush(QColor("#4a90d9"))
            painter.setPen(Qt.PenStyle.NoPen)
            # ボタンの高さの60%の長さで中央に配置
            line_height = int(self.height() * 0.5)
            y_offset = (self.height() - line_height) // 2
            painter.drawRoundedRect(0, y_offset, 3, line_height, 1, 1)
            painter.end()


class BotWorker(QThread):
    """Botを別スレッドで実行するワーカー（サーバーダウンロード方式）"""
    status_changed = Signal(int, str)
    result_changed = Signal(int, str)
    finished_task = Signal(int)
    raffle_result = Signal(int, list)
    log_updated = Signal(int, str)
    
    def __init__(self, row, task_data):
        super().__init__()
        self.row = row
        self.task_data = task_data
        self._is_running = True
        self._stopped_by_user = False
        self._stop_requested = False  # ストップリクエストフラグ
        self._raffle_results = []
        self.log_lines = []
        self.bot_instance = None  # botインスタンスを保持（stop用）
    
    def get_log(self):
        return "\n".join(self.log_lines)
    
    def _timestamp(self):
        from datetime import datetime
        return datetime.now().strftime("%H:%M:%S")
    
    def _capture_print(self, text):
        """printをキャプチャしてステータス更新"""
        text = str(text).strip()
        if not text:
            return
        
        # \x00STATUS:で始まる行はステータスを更新（ログには追加しない、コンソールにも表示しない）
        # status_changedシグナルを使用（webhookなし）
        if "\x00STATUS:" in text:
            try:
                status = text.split("\x00STATUS:")[1].strip()
                self.status_changed.emit(self.row, status)
            except:
                pass
            return  # ログには追加しない
        
        # 旧形式のSTATUS:もサポート（後方互換性）
        if text.startswith("STATUS:"):
            try:
                status = text.split("STATUS:")[1].strip()
                self.status_changed.emit(self.row, status)
            except:
                pass
            return  # ログには追加しない
        
        # ログに追加
        self.log_lines.append(f"[{self._timestamp()}] {text}")
        self.log_updated.emit(self.row, text)
    
    def run(self):
        import io
        import contextlib
        import threading
        
        try:
            self.status_changed.emit(self.row, "Starting")
            self.log_lines.append(f"[{self._timestamp()}] Task started")
            
            # サイトとモードを取得
            mode = self.task_data.get("Mode", "").lower()
            site = self.task_data.get("Site", "").lower()
            
            self.log_lines.append(f"[{self._timestamp()}] Loading bot: {site}_{mode}")
            
            # サーバーからBotモジュールを取得（メモリのみ）
            bot_module = get_bot_module(site, mode, self.task_data)
            
            if bot_module is None:
                self.result_changed.emit(self.row, "Error: Cannot load bot")
                return
            
            # クラス名を決定
            class_name = f"{site.capitalize()}{mode.capitalize()}"
            
            if not hasattr(bot_module, class_name):
                for attr_name in dir(bot_module):
                    if attr_name.lower() == class_name.lower():
                        class_name = attr_name
                        break
                else:
                    self.result_changed.emit(self.row, f"Error: Class {class_name} not found")
                    return
            
            bot_class = getattr(bot_module, class_name)
            
            # Botインスタンスを作成
            self.bot_instance = bot_class(self.task_data)
            
            self.result_changed.emit(self.row, "Running")
            
            # スレッドセーフなprintキャプチャ
            # 各スレッドのIDとコールバックを紐付ける
            current_thread_id = threading.current_thread().ident
            
            class ThreadSafePrintCapture:
                # クラス変数でスレッドごとのコールバックを管理
                _callbacks = {}
                _lock = threading.Lock()
                
                def __init__(self, original):
                    self.original = original
                
                @classmethod
                def register(cls, thread_id, callback):
                    with cls._lock:
                        cls._callbacks[thread_id] = callback
                
                @classmethod
                def unregister(cls, thread_id):
                    with cls._lock:
                        if thread_id in cls._callbacks:
                            del cls._callbacks[thread_id]
                
                def write(self, text):
                    if text.strip():
                        thread_id = threading.current_thread().ident
                        with self._lock:
                            callback = self._callbacks.get(thread_id)
                        if callback:
                            callback(text)
                    self.original.write(text)
                
                def flush(self):
                    self.original.flush()
            
            # printをキャプチャしながら実行
            import sys
            original_stdout = sys.stdout
            
            # 既にThreadSafePrintCaptureがインストールされているか確認
            if not isinstance(sys.stdout, ThreadSafePrintCapture):
                sys.stdout = ThreadSafePrintCapture(original_stdout)
            
            # このスレッドのコールバックを登録
            ThreadSafePrintCapture.register(current_thread_id, self._capture_print)
            
            result = None
            run_error = None
            try:
                result = self.bot_instance.run()
            except Exception as e:
                run_error = e
            finally:
                # このスレッドのコールバックを解除
                ThreadSafePrintCapture.unregister(current_thread_id)
            
            # Bot側のブラウザ閉じフラグをチェック
            if hasattr(self.bot_instance, '_browser_closed') and self.bot_instance._browser_closed:
                self._stop_requested = True
            if hasattr(self.bot_instance, '_stop_requested') and self.bot_instance._stop_requested:
                self._stop_requested = True
            
            # ストップフラグが設定されていた場合の処理
            if self._stopped_by_user or self._stop_requested:
                # botが成功を返した場合（Browserモードで正常終了など）はその結果を使う
                if isinstance(result, tuple) and len(result) >= 2 and result[0] == True:
                    pass  # 下の結果解析処理に進む
                else:
                    self.result_changed.emit(self.row, "Stopped")
                    return
            
            # 例外が発生していた場合は再raise
            if run_error:
                raise run_error
            
            # 結果を解析
            if isinstance(result, tuple):
                if len(result) >= 3:
                    success, results, error_status = result
                elif len(result) == 2:
                    success, second = result
                    # 2番目の要素が文字列ならerror_status（失敗時）またはsuccess_status（成功時）
                    if isinstance(second, str):
                        if success:
                            # 成功時の文字列はsuccess_statusとして扱う
                            error_status = second  # error_statusを再利用（result_changedで使用）
                        else:
                            error_status = second
                        results = None
                    else:
                        results = second
                        error_status = None
                else:
                    success = result[0] if result else False
                    results = []
                    error_status = None
            else:
                success = bool(result)
                results = []
                error_status = None
            
            # Raffle結果があれば送信（辞書形式のリストの場合のみ）
            if results and isinstance(results, list) and len(results) > 0:
                # 最初の要素が辞書かどうかでRaffle結果かを判定
                if isinstance(results[0], dict):
                    self._raffle_results = results
                    self.raffle_result.emit(self.row, results)
            
            # 結果を設定
            if error_status:
                self.result_changed.emit(self.row, error_status)
            elif success:
                self.result_changed.emit(self.row, "Success")
            else:
                self.result_changed.emit(self.row, "Failed")
            
        except Exception as e:
            # ストップによる例外の場合はStoppedとして処理
            if self._stopped_by_user or self._stop_requested:
                self.result_changed.emit(self.row, "Stopped")
            else:
                import traceback
                traceback.print_exc()
                self.result_changed.emit(self.row, f"Error: {str(e)[:50]}")
        finally:
            self.finished_task.emit(self.row)
    
    def stop(self):
        """タスクを停止（フラグを設定）"""
        self._is_running = False
        self._stopped_by_user = True
        self._stop_requested = True
        
        # Bot側にもストップフラグを設定
        if self.bot_instance:
            self.bot_instance._stop_requested = True
            print("Stop requested")


class TaskPage(QWidget):
    """タスクページ"""
    
    CONFIG_FILE = SETTINGS_DIR / "app_config.json"
    GENERAL_SETTINGS_FILE = SETTINGS_DIR / "general_settings.json"
    
    def __init__(self, proxy_page=None):
        super().__init__()
        self.proxy_page = proxy_page
        self.settings_page = None  # Webhook送信用の参照（MainWindowで設定）
        self.workers = {}
        self.current_excel_path = None
        self.last_checked_row = None
        self.all_task_data = []
        self.original_task_data = []  # Excelから読み込んだ元のデータ
        self.pending_tasks = []  # 待機中のタスク（並列制御用）
        self.setup_ui()
        self._load_last_excel()
    
    def _load_last_excel(self):
        """前回読み込んだExcelファイルを自動読み込み"""
        try:
            if self.CONFIG_FILE.exists():
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                last_excel = config.get("last_excel_path", "")
                if last_excel and os.path.exists(last_excel):
                    self._load_excel_file(last_excel, show_toast=False)
        except Exception as e:
            print(f"Failed to load last Excel: {e}")
    
    def _save_last_excel(self, path):
        """最後に読み込んだExcelファイルパスを保存"""
        try:
            self.CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
            config = {}
            if self.CONFIG_FILE.exists():
                with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            config["last_excel_path"] = path
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save last Excel path: {e}")
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        # トースト通知
        self.toast = ToastNotification(self)
        
        # ヘッダー部分
        header_layout = QHBoxLayout()
        header_layout.setSpacing(8)
        
        self.header_label = QLabel("Tasks(0)")
        self.header_label.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #ffffff;
        """)
        header_layout.addWidget(self.header_label)
        
        header_layout.addStretch()
        
        # Headlessチェックボックス
        headless_container = QWidget()
        headless_layout = QHBoxLayout(headless_container)
        headless_layout.setContentsMargins(0, 0, 15, 0)
        headless_layout.setSpacing(8)
        
        self.headless_checkbox = CheckmarkCheckBox()
        headless_layout.addWidget(self.headless_checkbox)
        
        headless_label = QLabel("Headless")
        headless_label.setStyleSheet("color: #ffffff; font-size: 14px; font-weight: bold;")
        headless_label.setCursor(Qt.CursorShape.PointingHandCursor)
        headless_label.mousePressEvent = lambda e: self.headless_checkbox.setChecked(not self.headless_checkbox.isChecked())
        headless_layout.addWidget(headless_label)
        
        header_layout.addWidget(headless_container)
        
        # 時計表示
        self.clock_label = QLabel()
        self.clock_label.setStyleSheet("color: #b0b0b0; font-size: 14px; font-family: monospace; margin-right: 15px;")
        header_layout.addWidget(self.clock_label)
        
        # 時計更新タイマー
        self.clock_timer = QTimer()
        self.clock_timer.timeout.connect(self._update_clock)
        self.clock_timer.start(1000)
        self._update_clock()
        
        # Excelファイル読み込み＆シート選択コンテナ
        excel_container = QWidget()
        excel_container.setStyleSheet("""
            QWidget {
                background-color: #2a2a3a;
                border-radius: 8px;
            }
        """)
        excel_layout = QHBoxLayout(excel_container)
        excel_layout.setContentsMargins(8, 4, 8, 4)
        excel_layout.setSpacing(6)
        
        # フォルダアイコンボタン
        load_btn = QPushButton()
        load_btn.setFixedSize(28, 28)
        load_btn.setToolTip("Load Excel")
        load_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        folder_icon = get_icon_from_base64("folder")
        if not folder_icon.isNull():
            load_btn.setIcon(folder_icon)
            load_btn.setIconSize(QSize(18, 18))
        else:
            load_btn.setText("📁")
        load_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #3a3a4a;
                border-radius: 4px;
            }
        """)
        load_btn.clicked.connect(self.load_excel)
        excel_layout.addWidget(load_btn)
        
        # シートセレクター
        self.sheet_selector = QComboBox()
        self.sheet_selector.setFixedHeight(24)
        self.sheet_selector.setMinimumWidth(120)
        self.sheet_selector.setMaximumWidth(250)
        self.sheet_selector.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.sheet_selector.setStyleSheet("""
            QComboBox {
                background-color: transparent;
                color: #ffffff;
                border: none;
                padding: 2px 4px;
                font-size: 14px;
                font-weight: bold;
            }
            QComboBox::drop-down {
                border: none;
                width: 16px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid #808080;
                margin-right: 4px;
            }
            QComboBox QAbstractItemView {
                background-color: #2a2a3a;
                color: #ffffff;
                selection-background-color: #3a3a4a;
                border: 1px solid #3a3a4a;
                outline: none;
                font-size: 14px;
                font-weight: bold;
            }
            QComboBox QAbstractItemView::item {
                padding: 6px 10px;
                min-height: 24px;
            }
            QComboBox QAbstractItemView::item:hover {
                background-color: #3a3a4a;
            }
            QComboBox QAbstractItemView QScrollBar:vertical {
                background-color: #2a2a3a;
                width: 8px;
                border: none;
                border-radius: 4px;
            }
            QComboBox QAbstractItemView QScrollBar::handle:vertical {
                background-color: #4a4a5a;
                border-radius: 4px;
                min-height: 20px;
            }
            QComboBox QAbstractItemView QScrollBar::handle:vertical:hover {
                background-color: #5a5a6a;
            }
            QComboBox QAbstractItemView QScrollBar::add-line:vertical,
            QComboBox QAbstractItemView QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QComboBox QAbstractItemView QScrollBar::add-page:vertical,
            QComboBox QAbstractItemView QScrollBar::sub-page:vertical {
                background-color: #2a2a3a;
            }
        """)
        self.sheet_selector.currentIndexChanged.connect(self._on_sheet_changed)
        self.sheet_selector.setPlaceholderText("No file")
        excel_layout.addWidget(self.sheet_selector)
        
        header_layout.addWidget(excel_container)
        
        # 現在のワークブックとシート情報を保持
        self.current_workbook = None
        self.current_sheet_names = []
        
        # リフレッシュボタン（シンプルなアイコン）
        refresh_btn = QPushButton()
        refresh_btn.setFixedSize(36, 36)
        refresh_btn.setToolTip("Refresh")
        refresh_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        reload_icon = get_icon_from_base64("reload")
        if not reload_icon.isNull():
            refresh_btn.setIcon(reload_icon)
            refresh_btn.setIconSize(QSize(20, 20))
        else:
            refresh_btn.setText("↻")
        refresh_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: none;
                border-radius: 8px;
                font-size: 22px;
                font-weight: bold;
                color: #4ade80;
            }
            QPushButton:hover {
                background-color: #2a2a3a;
            }
        """)
        refresh_btn.clicked.connect(self.refresh_tasks)
        header_layout.addWidget(refresh_btn)
        
        layout.addLayout(header_layout)
        
        # タスクテーブル
        self.task_table = QTableWidget(0, 8)
        self.task_table.setHorizontalHeaderLabels(["", "Profile", "Site", "Mode", "URL", "Proxy", "Status", "Action"])
        
        # ヘッダーにチェックボックスを追加
        self.header_checkbox = QCheckBox()
        self.header_checkbox.setStyleSheet("margin-left: 12px;")
        self.header_checkbox.stateChanged.connect(self.toggle_all_checkboxes)
        
        header_widget = QWidget()
        header_layout_cb = QHBoxLayout(header_widget)
        header_layout_cb.addWidget(self.header_checkbox)
        header_layout_cb.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout_cb.setContentsMargins(0, 0, 0, 0)
        self.task_table.setCellWidget(0, 0, header_widget)  # これは後でヘッダーに設定
        
        # カラム幅の設定
        header_view = self.task_table.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)  # チェックボックス
        header_view.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)  # Profile
        header_view.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)  # Site
        header_view.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)  # Mode
        header_view.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)  # URL
        header_view.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)  # Proxy
        header_view.setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)  # Status
        header_view.setSectionResizeMode(7, QHeaderView.ResizeMode.Fixed)  # Action
        
        self.task_table.setColumnWidth(0, 40)   # チェックボックス
        self.task_table.setColumnWidth(1, 80)   # Profile
        self.task_table.setColumnWidth(2, 80)   # Site
        self.task_table.setColumnWidth(3, 80)   # Mode
        self.task_table.setColumnWidth(5, 120)  # Proxy
        self.task_table.setColumnWidth(6, 160)  # Status
        self.task_table.setColumnWidth(7, 80)   # Action
        
        self.task_table.setStyleSheet(self._table_style())
        self.task_table.setShowGrid(False)
        self.task_table.verticalHeader().setVisible(False)
        self.task_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.task_table.setSelectionMode(QTableWidget.SelectionMode.NoSelection)
        self.task_table.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        
        # 右クリックメニューを有効化
        self.task_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.task_table.customContextMenuRequested.connect(self._show_task_context_menu)
        
        # ヘッダーの配置を設定
        header = self.task_table.horizontalHeader()
        header.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        
        # ホバーエフェクト用の設定
        self.task_table.setMouseTracking(True)
        self.task_table.viewport().setMouseTracking(True)
        self.task_table.viewport().installEventFilter(self)
        self.hovered_row = -1
        
        # ヘッダーの最初の列にチェックボックスを配置するためのカスタムヘッダー
        self._setup_header_checkbox()
        
        # StatusとActionヘッダーを中央揃えに
        header_item_status = QTableWidgetItem("Status")
        header_item_status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.task_table.setHorizontalHeaderItem(6, header_item_status)
        
        header_item_action = QTableWidgetItem("Action")
        header_item_action.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        self.task_table.setHorizontalHeaderItem(7, header_item_action)
        
        layout.addWidget(self.task_table)
        
        # 下部レイアウト
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(8)
        
        # ステータスカウンター（左側）- ボタンと同じスタイル
        self.idle_count = 0
        self.success_count = 0
        self.failed_count = 0
        
        self.idle_label = QLabel("Idle Task: 0")
        self.idle_label.setStyleSheet(self._status_label_style("#b0b0b0"))
        bottom_layout.addWidget(self.idle_label)
        
        self.success_label = QLabel("Success Task: 0")
        self.success_label.setStyleSheet(self._status_label_style("#2ecc71"))
        bottom_layout.addWidget(self.success_label)
        
        self.failed_label = QLabel("Failed Task: 0")
        self.failed_label.setStyleSheet(self._status_label_style("#e74c3c"))
        bottom_layout.addWidget(self.failed_label)
        
        # 検索ボックス
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search Task...")
        self.search_box.setFixedWidth(170)
        self.search_box.setStyleSheet("""
            QLineEdit {
                background-color: #3c3c3c;
                color: #ffffff;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 4px 25px 4px 8px;
                font-size: 12px;
                margin-left: 10px;
            }
            QLineEdit:focus {
                border: 1px solid #3498db;
            }
        """)
        self.search_box.textChanged.connect(self.filter_tasks)
        self.search_box.textChanged.connect(self._update_search_clear_button)
        
        # クリアボタン
        self.search_clear_btn = QPushButton("✕", self.search_box)
        self.search_clear_btn.setFixedSize(18, 18)
        self.search_clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.search_clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #555;
                color: #fff;
                border: none;
                border-radius: 9px;
                font-size: 10px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #777;
            }
        """)
        self.search_clear_btn.clicked.connect(lambda: self.search_box.clear())
        self.search_clear_btn.hide()
        
        # クリアボタンの位置を調整
        self.search_box.resizeEvent = lambda e: self.search_clear_btn.move(
            self.search_box.width() - 23, 
            (self.search_box.height() - self.search_clear_btn.height()) // 2
        )
        
        bottom_layout.addWidget(self.search_box)
        
        # ステータスフィルターボタン
        self.status_filter_btn = QPushButton("⊟")
        self.status_filter_btn.setFixedSize(26, 26)
        self.status_filter_btn.setToolTip("Filter by Status")
        self.status_filter_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.status_filter_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                border: 1px solid #555;
                border-radius: 4px;
                font-size: 11px;
                color: #b0b0b0;
            }
            QPushButton:hover {
                background-color: #3a3a4a;
                border-color: #666;
            }
        """)
        self.status_filter_btn.clicked.connect(self._show_status_filter_menu)
        bottom_layout.addWidget(self.status_filter_btn)
        
        # 現在のステータスフィルター（None = 全て表示）
        self.current_status_filter = None
        
        bottom_layout.addStretch()
        
        # Start All / Stop Allボタン（右側）
        start_all_btn = QPushButton("▶ Start All")
        start_all_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        start_all_btn.setStyleSheet(self._outline_button_style("#27ae60"))
        start_all_btn.clicked.connect(self.start_all_tasks)
        bottom_layout.addWidget(start_all_btn)
        
        stop_all_btn = QPushButton("■ Stop All")
        stop_all_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        stop_all_btn.setStyleSheet(self._outline_button_style("#c0392b"))
        stop_all_btn.clicked.connect(self.stop_all_tasks)
        bottom_layout.addWidget(stop_all_btn)
        
        layout.addLayout(bottom_layout)
    
    def _setup_header_checkbox(self):
        """ヘッダーの最初の列にチェックボックスを設置"""
        # カスタムヘッダーウィジェットを作成
        header = self.task_table.horizontalHeader()
        
        # チェックボックス用のコンテナ
        self.header_checkbox_container = QWidget(self.task_table)
        self.header_checkbox_container.setStyleSheet("background-color: transparent;")
        layout = QHBoxLayout(self.header_checkbox_container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.header_checkbox = CheckmarkCheckBox()
        self.header_checkbox.stateChanged.connect(self.toggle_all_checkboxes)
        layout.addWidget(self.header_checkbox)
        
        # ヘッダーの位置にチェックボックスを配置
        self._update_header_checkbox_position()
        header.sectionResized.connect(self._update_header_checkbox_position)
        header.sectionMoved.connect(self._update_header_checkbox_position)
    
    def _update_header_checkbox_position(self):
        """ヘッダーチェックボックスの位置を更新"""
        header = self.task_table.horizontalHeader()
        self.header_checkbox_container.setGeometry(
            header.sectionPosition(0),
            0,
            header.sectionSize(0),
            header.height()
        )
        self.header_checkbox_container.show()
    
    def toggle_all_checkboxes(self, state):
        """全てのチェックボックスをトグル（表示中のタスクのみ）"""
        checked = state == Qt.CheckState.Checked.value
        for row in range(self.task_table.rowCount()):
            # 非表示の行はスキップ
            if self.task_table.isRowHidden(row):
                continue
            checkbox_widget = self.task_table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox = checkbox_widget.findChild(QCheckBox)
                if checkbox:
                    checkbox.setChecked(checked)
    
    def filter_tasks(self, search_text):
        """検索テキストに基づいてタスクをフィルタリング"""
        search_text = search_text.lower().strip()
        visible_count = 0
        
        for row in range(self.task_table.rowCount()):
            show_row = True
            
            # 検索テキストフィルター
            if search_text:
                match_found = False
                for col in range(self.task_table.columnCount()):
                    # ウィジェットからテキストを取得
                    widget = self.task_table.cellWidget(row, col)
                    if widget:
                        label = widget.findChild(QLabel)
                        if label:
                            cell_text = label.text().lower()
                            if search_text in cell_text:
                                match_found = True
                                break
                if not match_found:
                    show_row = False
            
            # ステータスフィルター
            if show_row and self.current_status_filter:
                status_widget = self.task_table.cellWidget(row, 6)
                if status_widget:
                    status_label = status_widget.findChild(QLabel, "status_label")
                    if status_label:
                        current_status = status_label.text()
                        if self.current_status_filter == "Idle":
                            if current_status != "Idle":
                                show_row = False
                        elif self.current_status_filter == "Success":
                            if current_status != "Success":
                                show_row = False
                        elif self.current_status_filter == "Failed":
                            # Failed, Stopped, Error, Not supported, Failed ○○ を含める
                            is_failed = (
                                current_status in ["Failed", "Stopped", "Error", "Not supported"] or
                                current_status.startswith("Error:") or
                                current_status.startswith("Failed")
                            )
                            if not is_failed:
                                show_row = False
            
            self.task_table.setRowHidden(row, not show_row)
            if show_row:
                visible_count += 1
        
        # Tasks数を更新
        self.update_task_count()
    
    def _update_search_clear_button(self, text):
        """検索ボックスのクリアボタンの表示/非表示を更新"""
        if text:
            self.search_clear_btn.show()
        else:
            self.search_clear_btn.hide()
    
    def _show_status_filter_menu(self):
        """ステータスフィルターメニューを表示"""
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: #1e1e2e;
                border: 1px solid #3a3a4a;
                border-radius: 6px;
                padding: 4px;
            }
            QMenu::item {
                background-color: transparent;
                color: #ffffff;
                padding: 6px 20px;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
            }
            QMenu::item:selected {
                background-color: #2a2a3a;
            }
            QMenu::item:checked {
                background-color: #27ae60;
            }
            QMenu::separator {
                height: 1px;
                background: #3a3a4a;
                margin: 4px 8px;
            }
        """)
        
        # All（フィルターなし）
        all_action = menu.addAction("All")
        all_action.setCheckable(True)
        all_action.setChecked(self.current_status_filter is None)
        
        menu.addSeparator()
        
        # Idle
        idle_action = menu.addAction("Idle")
        idle_action.setCheckable(True)
        idle_action.setChecked(self.current_status_filter == "Idle")
        
        # Success
        success_action = menu.addAction("Success")
        success_action.setCheckable(True)
        success_action.setChecked(self.current_status_filter == "Success")
        
        # Failed
        failed_action = menu.addAction("Failed")
        failed_action.setCheckable(True)
        failed_action.setChecked(self.current_status_filter == "Failed")
        
        # メニューのサイズを取得してボタンの上に表示
        menu_height = menu.sizeHint().height()
        btn_pos = self.status_filter_btn.mapToGlobal(self.status_filter_btn.rect().topLeft())
        menu_pos = btn_pos - QPoint(0, menu_height)
        
        # メニューを表示
        action = menu.exec(menu_pos)
        
        if action == all_action:
            self._apply_status_filter(None)
        elif action == idle_action:
            self._apply_status_filter("Idle")
        elif action == success_action:
            self._apply_status_filter("Success")
        elif action == failed_action:
            self._apply_status_filter("Failed")
    
    def _apply_status_filter(self, status):
        """ステータスフィルターを適用"""
        self.current_status_filter = status
        
        # ボタンの見た目を更新（ステータスに応じて色を変更）
        if status:
            # ステータスに応じた色を設定
            if status == "Idle":
                bg_color = "#808080"
                hover_color = "#909090"
            elif status == "Success":
                bg_color = "#2ecc71"
                hover_color = "#27ae60"
            elif status == "Failed":
                bg_color = "#e74c3c"
                hover_color = "#c0392b"
            else:
                bg_color = "#27ae60"
                hover_color = "#2ecc71"
            
            self.status_filter_btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {bg_color};
                    border: 1px solid {bg_color};
                    border-radius: 4px;
                    font-size: 11px;
                    color: #ffffff;
                }}
                QPushButton:hover {{
                    background-color: {hover_color};
                    border-color: {hover_color};
                }}
            """)
            self.status_filter_btn.setToolTip(f"Filter: {status}")
        else:
            self.status_filter_btn.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    border: 1px solid #555;
                    border-radius: 4px;
                    font-size: 11px;
                    color: #b0b0b0;
                }
                QPushButton:hover {
                    background-color: #3a3a4a;
                    border-color: #666;
                }
            """)
            self.status_filter_btn.setToolTip("Filter by Status")
        
        # フィルターを再適用
        self.filter_tasks(self.search_box.text())
    
    def update_task_count(self):
        """表示中のタスク数を更新"""
        visible_count = 0
        for row in range(self.task_table.rowCount()):
            if not self.task_table.isRowHidden(row):
                visible_count += 1
        self.header_label.setText(f"Tasks({visible_count})")
    
    def _update_clock(self):
        """時計を更新"""
        from datetime import datetime
        now = datetime.now()
        self.clock_label.setText(now.strftime("%Y / %m / %d  %H : %M : %S"))
    
    def refresh_tasks(self):
        """タスクリストをリフレッシュ"""
        # 実行中のタスクを停止
        for row in list(self.workers.keys()):
            if self.workers[row].isRunning():
                self.workers[row].stop()
        self.workers.clear()
        
        # Excelを再読み込み（現在選択中のシートを維持）
        if self.current_excel_path and os.path.exists(self.current_excel_path):
            current_sheet = self.sheet_selector.currentText() if self.sheet_selector.isVisible() else None
            self._load_excel_file(self.current_excel_path, show_reload_toast=True, sheet_name=current_sheet)
        else:
            self.toast.show_toast("No file to reload", "warning", 2000)
    
    def load_excel(self):
        """Excel/CSVファイルを読み込んでテーブルに表示"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "ファイルを選択", "", "Excel/CSV Files (*.xlsx *.xls *.csv);;Excel Files (*.xlsx *.xls);;CSV Files (*.csv)"
        )
        
        if not file_path:
            return
        
        self._load_excel_file(file_path)
    
    def _load_excel_file(self, file_path, show_reload_toast=False, show_toast=True, sheet_name=None):
        """指定されたExcel/CSVファイルを読み込む
        
        Args:
            file_path: Excel/CSVファイルのパス
            show_reload_toast: リロード時のトースト表示フラグ
            show_toast: トースト表示フラグ
            sheet_name: 読み込むシート名（Noneの場合は最初のシート、CSVでは無視）
        """
        EXCEL_COLUMNS = [
            "Profile", "Site", "Mode", "URL", "Proxy",
            "Loginid", "Loginpass", "LastName", "FirstName",
            "LastNameKana", "FirstNameKana", "Country", "State",
            "City", "Address1", "Address2", "Zipcode", "Tell",
            "Mail", "Birthday", "Size", "Gender", "Cardfirstname", "Cardlastname",
            "Cardnumber", "Cardmonth", "Cardyear", "Securitycode",
            "Free1", "Free2"
        ]
        
        try:
            # CSVファイルかどうか判定
            is_csv = file_path.lower().endswith('.csv')
            
            if is_csv:
                # CSV読み込み
                import csv
                
                # シートセレクターを無効化（CSVにはシートがない）
                self.current_sheet_names = []
                self.sheet_selector.blockSignals(True)
                self.sheet_selector.clear()
                self.sheet_selector.addItem("(CSV)")
                self.sheet_selector.setEnabled(False)
                self.sheet_selector.blockSignals(False)
                
                rows_data = []
                # エンコーディングを試行（UTF-8 → Shift-JIS → CP932）
                encodings = ['utf-8', 'utf-8-sig', 'shift-jis', 'cp932']
                for encoding in encodings:
                    try:
                        with open(file_path, 'r', encoding=encoding, newline='') as f:
                            reader = csv.reader(f)
                            rows_data = list(reader)
                        break
                    except UnicodeDecodeError:
                        continue
                
                if not rows_data:
                    raise Exception("ファイルのエンコーディングを判定できませんでした")
                
                self.task_table.setRowCount(0)
                self.workers.clear()
                self.all_task_data = []
                self.original_task_data = []
                self.last_checked_row = None
                
                # カウンターをリセット
                self.idle_count = 0
                self.success_count = 0
                self.failed_count = 0
                
                task_count = 0
                
                # ヘッダー行をスキップ（2行目から読み込み）
                for row_idx, row in enumerate(rows_data[1:], start=0):
                    if not row or row[0] is None or row[0] == '':
                        break
                    
                    task_data = {}
                    for col_idx, col_name in enumerate(EXCEL_COLUMNS):
                        if col_idx < len(row):
                            task_data[col_name] = str(row[col_idx]) if row[col_idx] else ""
                        else:
                            task_data[col_name] = ""
                    
                    self.all_task_data.append(task_data)
                    self.original_task_data.append(task_data.copy())
                    
                    profile = task_data["Profile"]
                    site = task_data["Site"]
                    mode = task_data["Mode"]
                    url = task_data["URL"]
                    proxy = task_data["Proxy"]
                    
                    self.add_task_row(profile, site, mode, url, proxy)
                    task_count += 1
                
            else:
                # Excel読み込み
                wb = openpyxl.load_workbook(file_path)
                
                # シート名リストを取得
                sheet_names = wb.sheetnames
                
                # シートセレクターを有効化・更新（シート名指定がない場合のみ）
                self.sheet_selector.setEnabled(True)
                if sheet_name is None:
                    self.current_sheet_names = sheet_names
                    self.sheet_selector.blockSignals(True)
                    self.sheet_selector.clear()
                    self.sheet_selector.addItems(sheet_names)
                    self.sheet_selector.blockSignals(False)
                
                # 指定されたシート、またはアクティブシートを選択
                if sheet_name and sheet_name in sheet_names:
                    ws = wb[sheet_name]
                else:
                    ws = wb.active
                    # セレクターの選択を同期（blockSignalsで無限ループ防止）
                    if ws.title in sheet_names:
                        self.sheet_selector.blockSignals(True)
                        self.sheet_selector.setCurrentText(ws.title)
                        self.sheet_selector.blockSignals(False)
                
                self.task_table.setRowCount(0)
                self.workers.clear()
                self.all_task_data = []
                self.original_task_data = []
                self.last_checked_row = None
                
                # カウンターをリセット
                self.idle_count = 0
                self.success_count = 0
                self.failed_count = 0
                
                task_count = 0
                
                for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=0):
                    if row[0] is None:
                        break
                    
                    task_data = {}
                    for col_idx, col_name in enumerate(EXCEL_COLUMNS):
                        if col_idx < len(row):
                            task_data[col_name] = str(row[col_idx]) if row[col_idx] else ""
                        else:
                            task_data[col_name] = ""
                    
                    self.all_task_data.append(task_data)
                    self.original_task_data.append(task_data.copy())
                    
                    profile = task_data["Profile"]
                    site = task_data["Site"]
                    mode = task_data["Mode"]
                    url = task_data["URL"]
                    proxy = task_data["Proxy"]
                    
                    self.add_task_row(profile, site, mode, url, proxy)
                    task_count += 1
                
                wb.close()
            
            self.header_label.setText(f"Tasks({task_count})")
            self.current_excel_path = file_path
            self._save_last_excel(file_path)
            
            # Idleカウントを更新
            self.idle_count = task_count
            self.idle_label.setText(f"Idle Task: {self.idle_count}")
            self.success_label.setText(f"Success Task: {self.success_count}")
            self.failed_label.setText(f"Failed Task: {self.failed_count}")
            
            # トースト通知
            if show_toast:
                if show_reload_toast:
                    self.toast.show_toast("Successfully Reload Tasks", "success", 2000)
                else:
                    self.toast.show_toast("Successfully Loaded File", "success", 2000)
            
        except Exception as e:
            if show_toast:
                self.toast.show_toast(f"Failed to load file: {str(e)[:30]}", "error", 3000)
    
    def _on_sheet_changed(self, index):
        """シート選択が変更されたときの処理"""
        if index < 0 or not self.current_excel_path:
            return
        
        sheet_name = self.sheet_selector.currentText()
        if sheet_name:
            self._load_excel_file(self.current_excel_path, show_toast=True, sheet_name=sheet_name)
    
    def add_task_row(self, profile, site, mode, url, proxy):
        """タスク行を追加"""
        row = self.task_table.rowCount()
        self.task_table.insertRow(row)
        
        # チェックボックス（カスタムクラス使用）
        checkbox = CheckmarkCheckBox()
        checkbox.clicked.connect(lambda checked, r=row: self._on_checkbox_clicked(r, checked))
        checkbox_widget = QWidget()
        checkbox_widget.setStyleSheet("background-color: transparent;")
        checkbox_layout = QHBoxLayout(checkbox_widget)
        checkbox_layout.addWidget(checkbox)
        checkbox_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        checkbox_layout.setContentsMargins(0, 0, 0, 0)
        
        # チェックボックスにマウストラッキングを追加
        checkbox_widget.setMouseTracking(True)
        checkbox_widget.enterEvent = lambda event, r=row: self._update_row_hover(r)
        checkbox.setMouseTracking(True)
        checkbox.enterEvent = lambda event, r=row: self._update_row_hover(r)
        
        self.task_table.setCellWidget(row, 0, checkbox_widget)
        
        # テキストセル用のウィジェット作成関数
        def create_text_cell(text, align_center=False, tooltip=None):
            widget = QWidget()
            widget.setStyleSheet("background-color: transparent; margin: 0px; padding: 0px;")
            widget.setContentsMargins(0, 0, 0, 0)
            layout = QHBoxLayout(widget)
            layout.setContentsMargins(8, 0, 8, 0)
            layout.setSpacing(0)
            label = QLabel(text)
            label.setStyleSheet("color: #ffffff; background-color: transparent; margin: 0px; padding: 0px;")
            if align_center:
                layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            else:
                layout.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            if tooltip:
                label.setToolTip(tooltip)
            layout.addWidget(label)
            widget.setMouseTracking(True)
            widget.enterEvent = lambda event, r=row: self._update_row_hover(r)
            return widget
        
        # Profile
        self.task_table.setCellWidget(row, 1, create_text_cell(profile))
        
        # Site
        self.task_table.setCellWidget(row, 2, create_text_cell(site))
        
        # Mode
        self.task_table.setCellWidget(row, 3, create_text_cell(mode))
        
        # URL
        self.task_table.setCellWidget(row, 4, create_text_cell(url, tooltip=url))
        
        # Proxy
        self.task_table.setCellWidget(row, 5, create_text_cell(proxy, tooltip=proxy))
        
        # Status
        status_widget = QWidget()
        status_widget.setStyleSheet("background-color: transparent; margin: 0px; padding: 0px;")
        status_widget.setContentsMargins(0, 0, 0, 0)
        status_layout = QHBoxLayout(status_widget)
        status_layout.setContentsMargins(8, 0, 8, 0)
        status_layout.setSpacing(0)
        status_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_label = QLabel("Idle")
        status_label.setStyleSheet("color: #b0b0b0; background-color: transparent; margin: 0px; padding: 0px;")
        status_label.setObjectName("status_label")
        status_layout.addWidget(status_label)
        status_widget.setMouseTracking(True)
        status_widget.enterEvent = lambda event, r=row: self._update_row_hover(r)
        self.task_table.setCellWidget(row, 6, status_widget)
        
        # Action buttons
        action_widget = QWidget()
        action_widget.setStyleSheet("background-color: transparent;")
        action_layout = QHBoxLayout(action_widget)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(8)
        action_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        start_btn = QPushButton("▶")
        start_btn.setFixedSize(28, 28)
        start_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        start_btn.setStyleSheet(self._action_icon_style("#27ae60"))
        start_btn.clicked.connect(lambda checked, r=row: self.start_task(r))
        start_btn.setMouseTracking(True)
        start_btn.enterEvent = lambda event, r=row: self._update_row_hover(r)
        action_layout.addWidget(start_btn)
        
        stop_btn = QPushButton("■")
        stop_btn.setFixedSize(28, 28)
        stop_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        stop_btn.setStyleSheet(self._action_icon_style("#e74c3c"))
        stop_btn.clicked.connect(lambda checked, r=row: self.stop_task(r))
        stop_btn.setMouseTracking(True)
        stop_btn.enterEvent = lambda event, r=row: self._update_row_hover(r)
        action_layout.addWidget(stop_btn)
        
        # action_widget全体にもマウストラッキング
        action_widget.setMouseTracking(True)
        action_widget.enterEvent = lambda event, r=row: self._update_row_hover(r)
        
        self.task_table.setCellWidget(row, 7, action_widget)
        
        # 行の高さを設定
        self.task_table.setRowHeight(row, 45)
    
    def _on_checkbox_clicked(self, row, checked):
        """チェックボックスがクリックされた時の処理(シフトクリック対応)"""
        modifiers = QApplication.keyboardModifiers()
        
        if modifiers == Qt.KeyboardModifier.ShiftModifier and self.last_checked_row is not None:
            start_row = min(self.last_checked_row, row)
            end_row = max(self.last_checked_row, row)
            
            for r in range(start_row, end_row + 1):
                checkbox_widget = self.task_table.cellWidget(r, 0)
                if checkbox_widget:
                    cb = checkbox_widget.findChild(QCheckBox)
                    if cb:
                        cb.setChecked(checked)
        
        self.last_checked_row = row
    
    def start_task(self, row):
        """指定行のタスクを開始"""
        if row in self.workers and self.workers[row].isRunning():
            return
        
        # 古いfinished_workerがあれば削除
        if hasattr(self, 'finished_workers') and row in self.finished_workers:
            del self.finished_workers[row]
        
        # 古いworkerがあれば削除（実行中でない場合）
        if row in self.workers:
            old_worker = self.workers[row]
            try:
                old_worker.status_changed.disconnect(self.update_status)
            except:
                pass
            try:
                old_worker.result_changed.disconnect(self.update_result)
            except:
                pass
            try:
                old_worker.finished_task.disconnect(self.on_task_finished)
            except:
                pass
            try:
                old_worker.raffle_result.disconnect(self.on_raffle_result)
            except:
                pass
            del self.workers[row]
        
        if row >= len(self.all_task_data):
            self.update_status(row, "Failed")
            return
        
        task_data = self.all_task_data[row].copy()
        
        site = task_data.get("Site", "").lower()
        mode = task_data.get("Mode", "").lower()
        
        if not site or not mode:
            self.update_status(row, "Failed")
            return
        
        # サポートされているサイトとモードをチェック
        supported_sites = ["amazon", "icloud", "x", "rakuten"]
        supported_modes = {
            "amazon": ["browser", "signup", "addy", "card", "raffle"],
            "icloud": ["generate", "collect"],
            "x": ["repost", "browser", "password", "follow", "name", "bio", "icon", "header", "mail", "follower"],
            "rakuten": ["browser", "address", "card", "name", "raffle"]
        }
        
        if site not in supported_sites:
            self.update_status(row, "Not supported")
            return
        
        if mode not in supported_modes.get(site, []):
            self.update_status(row, "Not supported")
            return
        
        # プロキシ設定を適用
        if self.proxy_page:
            random_proxy = self.proxy_page.get_random_proxy()
            if random_proxy:
                task_data["Proxy"] = random_proxy
                # all_task_dataも更新
                if row < len(self.all_task_data):
                    self.all_task_data[row]["Proxy"] = random_proxy
                # テーブルのProxy列も更新
                self._update_proxy_cell(row, random_proxy)
                print(f"[Proxy] Using random proxy: {random_proxy}")
            else:
                # ProxyがOFF（またはグループ未選択）の場合、元のExcelデータを使用
                if row < len(self.original_task_data):
                    original_proxy = self.original_task_data[row].get("Proxy", "")
                    task_data["Proxy"] = original_proxy
                    if row < len(self.all_task_data):
                        self.all_task_data[row]["Proxy"] = original_proxy
                    # テーブルのProxy列も更新
                    self._update_proxy_cell(row, original_proxy)
                    if original_proxy:
                        print(f"[Proxy] Using Excel proxy: {original_proxy}")
                    else:
                        print(f"[Proxy] Using local (no proxy)")
        
        # Headless設定を追加
        task_data["Headless"] = self.headless_checkbox.isChecked()
        if task_data["Headless"]:
            print(f"[Headless] Running in headless mode")
        
        # サーバーからBotをダウンロードする方式に変更
        # bot_pathは不要（サーバーから取得するため）
        
        worker = BotWorker(row, task_data)
        worker.status_changed.connect(self.update_status)
        worker.result_changed.connect(self.update_result)
        worker.finished_task.connect(self.on_task_finished)
        worker.raffle_result.connect(self.on_raffle_result)  # Raffle結果用
        
        self.workers[row] = worker
        worker.start()
    
    def stop_task(self, row):
        """指定行のタスクを停止"""
        if row in self.workers and self.workers[row].isRunning():
            self.workers[row].stop()
    
    def start_all_tasks(self):
        """チェックされた全タスクを開始（並列数制限付き、表示中のタスクのみ）"""
        # 設定を読み込み
        parallel_count = 3  # デフォルト値
        try:
            if self.GENERAL_SETTINGS_FILE.exists():
                with open(self.GENERAL_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                parallel_count = settings.get("parallel_count", 3)
        except:
            pass
        
        # チェックされたタスクを収集（表示中のタスクのみ）
        checked_rows = []
        for row in range(self.task_table.rowCount()):
            # 非表示の行はスキップ
            if self.task_table.isRowHidden(row):
                continue
            checkbox_widget = self.task_table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox = checkbox_widget.findChild(QCheckBox)
                if checkbox and checkbox.isChecked():
                    # 既に実行中でないタスクのみ追加
                    if row not in self.workers or not self.workers[row].isRunning():
                        checked_rows.append(row)
        
        # 現在の実行中タスク数を確認
        running_count = len([w for w in self.workers.values() if w.isRunning()])
        available_slots = parallel_count - running_count
        
        # 即時実行するタスク
        immediate_tasks = checked_rows[:available_slots]
        # 待機キューに入れるタスク
        self.pending_tasks = checked_rows[available_slots:]
        
        # 待機キューのタスクのステータスを「Waiting task...」に設定
        for row in self.pending_tasks:
            self.update_status(row, "Waiting task...")
        
        # 即時実行（Thread数分は同時に開始）
        for row in immediate_tasks:
            self.start_task(row)
        
        if self.pending_tasks:
            print(f"Queued {len(self.pending_tasks)} tasks (parallel limit: {parallel_count})")
    
    def _check_pending_tasks(self):
        """待機中のタスクを確認して実行（Task Delay適用）"""
        if not self.pending_tasks:
            return
        
        # 設定を読み込み
        parallel_count = 3
        task_delay = 0
        try:
            if self.GENERAL_SETTINGS_FILE.exists():
                with open(self.GENERAL_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                parallel_count = settings.get("parallel_count", 3)
                task_delay = settings.get("task_delay", 0)
        except:
            pass
        
        # 現在の実行中タスク数を確認
        running_count = len([w for w in self.workers.values() if w.isRunning()])
        available_slots = parallel_count - running_count
        
        # 空きスロットがあれば次のタスクを実行（Task Delay後）
        if available_slots > 0 and self.pending_tasks:
            row = self.pending_tasks.pop(0)
            if task_delay > 0:
                print(f"Task Delay: waiting {task_delay}s before starting task {row + 1}")
                QTimer.singleShot(task_delay * 1000, lambda r=row: self._delayed_start_task(r))
            else:
                self.start_task(row)
    
    def _delayed_start_task(self, row):
        """遅延後にタスクを開始"""
        self.start_task(row)
        # 次の待機タスクもチェック
        self._check_pending_tasks()
    
    def stop_all_tasks(self):
        """全タスクを停止（実行中 + 待機中）"""
        # 1. 待機キューをクリアし、ステータスを「Stopped」に
        for row in self.pending_tasks:
            # 非表示の行はスキップ
            if self.task_table.isRowHidden(row):
                continue
            self.update_status(row, "Stopped")
        self.pending_tasks = []
        
        # 2. 実行中のタスクを停止
        for row in list(self.workers.keys()):
            # 非表示の行はスキップ
            if self.task_table.isRowHidden(row):
                continue
            self.stop_task(row)
    
    def _show_task_context_menu(self, pos):
        """タスクの右クリックメニューを表示"""
        # クリックされた行を取得
        row = self.task_table.rowAt(pos.y())
        if row < 0:
            return
        
        # メインメニュー
        menu = QMenu(self)
        menu.setStyleSheet(self._context_menu_style())
        
        # Logサブメニュー
        log_menu = QMenu("Log", menu)
        log_menu.setStyleSheet(self._context_menu_style())
        
        view_action = log_menu.addAction("View")
        copy_action = log_menu.addAction("Copy")
        
        menu.addMenu(log_menu)
        
        # メニューを表示
        action = menu.exec(self.task_table.viewport().mapToGlobal(pos))
        
        if action == view_action:
            self._view_task_log(row)
        elif action == copy_action:
            self._copy_task_log(row)
    
    def _context_menu_style(self):
        """コンテキストメニューのスタイル"""
        return """
            QMenu {
                background-color: #1e1e2e;
                border: 1px solid #3a3a4a;
                border-radius: 8px;
                padding: 6px;
            }
            QMenu::item {
                background-color: transparent;
                color: #ffffff;
                padding: 10px 28px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: bold;
            }
            QMenu::item:selected {
                background-color: #2a2a3a;
            }
            QMenu::separator {
                height: 1px;
                background: #3a3a4a;
                margin: 6px 10px;
            }
        """
    
    def _view_task_log(self, row):
        """タスクのログを表示"""
        log_text = self._get_task_log(row)
        
        # Profileの値を取得してタイトルに使用
        profile = self._get_cell_text(row, 1)  # Profileは1番目のカラム
        if not profile:
            profile = str(row + 1)
        
        # ログビューアダイアログを表示
        dialog = QDialog(self)
        dialog.setWindowTitle(f"Task {profile} Log")
        dialog.setMinimumSize(700, 500)
        dialog.setStyleSheet("""
            QDialog {
                background-color: #1e1e2e;
            }
        """)
        
        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        # ログ表示エリア
        log_view = QPlainTextEdit()
        log_view.setPlainText(log_text if log_text else "No log available")
        log_view.setReadOnly(True)
        log_view.setStyleSheet("""
            QPlainTextEdit {
                background-color: #0d0d0d;
                color: #00ff00;
                font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
                font-size: 12px;
                border: 1px solid #333333;
                border-radius: 6px;
                padding: 12px;
            }
        """)
        layout.addWidget(log_view)
        
        # ボタンエリア
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        copy_btn = QPushButton("Copy")
        copy_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        copy_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        copy_btn.clicked.connect(lambda: self._copy_log_from_dialog(log_text))
        btn_layout.addWidget(copy_btn)
        
        close_btn = QPushButton("Close")
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #404040;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #505050;
            }
        """)
        close_btn.clicked.connect(dialog.close)
        btn_layout.addWidget(close_btn)
        
        layout.addLayout(btn_layout)
        
        dialog.exec()
    
    def _copy_task_log(self, row):
        """タスクのログをクリップボードにコピー"""
        log_text = self._get_task_log(row)
        
        if log_text:
            clipboard = QApplication.clipboard()
            clipboard.setText(log_text)
            self.toast.show_toast("Log copied to clipboard", "success")
        else:
            self.toast.show_toast("No log available", "warning")
    
    def _copy_log_from_dialog(self, log_text):
        """ダイアログからログをコピー"""
        if log_text:
            clipboard = QApplication.clipboard()
            clipboard.setText(log_text)
            self.toast.show_toast("Log copied to clipboard", "success")
    
    def _get_task_log(self, row):
        """指定された行のタスクのログを取得"""
        # 現在実行中のワーカーからログを取得
        if row in self.workers:
            return self.workers[row].get_log()
        
        # 完了済みワーカーの履歴から取得
        if hasattr(self, 'finished_workers') and row in self.finished_workers:
            return self.finished_workers[row].get_log()
        
        return None
    
    def update_status(self, row, status):
        """ステータスを更新"""
        if row < self.task_table.rowCount():
            # ウィジェットベースのステータス更新
            status_widget = self.task_table.cellWidget(row, 6)
            if status_widget:
                status_label = status_widget.findChild(QLabel, "status_label")
                if status_label:
                    old_status = status_label.text()
                    status_label.setText(status)
                    
                    # 色を設定
                    color = "#b0b0b0"  # デフォルト
                    
                    # Failedで始まるステータスは最優先で赤色
                    if status.startswith("Failed") or status in ["Login Failed", "Repost Failed", "Browse Failed", "Tweet Not Found", "Browser Closed", "Password Change Failed", "No New Password", "Follow Failed", "Already Followed", "Name Change Failed", "No New Name", "Bio Change Failed", "No New Bio", "Icon Change Failed", "No Icon File", "Icon Not Found", "Header Change Failed", "No Header File", "Header Not Found", "Address Change Failed", "No Zipcode", "No Tell", "Card Registration Failed", "No Card Number", "No Card Expiry", "No Name"]:
                        color = "#e74c3c"
                    elif status == "Idle":
                        color = "#b0b0b0"
                    elif status == "Queued":
                        color = "#f39c12"
                    elif status == "Waiting task...":
                        color = "#f39c12"  # オレンジ（待機中）
                    elif status == "Browsing":
                        color = "#3498db"
                    elif status in ["Starting Task", "Checking Login", "Logging In", "Completing"] or status.startswith("Raffle "):
                        color = "#3498db"  # 進捗ステータス（Browsingと同じ青）
                    elif status in [
                        # Card用進捗ステータス（青）- 詳細版
                        "Opening Page", "Checking Existing Card", "Existing Card Found", "Clicking Delete",
                        "Adding New Card", "Clicked Add Payment", "Clicked Add Card", "Waiting Card Form",
                        "Entering Card Number", "Entering Expiry", "Entering Security Code", "Entering Card Name",
                        "Clicked Next", "Clicked Use Address", "Widget Error Retry",
                        "Enable Default Payment", "Clicked Complete", "Card Registration", "Cookies Saved"
                    ]:
                        color = "#3498db"  # Card用進捗ステータス（青）
                    elif status in [
                        # Addy用進捗ステータス（青）- 詳細版
                        "Opening Address", "Checking Address", "Replacing Address", "Adding Address",
                        "Checking Addresses", "Address Found",
                        "Adding New", "Entering Name", "Entering Phone", "Entering Zipcode",
                        "Entering Address", "Selecting State", "Entering City", "Clicking Submit",
                        "Address Added", "Deleting Old", "Setting Default", "Address Replaced", "Saving Cookies"
                    ]:
                        color = "#3498db"  # Addy用進捗ステータス（青）
                    elif status in [
                        # Signup用進捗ステータス（青）- 詳細版
                        "Entering Email", "Clicking Next", "Clicking Create",
                        "Entering Password", "Confirming Password", "Clicking Email",
                        "Waiting CAPTCHA", "Quiz Clicked", "Sending CAPTCHA", "CAPTCHA Solved",
                        "Fetching OTP", "Waiting OTP", "Found OTP", "Entering OTP",
                        "Waiting SMS", "Got SMS", "Entering SMS",
                        "Checking Success", "Account Created",
                        "Navigating", "Clicking Edit",
                        "Confirming Delete", "Deletion Done"
                    ] or status.startswith("CAPTCHA "):
                        color = "#3498db"  # Signup用進捗ステータス（青）
                    elif status == "Navigating":
                        color = "#3498db"  # Browser用進捗ステータス（青）
                    elif status in [
                        # iCloud用進捗ステータス（青）
                        "Clicking Sign In", "Clicking Continue", "Checking 2FA", "Waiting 2FA",
                        "Login Success", "Collecting Emails", "Navigating to HME",
                        "Fetching Emails", "Generating Email", "Waiting Cooldown"
                    ] or status.startswith("Generating ") or status.startswith("Waiting "):
                        color = "#3498db"  # iCloud用進捗ステータス（青）
                    elif status in [
                        # X Repost/Browser/Password/Follow/Name/Bio/Icon/Header/Mail用進捗ステータス（青）
                        "Opening Page", "Following", "Liking", "Reposting", "Verifying",
                        "Browsing", "Solving Cloudflare", "Solving CF Challenge", "OTP Required", "Cloudflare Detected",
                        "Entering Current Password", "Entering New Password", "Confirming Password", "Saving Password",
                        "Checking Follow Status", "Opening Edit Profile", "Entering New Name", "Saving", "Entering New Bio",
                        "Uploading Icon", "Applying Icon", "Uploading Header", "Applying Header",
                        # X Mail用進捗ステータス（青）
                        "Opening Account Page", "Entering Password", "Clicking Confirm", "Clicking Email Link",
                        "Clicking Update Email", "Checking Restrictions", "Entering New Email", "Clicking Next",
                        "Fetching Email OTP", "Entering OTP", "Clicking Verify", "Verifying Change",
                        # X Follower用進捗ステータス（青）
                        "Starting Task", "Checking Login", "Opening Profile", "Getting Followers",
                        # Rakuten用進捗ステータス（青）
                        "Clicking Login", "Entering Email", "Clicking Next", "Entering Password", "Submitting Login", "Verifying Login",
                        "Opening My Rakuten", "Opening Member Info", "Opening Address Page", "Clicking Edit",
                        "Entering Zipcode", "Entering Address", "Entering Tell",
                        "Opening Payment Page", "Checking Existing Card", "Adding New Card", "Entering Card Info",
                        "Entering Card Number", "Entering Expiry Month", "Entering Expiry Year", "Submitting Card",
                        "Re-entering Password",
                        "Opening Personal Info", "Entering LastName", "Entering FirstName", "Entering LastNameKana", "Entering FirstNameKana",
                        # Rakuten Raffle用進捗ステータス（青）
                        "Opening Login Page", "Clicking Entry", "Sending Webhook", "Cookies Saved",
                        # Google用進捗ステータス（青）
                        "Clicking Login", "Entering Email", "Clicking Next", "Checking CAPTCHA", "Waiting CAPTCHA",
                        "Entering Password", "Checking Recovery Email", "Entering Recovery Email",
                        "Checking Phone Verification", "Getting Phone Number", "Entering Phone",
                        "Waiting SMS Code", "Entering SMS Code", "Skipping Prompts"
                    ]:
                        color = "#3498db"  # X用進捗ステータス（青）
                    elif status == "Success" or status.startswith("Success Follower"):
                        color = "#2ecc71"
                        # 成功したらチェックボックスを外す
                        self._uncheck_row(row)
                    elif status in [
                        "Failed", "Stopped", "Not supported", "Already Raffled", "Not Find", "Timeout", "Already Reposted",
                        # X Mail用エラーステータス（赤）
                        "No New Email", "Restricted 48h", "OTP Failed", "Email Change Failed",
                        # X Follower用エラーステータス（赤）
                        "Followers Not Found", "Get Followers Failed", "Parse Failed",
                        # Rakuten Raffle用エラーステータス（赤）
                        "Already Entered", "Not Found",
                        # Google用エラーステータス（赤）
                        "SMS Failed"
                    ]:
                        color = "#e74c3c"
                    elif status.startswith("Raffled("):
                        color = "#e74c3c"
                    elif status.startswith("Failed"):
                        color = "#e74c3c"
                    
                    status_label.setStyleSheet(f"color: {color}; background-color: transparent;")
                    
                    # カウンター更新
                    self._update_status_counter(old_status, status)
    
    def _uncheck_row(self, row):
        """指定行のチェックボックスを外す"""
        try:
            checkbox_widget = self.task_table.cellWidget(row, 0)
            if checkbox_widget:
                checkbox = checkbox_widget.findChild(QCheckBox)
                if checkbox and checkbox.isChecked():
                    checkbox.setChecked(False)
        except Exception as e:
            print(f"Failed to uncheck row {row}: {e}")
    
    def _update_proxy_cell(self, row, proxy_text):
        """Proxy列のテキストを更新（ウィジェットベース対応）"""
        proxy_widget = self.task_table.cellWidget(row, 5)
        if proxy_widget:
            for child in proxy_widget.children():
                if isinstance(child, QLabel):
                    child.setText(proxy_text)
                    child.setToolTip(proxy_text)
                    break
    
    def _get_cell_text(self, row, col):
        """セルのテキストを取得（ウィジェットベース対応）"""
        widget = self.task_table.cellWidget(row, col)
        if widget:
            # QLabel を探す（children()から検索）
            for child in widget.children():
                if isinstance(child, QLabel):
                    text = child.text()
                    return text
        # フォールバック: itemを試す
        item = self.task_table.item(row, col)
        if item:
            return item.text()
        return ""
    
    def _update_status_counter(self, old_status, new_status):
        """ステータスカウンターを更新"""
        # 古いステータスのカウントを減らす
        if old_status == "Idle":
            self.idle_count = max(0, self.idle_count - 1)
        elif old_status == "Success":
            self.success_count = max(0, self.success_count - 1)
        elif old_status in ["Failed", "Stopped", "Not supported"] or old_status.startswith("Failed"):
            self.failed_count = max(0, self.failed_count - 1)
        
        # 新しいステータスのカウントを増やす
        if new_status == "Idle":
            self.idle_count += 1
        elif new_status == "Success":
            self.success_count += 1
        elif new_status in ["Failed", "Stopped", "Not supported"] or new_status.startswith("Failed"):
            self.failed_count += 1
        
        # ラベル更新
        self.idle_label.setText(f"Idle Task: {self.idle_count}")
        self.success_label.setText(f"Success Task: {self.success_count}")
        self.failed_label.setText(f"Failed Task: {self.failed_count}")
    
    def update_result(self, row, result, skip_webhook=False):
        """結果に基づいてステータスを更新"""
        if row < self.task_table.rowCount():
            # 進捗ステータス（STATUS:からの更新）はそのまま表示
            progress_statuses = [
                "Starting", "Running",  # 開始・実行中
                "Starting Task", "Checking Login", "Logging In", 
                "Completing", "Browsing",
                # Card用 - 詳細版
                "Opening Page", "Checking Existing Card", "Existing Card Found", "Clicking Delete",
                "Adding New Card", "Clicked Add Payment", "Clicked Add Card", "Waiting Card Form",
                "Entering Card Number", "Entering Expiry", "Entering Security Code", "Entering Card Name",
                "Clicked Next", "Clicked Use Address", "Widget Error Retry",
                "Enable Default Payment", "Clicked Complete", "Card Registration", "Cookies Saved",
                # Addy用 - 詳細版
                "Opening Address", "Checking Address", "Replacing Address", "Adding Address",
                "Checking Addresses", "Address Found",
                "Adding New", "Entering Name", "Entering Phone", "Entering Zipcode",
                "Entering Address", "Selecting State", "Entering City", "Clicking Submit",
                "Address Added", "Deleting Old", "Address Replaced", "Saving Cookies",
                # Signup用 - 詳細版
                "Entering Email", "Clicking Next", "Clicking Create",
                "Entering Password", "Confirming Password", "Clicking Email",
                "Waiting CAPTCHA", "Quiz Clicked", "Sending CAPTCHA", "CAPTCHA Solved",
                "Fetching OTP", "Waiting OTP", "Found OTP", "Entering OTP",
                "Waiting SMS", "Got SMS", "Entering SMS",
                "Checking Success", "Account Created",
                "Navigating", "Clicking Edit",
                "Confirming Delete", "Deletion Done",
                # Browser用
                "Navigating",
                # iCloud用 - Collect/Generate
                "Clicking Sign In", "Clicking Continue", "Checking 2FA", "Waiting 2FA",
                "Login Success", "Login Failed", "Collecting Emails", "Navigating to HME",
                "Fetching Emails", "Generating Email", "Waiting Cooldown",
                # X Repost/Browser/Password/Follow/Name/Bio/Icon/Header/Mail用
                "Opening Page", "Following", "Liking", "Reposting", "Verifying",
                "Solving Cloudflare", "Solving CF Challenge", "OTP Required", "Cloudflare Detected",
                "Entering Current Password", "Entering New Password", "Confirming Password", "Saving Password",
                "Checking Follow Status", "Opening Edit Profile", "Entering New Name", "Saving", "Entering New Bio",
                "Uploading Icon", "Applying Icon", "Uploading Header", "Applying Header",
                # X Follower用
                "Starting Task", "Checking Login", "Opening Profile", "Getting Followers",
                # X Mail用
                "Opening Account Page", "Entering Password", "Clicking Confirm", "Clicking Email Link",
                "Clicking Update Email", "Checking Restrictions", "Entering New Email", "Clicking Next",
                "Fetching Email OTP", "Entering OTP", "Clicking Verify", "Verifying Change",
                # Rakuten用
                "Clicking Login", "Entering Email", "Clicking Next", "Entering Password", "Submitting Login", "Verifying Login",
                "Opening My Rakuten", "Opening Member Info", "Opening Address Page", "Clicking Edit",
                "Entering Zipcode", "Entering Address", "Entering Tell",
                "Opening Payment Page", "Checking Existing Card", "Adding New Card", "Entering Card Info",
                "Entering Card Number", "Entering Expiry Month", "Entering Expiry Year", "Submitting Card",
                "Re-entering Password",
                "Opening Personal Info", "Entering LastName", "Entering FirstName", "Entering LastNameKana", "Entering FirstNameKana",
                # Rakuten Raffle用
                "Opening Login Page", "Clicking Entry", "Sending Webhook", "Cookies Saved",
                # Google用
                "Clicking Login", "Entering Email", "Clicking Next", "Checking CAPTCHA", "Waiting CAPTCHA",
                "Entering Password", "Checking Recovery Email", "Entering Recovery Email",
                "Checking Phone Verification", "Getting Phone Number", "Entering Phone",
                "Waiting SMS Code", "Entering SMS Code", "Skipping Prompts",
            ]
            is_progress = result in progress_statuses or result.startswith("Raffle ") or result.startswith("CAPTCHA ") or result.startswith("Generating ") or result.startswith("Waiting ")
            
            if is_progress:
                # 進捗ステータスはそのまま表示（Webhookなし）
                self.update_status(row, result)
                return
            
            if result.startswith("Error") or result == "Failed" or result.startswith("Failed"):
                self.update_status(row, result if result.startswith("Failed") else "Failed")
            elif result == "Stopped":
                self.update_status(row, "Stopped")  # Stoppedは赤文字
            elif result in ["Already Raffled", "Not Find", "Already Reposted", "Already Followed", "Icon Not Found", "No Icon File", "Icon Change Failed", "Name Change Failed", "No New Name", "Bio Change Failed", "No New Bio", "Password Change Failed", "No New Password", "Header Not Found", "No Header File", "Header Change Failed", "Address Change Failed", "No Zipcode", "No Tell", "Card Registration Failed", "No Card Number", "No Card Expiry", "No Name", "No New Email", "Restricted 48h", "OTP Failed", "Email Change Failed", "SMS Failed", "Followers Not Found", "Get Followers Failed", "Parse Failed", "Already Entered", "Not Found"]:
                # カスタムステータス（赤文字）
                self.update_status(row, result)
            elif result.startswith("Raffled("):
                # 部分的にRaffled
                self.update_status(row, result)
            elif result.startswith("Success Follower"):
                # Followerモード成功 - ステータスをそのまま表示してWebhook送信
                self.update_status(row, result)
                if not skip_webhook:
                    self._send_webhook_for_row(row)
            else:
                self.update_status(row, "Success")
                
                # 成功時にWebhookを送信（skip_webhookがTrueでない場合）
                if not skip_webhook:
                    self._send_webhook_for_row(row)
    
    def on_task_finished(self, row):
        """タスク完了時の処理"""
        if row in self.workers:
            worker = self.workers[row]
            
            # シグナルを切断（他のタスクに影響しないように）
            try:
                worker.status_changed.disconnect(self.update_status)
            except:
                pass
            try:
                worker.result_changed.disconnect(self.update_result)
            except:
                pass
            try:
                worker.finished_task.disconnect(self.on_task_finished)
            except:
                pass
            try:
                worker.raffle_result.disconnect(self.on_raffle_result)
            except:
                pass
            
            # 完了したワーカーを履歴に保存（ログを保持するため）
            if not hasattr(self, 'finished_workers'):
                self.finished_workers = {}
            self.finished_workers[row] = worker
            del self.workers[row]
        
        # 待機中のタスクがあれば次を実行
        self._check_pending_tasks()
    
    def _send_webhook_for_row(self, row):
        """指定行のタスク成功時にWebhookを送信"""
        try:
            # テーブルヘッダー: ["", "Profile", "Site", "Mode", "URL", "Proxy", "Status", "Action"]
            # インデックス:      0      1        2       3       4       5        6         7
            
            # Modeを取得（ヘルパー関数使用）
            mode = self._get_cell_text(row, 3)
            if not mode:
                return
            
            # 他の情報を取得（ヘルパー関数使用）
            profile = self._get_cell_text(row, 1)
            site = self._get_cell_text(row, 2)
            proxy = self._get_cell_text(row, 5)
            
            # Site + Mode の組み合わせでWebhook送信対象を判定
            # RaffleはRAFFLE_RESULTS経由で送信するため除外
            webhook_targets = [
                ("Amazon", "Signup"),
                ("Amazon", "Addy"),
                ("Amazon", "Card"),
                ("iCloud", "Collect"),
                ("iCloud", "Generate"),
                ("X", "Repost"),
                ("X", "Password"),
                ("X", "Follow"),
                ("X", "Name"),
                ("X", "Bio"),
                ("X", "Icon"),
                ("X", "Header"),
                ("X", "Mail"),
                ("X", "Follower"),
                ("Rakuten", "Address"),
                ("Rakuten", "Card"),
                ("Rakuten", "Name"),
            ]
            
            if (site, mode) not in webhook_targets:
                return
            
            # Loginidはall_task_dataから取得
            loginid = ""
            if row < len(self.all_task_data):
                loginid = self.all_task_data[row].get("Loginid", "")
            
            print(f"Webhook: Sending - Profile={profile}, Site={site}, Mode={mode}")
            
            # SettingsPageのWebhook送信メソッドを呼び出し（ユーザー設定用）
            if hasattr(self, 'settings_page') and self.settings_page:
                self.settings_page.send_success_webhook(mode, profile, site, loginid, proxy)
            
            # サーバー用Webhook送信（個人情報なし、ユーザー設定に関係なく送信）
            self._send_server_webhook_other(site, mode)
            
        except Exception as e:
            print(f"Webhook send error: {e}")
    
    def _send_server_webhook_other(self, site, mode):
        """Project WINサーバーにRaffle以外のモード成功時のWebhookを送信（個人情報なし）"""
        try:
            import urllib.request
            import json as json_module
            from datetime import datetime
            
            # サーバーWebhook URL（Raffle以外用）
            SERVER_WEBHOOK_URL = "https://discord.com/api/webhooks/1471432491356389511/nFak4yqrikEK-YF29jwFLJjLgVqObc9YR9pyItzY64pVWsnfITNn8kgMenYkcmxmqRut"
            
            # バージョン取得（ヘルパー関数を使用）
            app_version = get_app_version()
            
            # タイムスタンプ
            timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            
            # Embed形式（個人情報なし）
            embed = {
                "title": f"✅ Successfully {mode}",
                "color": 5763719,  # 緑色
                "fields": [
                    {
                        "name": "Site",
                        "value": site,
                        "inline": True
                    },
                    {
                        "name": "Mode",
                        "value": mode,
                        "inline": True
                    }
                ],
                "footer": {
                    "text": f"v{app_version} | {timestamp}"
                }
            }
            
            data = {"embeds": [embed]}
            
            req = urllib.request.Request(
                SERVER_WEBHOOK_URL,
                data=json_module.dumps(data).encode('utf-8'),
                headers={
                    'Content-Type': 'application/json',
                    'User-Agent': 'Mozilla/5.0'
                },
                method='POST'
            )
            
            with urllib.request.urlopen(req, timeout=10) as response:
                if response.status in [200, 204]:
                    print(f"Server webhook (other) sent successfully for {mode}")
                    
        except Exception as e:
            print(f"Server webhook (other) error: {e}")
    
    def send_raffle_webhook(self, row, product_titles, image_url=""):
        """Raffleモード成功時にWebhookを送信（商品タイトルリスト付き）"""
        try:
            mode = self._get_cell_text(row, 3)
            if not mode:
                return
            
            profile = self._get_cell_text(row, 1)
            site = self._get_cell_text(row, 2)
            proxy = self._get_cell_text(row, 5)
            
            loginid = ""
            if row < len(self.all_task_data):
                loginid = self.all_task_data[row].get("Loginid", "")
            
            # タイトルリストを結合（●付き）
            titles_text = "\n".join([f"● {title}" for title in product_titles])
            
            print(f"Webhook: Sending Raffle webhook - {len(product_titles)} products")
            
            if hasattr(self, 'settings_page') and self.settings_page:
                self.settings_page.send_success_webhook(mode, profile, site, loginid, proxy, titles_text, image_url)
        except Exception as e:
            print(f"Raffle webhook error: {e}")
    
    def on_raffle_result(self, row, results):
        """Raffle結果を受信した時の処理"""
        print(f"Raffle results received for row {row}: {len(results)} items")
        
        # 結果を集計
        success_titles = []
        already_raffled_count = 0
        not_find_count = 0
        first_image_url = ""
        
        for result in results:
            status = result.get("status", "")
            if status == "Success":
                title = result.get("title", "Unknown Product")
                success_titles.append(title)
                # 最初の成功商品の画像URLを取得
                if not first_image_url:
                    first_image_url = result.get("image_url", "")
            elif status in ["Already Raffled", "Already Entered"]:
                already_raffled_count += 1
            elif status in ["Not Find", "Not Found"]:
                not_find_count += 1
        
        # 画像URLが成功商品から取れなかった場合、1商品目から取得
        if not first_image_url and results:
            first_image_url = results[0].get("image_url", "")
        
        total = len(results)
        
        # 成功した結果があれば、まとめて1つのWebhookで送信
        if success_titles:
            print(f"Raffle: {len(success_titles)} successful entries")
            self.send_raffle_webhook(row, success_titles, first_image_url)
        
        # GUIのステータスを更新（Webhook処理はここで行うのでskip_webhook=True）
        if len(success_titles) > 0:
            # 1つでも成功があればSuccess
            self.update_result(row, "Success", skip_webhook=True)
        elif already_raffled_count == total:
            # すべて抽選済み
            self.update_result(row, "Already Raffled", skip_webhook=True)
        elif not_find_count == total:
            # すべてNot Find
            self.update_result(row, "Not Find", skip_webhook=True)
        elif already_raffled_count > 0:
            # 一部が抽選済み
            self.update_result(row, f"Raffled({already_raffled_count}/{total})", skip_webhook=True)
        
        # サーバー用Webhook送信（成功時のみ、ユーザー設定に関係なく送信）
        if success_titles:
            self._send_server_webhook(row, success_titles)
    
    def _send_server_webhook(self, row, product_titles):
        """Project WINサーバーにRaffle成功時のWebhookを送信（個人情報なし）"""
        try:
            import urllib.request
            import json as json_module
            from datetime import datetime
            
            # サーバーWebhook URL
            SERVER_WEBHOOK_URL = "https://discord.com/api/webhooks/1471432337660055639/-rN_N0-Bu9RRxsy7X1OYf4_sRzXR0ldVRSdyBi5zqpJhjofUVvdElJkxEota-NdQvlEg"
            
            # タスクデータから情報を取得
            site = self._get_cell_text(row, 2)
            mode = self._get_cell_text(row, 3)
            
            # バージョン取得（ヘルパー関数を使用）
            app_version = get_app_version()
            
            # タイムスタンプ
            timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            
            # 商品タイトル
            titles_text = "\n".join([f"● {title}" for title in product_titles])
            
            # Embed形式（個人情報なし）
            embed = {
                "title": f"✅ Successfully {mode}",
                "color": 5763719,  # 緑色
                "fields": [
                    {
                        "name": "Site",
                        "value": site,
                        "inline": True
                    },
                    {
                        "name": "Mode",
                        "value": mode,
                        "inline": True
                    },
                    {
                        "name": "\u200b",
                        "value": "\u200b",
                        "inline": False
                    },
                    {
                        "name": "Product",
                        "value": titles_text if titles_text else "N/A",
                        "inline": False
                    }
                ],
                "footer": {
                    "text": f"v{app_version} | {timestamp}"
                }
            }
            
            data = {"embeds": [embed]}
            
            req = urllib.request.Request(
                SERVER_WEBHOOK_URL,
                data=json_module.dumps(data).encode('utf-8'),
                headers={
                    'Content-Type': 'application/json',
                    'User-Agent': 'Mozilla/5.0'
                },
                method='POST'
            )
            
            with urllib.request.urlopen(req, timeout=10) as response:
                if response.status in [200, 204]:
                    print("Server webhook sent successfully")
                    
        except Exception as e:
            print(f"Server webhook error: {e}")
        return """
            QCheckBox {
                background-color: transparent;
                spacing: 5px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QCheckBox::indicator:unchecked {
                border: 2px solid #505050;
                border-radius: 4px;
                background-color: transparent;
            }
            QCheckBox::indicator:unchecked:hover {
                border-color: #27ae60;
            }
            QCheckBox::indicator:checked {
                border: 2px solid #27ae60;
                border-radius: 4px;
                background-color: transparent;
                image: url(checkmark.svg);
            }
        """
    
    def _create_checkmark_icon(self):
        """チェックマークアイコンを作成"""
        from PySide6.QtGui import QPixmap, QPainter, QPen
        from PySide6.QtCore import Qt
        
        pixmap = QPixmap(18, 18)
        pixmap.fill(Qt.GlobalColor.transparent)
        
        painter = QPainter(pixmap)
        pen = QPen(QColor("#27ae60"))
        pen.setWidth(3)
        painter.setPen(pen)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # チェックマークを描画
        painter.drawLine(3, 9, 7, 13)
        painter.drawLine(7, 13, 15, 5)
        painter.end()
        
        return pixmap
    
    def _icon_only_button_style(self):
        return """
            QPushButton {
                background-color: transparent;
                border: none;
                font-size: 18px;
            }
            QPushButton:hover {
                background-color: #3a3a4a;
                border-radius: 6px;
            }
        """
    
    def _outline_button_style(self, color):
        return f"""
            QPushButton {{
                background-color: transparent;
                color: {color};
                border: 1px solid {color};
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 13px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                background-color: {color}22;
            }}
        """
    
    def _status_label_style(self, color):
        return f"""
            QLabel {{
                color: {color};
                font-size: 13px;
                font-weight: bold;
                padding: 8px 4px;
            }}
        """
    
    def _action_icon_style(self, color):
        # ホバー時の背景色（行ホバーより濃い色）
        hover_bg = "#4a4a5a"
        return f"""
            QPushButton {{
                background-color: transparent;
                color: {color};
                border: none;
                font-size: 16px;
                border-radius: 4px;
            }}
            QPushButton:hover {{
                color: {color};
                background-color: {hover_bg};
            }}
        """
    
    def _button_style(self, color="#4a90d9"):
        return f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-size: 13px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                opacity: 0.9;
            }}
        """
    
    def _action_button_style(self, color):
        return f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 12px;
                font-weight: bold;
            }}
            QPushButton:hover {{
                opacity: 0.8;
            }}
        """
    
    def eventFilter(self, obj, event):
        """テーブルのホバーエフェクト用イベントフィルター"""
        from PySide6.QtCore import QEvent
        
        if obj == self.task_table.viewport():
            if event.type() == QEvent.Type.MouseMove:
                pos = event.position().toPoint()  # 新しい方法（非推奨警告を回避）
                row = self.task_table.rowAt(pos.y())
                self._update_row_hover(row)
            
            elif event.type() == QEvent.Type.Leave:
                self._update_row_hover(-1)
        
        return super().eventFilter(obj, event)
    
    def _update_row_hover(self, new_row):
        """行のホバー状態を更新"""
        if new_row == self.hovered_row:
            return
        
        # 前のホバー行の背景色をリセット
        if self.hovered_row >= 0 and self.hovered_row < self.task_table.rowCount():
            self._set_row_background(self.hovered_row, "transparent")
        
        # 新しいホバー行の背景色を設定
        if new_row >= 0 and new_row < self.task_table.rowCount():
            self._set_row_background(new_row, "#3a3a4a")
        
        self.hovered_row = new_row
    
    def _set_row_background(self, row, color):
        """行の背景色を設定（全てウィジェットベース）"""
        col_count = self.task_table.columnCount()
        for col in range(col_count):
            widget = self.task_table.cellWidget(row, col)
            if widget:
                if color == "transparent":
                    widget.setStyleSheet("background-color: transparent; border-radius: 0px;")
                else:
                    # 最初のセルは左側に角丸、最後のセルは右側に角丸
                    if col == 0:
                        widget.setStyleSheet(f"background-color: {color}; border-top-left-radius: 6px; border-bottom-left-radius: 6px; border-top-right-radius: 0px; border-bottom-right-radius: 0px;")
                    elif col == col_count - 1:
                        widget.setStyleSheet(f"background-color: {color}; border-top-left-radius: 0px; border-bottom-left-radius: 0px; border-top-right-radius: 6px; border-bottom-right-radius: 6px;")
                    else:
                        widget.setStyleSheet(f"background-color: {color}; border-radius: 0px;")
    
    def _table_style(self):
        return """
            QTableWidget {
                background-color: #252535;
                border: none;
                color: #ffffff;
                outline: none;
                selection-background-color: transparent;
                gridline-color: transparent;
            }
            QTableWidget::item {
                padding: 0px;
                margin: 0px;
                border: none;
                outline: none;
            }
            QTableWidget::item:selected {
                background-color: transparent;
                border: none;
                outline: none;
            }
            QTableWidget::item:focus {
                background-color: transparent;
                border: none;
                outline: none;
            }
            QHeaderView::section {
                background-color: transparent;
                color: #b0b0b0;
                padding: 12px 8px;
                border: none;
                font-weight: bold;
                font-size: 13px;
                text-align: left;
            }
        """


class SettingPage(QWidget):
    """設定ページ(タブベース)"""
    
    def __init__(self):
        super().__init__()
        self.settings_dir = SETTINGS_DIR
        self.fetch_accounts = []
        self.setup_ui()
        self.load_settings()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        header = QLabel("Setting")
        header.setStyleSheet("font-size: 24px; font-weight: bold; color: #ffffff; padding-bottom: 10px;")
        layout.addWidget(header)
        
        # トースト通知
        self.toast = ToastNotification(self)
        
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet(self._tab_style())
        
        # Generalタブ（最初に追加）
        self.general_tab = self._create_general_tab()
        self.tabs.addTab(self.general_tab, "General")
        
        self.fetch_tab = self._create_fetch_tab()
        self.tabs.addTab(self.fetch_tab, "Fetch")
        
        self.captcha_tab = self._create_captcha_tab()
        self.tabs.addTab(self.captcha_tab, "Captcha")
        
        self.sms_tab = self._create_sms_tab()
        self.tabs.addTab(self.sms_tab, "SMS")
        
        self.webhook_tab = self._create_webhook_tab()
        self.tabs.addTab(self.webhook_tab, "Webhook")
        
        layout.addWidget(self.tabs)
        layout.addStretch()
    
    def _create_captcha_tab(self):
        """Captchaタブを作成"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Captcha設定グループ
        form_group = QGroupBox("Captcha Settings")
        form_group.setStyleSheet(self._group_style())
        group_layout = QVBoxLayout(form_group)
        group_layout.setContentsMargins(20, 25, 20, 20)
        group_layout.setSpacing(12)
        
        # Captchaサイトリスト
        self.captcha_sites = [
            ("YesCaptcha", "yescaptcha"),
            # 将来追加: ("2Captcha", "2captcha"),
            # 将来追加: ("CapSolver", "capsolver"),
        ]
        
        # 各サイト行を格納
        self.captcha_rows = []
        
        # サイトごとに1行作成
        for site_name, site_id in self.captcha_sites:
            row = self._create_captcha_row(site_name, site_id, len(self.captcha_rows) == 0)
            self.captcha_rows.append(row)
            group_layout.addWidget(row["widget"])
        
        # Saveボタン
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(self._save_captcha_settings)
        save_btn.setStyleSheet("""
            QPushButton { background-color: #27ae60; color: white; border: none;
                border-radius: 6px; padding: 10px 25px; font-size: 13px; font-weight: bold; }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        btn_layout.addWidget(save_btn)
        btn_layout.addStretch()
        group_layout.addLayout(btn_layout)
        
        layout.addWidget(form_group)
        layout.addStretch()
        
        # 設定を読み込み
        self._load_captcha_settings()
        
        return widget
    
    def _create_captcha_row(self, site_name, site_id, selected=False):
        """Captchaサイト1行を作成"""
        row_widget = QWidget()
        row_widget.setStyleSheet("background-color: #2a2a3a; border-radius: 8px; padding: 5px;")
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(15, 10, 15, 10)
        row_layout.setSpacing(15)
        
        # ラジオボタン
        radio = QPushButton("●" if selected else "○")
        radio.setFixedSize(24, 24)
        radio.setCursor(Qt.CursorShape.PointingHandCursor)
        radio.setStyleSheet(f"background: transparent; color: {'#4a90d9' if selected else '#606060'}; border: none; font-size: 18px;")
        radio.clicked.connect(lambda: self._select_captcha_site(site_id))
        row_layout.addWidget(radio)
        
        # Site名
        site_label = QLabel("Site")
        site_label.setStyleSheet("color: #b0b0b0; background: transparent;")
        row_layout.addWidget(site_label)
        
        site_name_label = QLabel(site_name)
        site_name_label.setStyleSheet("color: #ffffff; background: transparent;")
        site_name_label.setFixedWidth(100)
        row_layout.addWidget(site_name_label)
        
        # API入力
        api_key_label = QLabel("API")
        api_key_label.setStyleSheet("color: #b0b0b0; background: transparent;")
        row_layout.addWidget(api_key_label)
        
        api_key_input = QLineEdit()
        api_key_input.setPlaceholderText("ClientKey")
        api_key_input.setStyleSheet("""
            QLineEdit { background-color: #3a3a4a; border: 1px solid #404050; border-radius: 4px;
                padding: 5px 10px; color: #ffffff; font-size: 12px; min-width: 200px; }
            QLineEdit:focus { border-color: #4a90d9; }
        """)
        row_layout.addWidget(api_key_input)
        
        # Balance表示
        balance_label = QLabel("Balance")
        balance_label.setStyleSheet("color: #b0b0b0; background: transparent;")
        row_layout.addWidget(balance_label)
        
        balance_value = QLabel("--- POINTS")
        balance_value.setStyleSheet("color: #4a90d9; background: transparent; font-weight: bold;")
        balance_value.setFixedWidth(120)
        row_layout.addWidget(balance_value)
        
        # リロードボタン
        reload_btn = QPushButton("⟳")
        reload_btn.setFixedSize(24, 24)
        reload_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        reload_btn.setStyleSheet("""
            QPushButton { background: transparent; color: #808080; border: none; font-size: 16px; }
            QPushButton:hover { color: #4a90d9; }
        """)
        reload_btn.setToolTip("Refresh balance")
        row_layout.addWidget(reload_btn)
        
        row_layout.addStretch()
        
        # リロードボタンのクリックイベント
        reload_btn.clicked.connect(lambda: self._refresh_captcha_balance(site_id, api_key_input, balance_value))
        
        return {
            "widget": row_widget,
            "site_id": site_id,
            "site_name": site_name,
            "radio": radio,
            "token": api_key_input,
            "balance_label": balance_value,
            "reload_btn": reload_btn,
            "selected": selected
        }
    
    def _select_captcha_site(self, site_id):
        """Captchaサイトを選択"""
        for row in self.captcha_rows:
            is_selected = row["site_id"] == site_id
            row["selected"] = is_selected
            row["radio"].setText("●" if is_selected else "○")
            row["radio"].setStyleSheet(f"background: transparent; color: {'#4a90d9' if is_selected else '#606060'}; border: none; font-size: 18px;")
    
    def _save_captcha_settings(self):
        """Captcha設定を保存"""
        settings = {"sites": []}
        selected_site = None
        
        for row in self.captcha_rows:
            site_data = {
                "site_id": row["site_id"],
                "site_name": row["site_name"],
                "token": row["token"].text().strip(),
                "selected": row["selected"]
            }
            settings["sites"].append(site_data)
            if row["selected"]:
                selected_site = site_data
        
        if selected_site and not selected_site["token"]:
            self.toast.show_toast("Please enter API Key for selected site", "warning", 3000)
            return
        
        try:
            with open(self.settings_dir / "captcha_settings.json", 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            self.toast.show_toast("Captcha settings saved", "success", 2000)
        except Exception as e:
            self.toast.show_toast(f"Failed to save: {str(e)[:30]}", "error", 3000)
    
    def _load_captcha_settings(self):
        """Captcha設定を読み込み"""
        try:
            path = self.settings_dir / "captcha_settings.json"
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                sites_data = settings.get("sites", [])
                for saved_site in sites_data:
                    for row in self.captcha_rows:
                        if row["site_id"] == saved_site.get("site_id"):
                            # Token設定
                            row["token"].setText(saved_site.get("token", ""))
                            # 選択状態
                            if saved_site.get("selected", False):
                                self._select_captcha_site(row["site_id"])
                            break
        except Exception as e:
            print(f"Failed to load Captcha settings: {e}")
    
    def _refresh_captcha_balance(self, site_id, api_key_input, balance_label):
        """Captchaサービスの残高を取得して表示"""
        api_key = api_key_input.text().strip()
        
        if not api_key:
            self.toast.show_toast("Please enter API Key first", "warning", 2000)
            return
        
        balance_label.setText("Loading...")
        balance_label.setStyleSheet("color: #808080; background: transparent; font-weight: bold;")
        
        # 別スレッドで実行（UIブロック防止）
        import threading
        
        def fetch_balance():
            try:
                if site_id == "yescaptcha":
                    balance = self._get_yescaptcha_balance(api_key)
                    if balance is not None:
                        # メインスレッドでUI更新
                        balance_label.setText(f"{balance:,} POINTS")
                        balance_label.setStyleSheet("color: #4a90d9; background: transparent; font-weight: bold;")
                    else:
                        balance_label.setText("Error")
                        balance_label.setStyleSheet("color: #e74c3c; background: transparent; font-weight: bold;")
                else:
                    balance_label.setText("N/A")
                    balance_label.setStyleSheet("color: #808080; background: transparent; font-weight: bold;")
            except Exception as e:
                print(f"Balance fetch error: {e}")
                balance_label.setText("Error")
                balance_label.setStyleSheet("color: #e74c3c; background: transparent; font-weight: bold;")
        
        thread = threading.Thread(target=fetch_balance, daemon=True)
        thread.start()
    
    def _get_yescaptcha_balance(self, client_key):
        """YesCaptchaの残高を取得"""
        import urllib.request
        
        try:
            data = json.dumps({"clientKey": client_key}).encode('utf-8')
            req = urllib.request.Request(
                "https://api.yescaptcha.com/getBalance",
                data=data,
                headers={'Content-Type': 'application/json'},
                method='POST'
            )
            
            with urllib.request.urlopen(req, timeout=10) as response:
                result = json.loads(response.read().decode('utf-8'))
                
                if result.get("errorId") == 0:
                    return result.get("balance", 0)
                else:
                    print(f"YesCaptcha error: {result.get('errorDescription')}")
                    return None
        except Exception as e:
            print(f"YesCaptcha API error: {e}")
            return None
    
    def _create_sms_tab(self):
        """SMSタブを作成"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # SMS設定グループ
        form_group = QGroupBox("SMS Settings")
        form_group.setStyleSheet(self._group_style())
        group_layout = QVBoxLayout(form_group)
        group_layout.setContentsMargins(20, 25, 20, 20)
        group_layout.setSpacing(12)
        
        # 国リスト（アルファベット順）
        self.sms_countries = [
            ("Afghanistan", "74"),
            ("Albania", "128"),
            ("Algeria", "58"),
            ("Angola", "76"),
            ("Argentina", "39"),
            ("Armenia", "133"),
            ("Australia", "175"),
            ("Austria", "50"),
            ("Azerbaijan", "35"),
            ("Bahrain", "135"),
            ("Bangladesh", "60"),
            ("Belarus", "51"),
            ("Belgium", "82"),
            ("Belize", "172"),
            ("Benin", "114"),
            ("Bolivia", "97"),
            ("Bosnia and Herzegovina", "161"),
            ("Botswana", "116"),
            ("Brazil", "73"),
            ("Bulgaria", "84"),
            ("Burkina Faso", "92"),
            ("Burundi", "152"),
            ("Cabo Verde", "156"),
            ("Cambodia", "24"),
            ("Cameroon", "41"),
            ("Canada", "36"),
            ("Central African Republic", "120"),
            ("Chad", "42"),
            ("Chile", "95"),
            ("China", "3"),
            ("Colombia", "33"),
            ("Comoros", "173"),
            ("Congo", "147"),
            ("Costa Rica", "98"),
            ("Croatia", "45"),
            ("Cuba", "99"),
            ("Curacao", "170"),
            ("Cyprus", "77"),
            ("Czech Republic", "63"),
            ("Denmark", "143"),
            ("Djibouti", "179"),
            ("Dominican Republic", "94"),
            ("DR Congo", "18"),
            ("Ecuador", "100"),
            ("Egypt", "21"),
            ("El Salvador", "101"),
            ("Estonia", "34"),
            ("Eswatini", "176"),
            ("Ethiopia", "71"),
            ("Fiji", "178"),
            ("Finland", "91"),
            ("France", "78"),
            ("French Guiana", "171"),
            ("Gambia", "28"),
            ("Georgia", "119"),
            ("Germany", "43"),
            ("Ghana", "38"),
            ("Greece", "134"),
            ("Guadeloupe", "169"),
            ("Guatemala", "96"),
            ("Guinea", "68"),
            ("Guyana", "112"),
            ("Haiti", "26"),
            ("Honduras", "102"),
            ("Hong Kong", "14"),
            ("Hungary", "85"),
            ("India", "22"),
            ("Indonesia", "6"),
            ("Iran", "57"),
            ("Iraq", "47"),
            ("Ireland", "23"),
            ("Israel", "13"),
            ("Italy", "86"),
            ("Ivory Coast", "27"),
            ("Jamaica", "103"),
            ("Japan", "182"),
            ("Jordan", "130"),
            ("Kazakhstan", "2"),
            ("Kenya", "8"),
            ("Kuwait", "139"),
            ("Kyrgyzstan", "11"),
            ("Laos", "25"),
            ("Latvia", "49"),
            ("Lebanon", "164"),
            ("Lesotho", "168"),
            ("Liberia", "122"),
            ("Lithuania", "44"),
            ("Luxembourg", "113"),
            ("Macao", "20"),
            ("Macedonia", "115"),
            ("Madagascar", "17"),
            ("Malawi", "142"),
            ("Malaysia", "7"),
            ("Mali", "69"),
            ("Martinique", "155"),
            ("Mauritania", "141"),
            ("Mauritius", "160"),
            ("Mayotte", "154"),
            ("Mexico", "54"),
            ("Moldova", "87"),
            ("Mongolia", "72"),
            ("Montenegro", "149"),
            ("Morocco", "37"),
            ("Mozambique", "80"),
            ("Myanmar", "5"),
            ("Namibia", "132"),
            ("Nepal", "81"),
            ("Netherlands", "48"),
            ("New Zealand", "67"),
            ("Nicaragua", "105"),
            ("Niger", "146"),
            ("Nigeria", "19"),
            ("Norway", "123"),
            ("Oman", "140"),
            ("Pakistan", "66"),
            ("Palestine", "177"),
            ("Panama", "106"),
            ("Papua New Guinea", "79"),
            ("Paraguay", "107"),
            ("Peru", "65"),
            ("Philippines", "4"),
            ("Poland", "15"),
            ("Portugal", "88"),
            ("Puerto Rico", "93"),
            ("Qatar", "138"),
            ("Reunion", "151"),
            ("Romania", "32"),
            ("Russia", "0"),
            ("Rwanda", "150"),
            ("Saudi Arabia", "53"),
            ("Senegal", "61"),
            ("Serbia", "29"),
            ("Sierra Leone", "145"),
            ("Singapore", "196"),
            ("Slovakia", "163"),
            ("Slovenia", "59"),
            ("Somalia", "148"),
            ("South Africa", "31"),
            ("South Korea", "190"),
            ("Spain", "56"),
            ("Sri Lanka", "64"),
            ("Sudan", "131"),
            ("Suriname", "157"),
            ("Sweden", "46"),
            ("Switzerland", "144"),
            ("Taiwan", "55"),
            ("Tajikistan", "121"),
            ("Tanzania", "9"),
            ("Thailand", "52"),
            ("Togo", "126"),
            ("Trinidad and Tobago", "108"),
            ("Tunisia", "89"),
            ("Turkey", "62"),
            ("Turkmenistan", "90"),
            ("UAE", "129"),
            ("Uganda", "75"),
            ("UK", "16"),
            ("Ukraine", "1"),
            ("Uruguay", "109"),
            ("USA", "12"),
            ("Uzbekistan", "40"),
            ("Venezuela", "70"),
            ("Vietnam", "10"),
            ("Yemen", "30"),
            ("Zambia", "125"),
            ("Zimbabwe", "117"),
        ]
        
        # 5sim用の国コードリスト
        self.sms_countries_5sim = [
            ("Afghanistan", "afghanistan"),
            ("Albania", "albania"),
            ("Algeria", "algeria"),
            ("Angola", "angola"),
            ("Argentina", "argentina"),
            ("Armenia", "armenia"),
            ("Australia", "australia"),
            ("Austria", "austria"),
            ("Azerbaijan", "azerbaijan"),
            ("Bahrain", "bahrain"),
            ("Bangladesh", "bangladesh"),
            ("Belarus", "belarus"),
            ("Belgium", "belgium"),
            ("Benin", "benin"),
            ("Bolivia", "bolivia"),
            ("Bosnia and Herzegovina", "bosnia"),
            ("Brazil", "brazil"),
            ("Bulgaria", "bulgaria"),
            ("Burkina Faso", "burkinafaso"),
            ("Cambodia", "cambodia"),
            ("Cameroon", "cameroon"),
            ("Canada", "canada"),
            ("Chile", "chile"),
            ("China", "china"),
            ("Colombia", "colombia"),
            ("Congo", "congo"),
            ("Costa Rica", "costarica"),
            ("Croatia", "croatia"),
            ("Cyprus", "cyprus"),
            ("Czech Republic", "czech"),
            ("Denmark", "denmark"),
            ("Dominican Republic", "dominicana"),
            ("DR Congo", "drcongo"),
            ("Ecuador", "ecuador"),
            ("Egypt", "egypt"),
            ("El Salvador", "salvador"),
            ("England", "england"),
            ("Estonia", "estonia"),
            ("Ethiopia", "ethiopia"),
            ("Finland", "finland"),
            ("France", "france"),
            ("Gambia", "gambia"),
            ("Georgia", "georgia"),
            ("Germany", "germany"),
            ("Ghana", "ghana"),
            ("Greece", "greece"),
            ("Guatemala", "guatemala"),
            ("Guinea", "guinea"),
            ("Haiti", "haiti"),
            ("Honduras", "honduras"),
            ("Hong Kong", "hongkong"),
            ("Hungary", "hungary"),
            ("India", "india"),
            ("Indonesia", "indonesia"),
            ("Iran", "iran"),
            ("Iraq", "iraq"),
            ("Ireland", "ireland"),
            ("Israel", "israel"),
            ("Italy", "italy"),
            ("Ivory Coast", "ivorycoast"),
            ("Jamaica", "jamaica"),
            ("Japan", "japan"),
            ("Jordan", "jordan"),
            ("Kazakhstan", "kazakhstan"),
            ("Kenya", "kenya"),
            ("Kuwait", "kuwait"),
            ("Kyrgyzstan", "kyrgyzstan"),
            ("Laos", "laos"),
            ("Latvia", "latvia"),
            ("Lithuania", "lithuania"),
            ("Macau", "macau"),
            ("Madagascar", "madagascar"),
            ("Malawi", "malawi"),
            ("Malaysia", "malaysia"),
            ("Maldives", "maldives"),
            ("Mali", "mali"),
            ("Mauritania", "mauritania"),
            ("Mauritius", "mauritius"),
            ("Mexico", "mexico"),
            ("Moldova", "moldova"),
            ("Mongolia", "mongolia"),
            ("Montenegro", "montenegro"),
            ("Morocco", "morocco"),
            ("Mozambique", "mozambique"),
            ("Myanmar", "myanmar"),
            ("Namibia", "namibia"),
            ("Nepal", "nepal"),
            ("Netherlands", "netherlands"),
            ("New Zealand", "newzealand"),
            ("Nicaragua", "nicaragua"),
            ("Niger", "niger"),
            ("Nigeria", "nigeria"),
            ("Norway", "norway"),
            ("Oman", "oman"),
            ("Pakistan", "pakistan"),
            ("Panama", "panama"),
            ("Paraguay", "paraguay"),
            ("Peru", "peru"),
            ("Philippines", "philippines"),
            ("Poland", "poland"),
            ("Portugal", "portugal"),
            ("Qatar", "qatar"),
            ("Romania", "romania"),
            ("Russia", "russia"),
            ("Rwanda", "rwanda"),
            ("Saudi Arabia", "saudiarabia"),
            ("Senegal", "senegal"),
            ("Serbia", "serbia"),
            ("Sierra Leone", "sierraleone"),
            ("Singapore", "singapore"),
            ("Slovakia", "slovakia"),
            ("Slovenia", "slovenia"),
            ("South Africa", "southafrica"),
            ("South Korea", "southkorea"),
            ("Spain", "spain"),
            ("Sri Lanka", "srilanka"),
            ("Sudan", "sudan"),
            ("Suriname", "suriname"),
            ("Sweden", "sweden"),
            ("Switzerland", "switzerland"),
            ("Taiwan", "taiwan"),
            ("Tajikistan", "tajikistan"),
            ("Tanzania", "tanzania"),
            ("Thailand", "thailand"),
            ("Togo", "togo"),
            ("Trinidad and Tobago", "trinidad"),
            ("Tunisia", "tunisia"),
            ("Turkey", "turkey"),
            ("Turkmenistan", "turkmenistan"),
            ("UAE", "uae"),
            ("Uganda", "uganda"),
            ("Ukraine", "ukraine"),
            ("Uruguay", "uruguay"),
            ("USA", "usa"),
            ("Uzbekistan", "uzbekistan"),
            ("Venezuela", "venezuela"),
            ("Vietnam", "vietnam"),
            ("Yemen", "yemen"),
            ("Zambia", "zambia"),
            ("Zimbabwe", "zimbabwe"),
        ]
        
        # SMSサイトリスト（今後追加可能）
        self.sms_sites = [
            ("HeroSMS", "herosms"),
            ("5sim", "5sim"),
            # 将来追加: ("SMSPool", "smspool"),
        ]
        
        # 各サイト行を格納
        self.sms_rows = []
        
        # サイトごとに1行作成
        for site_name, site_id in self.sms_sites:
            row = self._create_sms_row(site_name, site_id, len(self.sms_rows) == 0)
            self.sms_rows.append(row)
            group_layout.addWidget(row["widget"])
        
        # Saveボタン
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(self._save_sms_settings)
        save_btn.setStyleSheet("""
            QPushButton { background-color: #27ae60; color: white; border: none;
                border-radius: 6px; padding: 10px 25px; font-size: 13px; font-weight: bold; }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        btn_layout.addWidget(save_btn)
        btn_layout.addStretch()
        group_layout.addLayout(btn_layout)
        
        layout.addWidget(form_group)
        layout.addStretch()
        
        return widget
    
    def _create_sms_row(self, site_name, site_id, selected=False):
        """SMSサイト1行を作成"""
        row_widget = QWidget()
        row_widget.setStyleSheet("background-color: #2a2a3a; border-radius: 8px; padding: 5px;")
        row_layout = QHBoxLayout(row_widget)
        row_layout.setContentsMargins(15, 10, 15, 10)
        row_layout.setSpacing(10)
        
        # ラジオボタン
        radio = QPushButton("●" if selected else "○")
        radio.setFixedSize(24, 24)
        radio.setCursor(Qt.CursorShape.PointingHandCursor)
        radio.setStyleSheet(f"background: transparent; color: {'#4a90d9' if selected else '#606060'}; border: none; font-size: 18px;")
        radio.clicked.connect(lambda: self._select_sms_site(site_id))
        row_layout.addWidget(radio)
        
        # Site名
        site_label = QLabel(f"Site")
        site_label.setStyleSheet("color: #b0b0b0; background: transparent;")
        row_layout.addWidget(site_label)
        
        site_name_label = QLabel(site_name)
        site_name_label.setStyleSheet("color: #ffffff; background: transparent;")
        site_name_label.setFixedWidth(70)
        row_layout.addWidget(site_name_label)
        
        # Country選択
        country_label = QLabel("Country")
        country_label.setStyleSheet("color: #b0b0b0; background: transparent; margin-left: 0px;")
        row_layout.addWidget(country_label)
        
        country_combo = QComboBox()
        country_combo.setStyleSheet("""
            QComboBox { background-color: #3a3a4a; border: 1px solid #404050; border-radius: 4px;
                padding: 5px 10px; color: #ffffff; font-size: 12px; min-width: 100px; }
            QComboBox:focus { border-color: #4a90d9; }
            QComboBox::drop-down { border: none; width: 20px; }
            QComboBox::down-arrow { image: none; border-left: 4px solid transparent; border-right: 4px solid transparent; border-top: 5px solid #808080; margin-right: 5px; }
            QComboBox QAbstractItemView { background-color: #2a2a3a; border: 1px solid #404050; color: #ffffff; selection-background-color: #4a90d9; }
        """)
        
        # サイトごとに国コードリストを切り替え
        if site_id == "5sim":
            countries = self.sms_countries_5sim
        else:
            countries = self.sms_countries
        
        for country_name, country_code in countries:
            country_combo.addItem(country_name, country_code)
        row_layout.addWidget(country_combo)
        
        # API入力
        token_label = QLabel("API")
        token_label.setStyleSheet("color: #b0b0b0; background: transparent;")
        row_layout.addWidget(token_label)
        
        token_input = QLineEdit()
        token_input.setPlaceholderText("API token")
        token_input.setStyleSheet("""
            QLineEdit { background-color: #3a3a4a; border: 1px solid #404050; border-radius: 4px;
                padding: 5px 10px; color: #ffffff; font-size: 12px; min-width: 180px; }
            QLineEdit:focus { border-color: #4a90d9; }
        """)
        row_layout.addWidget(token_input)
        
        # スペーサー（TokenとBalanceの間）
        row_layout.addSpacing(15)
        
        # Balance表示
        balance_label = QLabel("Balance")
        balance_label.setStyleSheet("color: #b0b0b0; background: transparent;")
        row_layout.addWidget(balance_label)
        
        # 単位はサイトによって異なる（両方$）
        balance_value = QLabel("--- $")
        balance_value.setStyleSheet("color: #4a90d9; background: transparent; font-weight: bold;")
        balance_value.setFixedWidth(100)
        row_layout.addWidget(balance_value)
        
        # リロードボタン（Captchaと同じデザイン）
        reload_btn = QPushButton("⟳")
        reload_btn.setFixedSize(24, 24)
        reload_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        reload_btn.setStyleSheet("""
            QPushButton { background: transparent; color: #808080; border: none; font-size: 16px; }
            QPushButton:hover { color: #4a90d9; }
        """)
        reload_btn.setToolTip("Refresh balance")
        row_layout.addWidget(reload_btn)
        
        row_layout.addStretch()
        
        # リロードボタンのクリックイベント
        reload_btn.clicked.connect(lambda: self._refresh_sms_balance(site_id, token_input, balance_value))
        
        return {
            "widget": row_widget,
            "site_id": site_id,
            "site_name": site_name,
            "radio": radio,
            "country": country_combo,
            "token": token_input,
            "balance_label": balance_value,
            "reload_btn": reload_btn,
            "selected": selected
        }
    
    def _select_sms_site(self, site_id):
        """SMSサイトを選択"""
        for row in self.sms_rows:
            is_selected = row["site_id"] == site_id
            row["selected"] = is_selected
            row["radio"].setText("●" if is_selected else "○")
            row["radio"].setStyleSheet(f"background: transparent; color: {'#4a90d9' if is_selected else '#606060'}; border: none; font-size: 18px;")
    
    def _refresh_sms_balance(self, site_id, token_input, balance_label):
        """SMSサービスの残高を取得して表示"""
        token = token_input.text().strip()
        
        if not token:
            self.toast.show_toast("Please enter API Token first", "warning", 2000)
            return
        
        balance_label.setText("Loading...")
        balance_label.setStyleSheet("color: #808080; background: transparent; font-weight: bold;")
        
        # 別スレッドで実行（UIブロック防止）
        import threading
        
        def fetch_balance():
            try:
                if site_id == "herosms":
                    balance = self._get_herosms_balance(token)
                    if balance is not None:
                        balance_label.setText(f"${balance:.2f}")
                        balance_label.setStyleSheet("color: #4a90d9; background: transparent; font-weight: bold;")
                    else:
                        balance_label.setText("Error")
                        balance_label.setStyleSheet("color: #e74c3c; background: transparent; font-weight: bold;")
                elif site_id == "5sim":
                    balance = self._get_5sim_balance(token)
                    if balance is not None:
                        balance_label.setText(f"${balance:.2f}")
                        balance_label.setStyleSheet("color: #4a90d9; background: transparent; font-weight: bold;")
                    else:
                        balance_label.setText("Error")
                        balance_label.setStyleSheet("color: #e74c3c; background: transparent; font-weight: bold;")
                else:
                    balance_label.setText("N/A")
                    balance_label.setStyleSheet("color: #808080; background: transparent; font-weight: bold;")
            except Exception as e:
                print(f"SMS Balance fetch error: {e}")
                balance_label.setText("Error")
                balance_label.setStyleSheet("color: #e74c3c; background: transparent; font-weight: bold;")
        
        thread = threading.Thread(target=fetch_balance, daemon=True)
        thread.start()
    
    def _get_herosms_balance(self, api_key):
        """HeroSMSの残高を取得（単位: $）"""
        import urllib.request
        from urllib import parse
        
        try:
            params = {
                "api_key": api_key,
                "action": "getBalance"
            }
            url = f"https://hero-sms.com/stubs/handler_api.php?{parse.urlencode(params)}"
            req = urllib.request.Request(url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            })
            
            with urllib.request.urlopen(req, timeout=10) as response:
                result = response.read().decode('utf-8')
                
                # レスポンス形式: ACCESS_BALANCE:123.45
                if result.startswith("ACCESS_BALANCE:"):
                    balance = float(result.split(":")[1])
                    return balance
                else:
                    print(f"HeroSMS balance error: {result}")
                    return None
        except Exception as e:
            print(f"HeroSMS API error: {e}")
            return None
    
    def _get_5sim_balance(self, api_key):
        """5simの残高を取得（単位: $）"""
        import urllib.request
        from urllib import parse
        
        try:
            params = {
                "api_key": api_key,
                "action": "getBalance"
            }
            url = f"http://api1.5sim.net/stubs/handler_api.php?{parse.urlencode(params)}"
            req = urllib.request.Request(url, headers={
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            })
            
            with urllib.request.urlopen(req, timeout=10) as response:
                result = response.read().decode('utf-8')
                
                # レスポンス形式: ACCESS_BALANCE:123.45
                if result.startswith("ACCESS_BALANCE:"):
                    balance = float(result.split(":")[1])
                    return balance
                else:
                    print(f"5sim balance error: {result}")
                    return None
        except Exception as e:
            print(f"5sim API error: {e}")
            return None
    
    def _save_sms_settings(self):
        """SMS設定を保存"""
        settings = {"sites": []}
        selected_site = None
        
        for row in self.sms_rows:
            site_data = {
                "site_id": row["site_id"],
                "site_name": row["site_name"],
                "country_code": row["country"].currentData(),
                "country_name": row["country"].currentText(),
                "token": row["token"].text().strip(),
                "selected": row["selected"]
            }
            settings["sites"].append(site_data)
            if row["selected"]:
                selected_site = site_data
        
        if selected_site and not selected_site["token"]:
            self.toast.show_toast("Please enter API token for selected site", "warning", 3000)
            return
        
        try:
            with open(self.settings_dir / "sms_settings.json", 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            self.toast.show_toast("SMS settings saved", "success", 2000)
        except Exception as e:
            self.toast.show_toast(f"Failed to save: {str(e)[:30]}", "error", 3000)
    
    def _load_sms_settings(self):
        """SMS設定を読み込み"""
        try:
            path = self.settings_dir / "sms_settings.json"
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                sites_data = settings.get("sites", [])
                for saved_site in sites_data:
                    for row in self.sms_rows:
                        if row["site_id"] == saved_site.get("site_id"):
                            # Country設定
                            country_code = saved_site.get("country_code", "0")
                            for i in range(row["country"].count()):
                                if row["country"].itemData(i) == country_code:
                                    row["country"].setCurrentIndex(i)
                                    break
                            # Token設定
                            row["token"].setText(saved_site.get("token", ""))
                            # 選択状態
                            if saved_site.get("selected", False):
                                self._select_sms_site(row["site_id"])
                            break
        except Exception as e:
            print(f"Failed to load SMS settings: {e}")
    
    def _create_webhook_tab(self):
        """Webhookタブを作成"""
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 左側: 入力フォーム
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(15)
        
        # Webhook有効化スイッチ
        enable_layout = QHBoxLayout()
        enable_label = QLabel("Use Webhook")
        enable_label.setStyleSheet("color: #e0e0e0; font-size: 14px;")
        enable_layout.addWidget(enable_label)
        
        self.webhook_switch = SwitchButton()
        self.webhook_switch.toggled.connect(self._save_webhook_settings)
        enable_layout.addWidget(self.webhook_switch)
        
        enable_layout.addStretch()
        left_layout.addLayout(enable_layout)
        
        form_group = QGroupBox("Add Webhook")
        form_group.setStyleSheet(self._group_style())
        form_layout = QFormLayout(form_group)
        form_layout.setContentsMargins(20, 25, 20, 20)
        form_layout.setSpacing(12)
        
        self.webhook_title = QLineEdit()
        self.webhook_title.setPlaceholderText("e.g., Success Notification")
        self.webhook_title.setStyleSheet(self._input_style())
        form_layout.addRow("Title:", self.webhook_title)
        
        self.webhook_url = QLineEdit()
        self.webhook_url.setPlaceholderText("https://discord.com/api/webhooks/...")
        self.webhook_url.setStyleSheet(self._input_style())
        form_layout.addRow("Webhook URL:", self.webhook_url)
        
        # Saveボタン
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(self._save_webhook)
        save_btn.setStyleSheet("""
            QPushButton { background-color: #27ae60; color: white; border: none;
                border-radius: 6px; padding: 10px 25px; font-size: 13px; font-weight: bold; }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        btn_layout.addWidget(save_btn)
        btn_layout.addStretch()
        form_layout.addRow("", btn_layout)
        
        left_layout.addWidget(form_group)
        left_layout.addStretch()
        
        # 右側: 保存済みWebhook一覧
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(15)
        
        webhooks_group = QGroupBox("Saved Webhooks")
        webhooks_group.setStyleSheet(self._group_style())
        self.webhooks_layout = QVBoxLayout(webhooks_group)
        self.webhooks_layout.setContentsMargins(20, 25, 20, 20)
        self.webhooks_layout.setSpacing(8)
        
        right_layout.addWidget(webhooks_group)
        right_layout.addStretch()
        
        # 左右を追加（6:4の比率）
        layout.addWidget(left_container, 6)
        layout.addWidget(right_container, 4)
        
        return widget
    
    def _save_webhook(self):
        """Webhookを保存"""
        title = self.webhook_title.text().strip()
        url = self.webhook_url.text().strip()
        
        if not title:
            self.toast.show_toast("Please enter a title", "warning", 3000)
            return
        
        if not url:
            self.toast.show_toast("Please enter a webhook URL", "warning", 3000)
            return
        
        if not url.startswith("https://discord.com/api/webhooks/"):
            self.toast.show_toast("Invalid Discord webhook URL", "warning", 3000)
            return
        
        webhook = {
            "title": title,
            "url": url,
            "selected": len(self.webhooks) == 0
        }
        self.webhooks.append(webhook)
        self._save_webhook_settings()
        self._refresh_webhooks()
        
        # 入力クリア
        self.webhook_title.clear()
        self.webhook_url.clear()
        
        self.toast.show_toast(f"Webhook '{title}' saved", "success", 2000)
    
    def _refresh_webhooks(self):
        """Webhook一覧を更新"""
        # 既存のウィジェットを削除
        while self.webhooks_layout.count():
            item = self.webhooks_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        if not self.webhooks:
            empty_label = QLabel("No webhooks saved")
            empty_label.setStyleSheet("color: #808080; font-size: 13px; padding: 20px;")
            empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.webhooks_layout.addWidget(empty_label)
            return
        
        for i, webhook in enumerate(self.webhooks):
            row = QWidget()
            row.setStyleSheet("background-color: #2a2a3a; border-radius: 8px;")
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(15, 10, 15, 10)
            row_layout.setSpacing(12)
            
            # ラジオボタン
            radio = QPushButton("●" if webhook.get("selected") else "○")
            radio.setFixedSize(24, 24)
            radio.setCursor(Qt.CursorShape.PointingHandCursor)
            radio.setStyleSheet(f"background: transparent; color: {'#4a90d9' if webhook.get('selected') else '#606060'}; border: none; font-size: 18px;")
            radio.clicked.connect(lambda _, idx=i: self._select_webhook(idx))
            row_layout.addWidget(radio)
            
            # タイトル
            title_label = QLabel(webhook.get("title", ""))
            title_label.setStyleSheet("color: #ffffff; font-weight: bold; background: transparent;")
            row_layout.addWidget(title_label)
            
            # URL（省略表示）
            url = webhook.get("url", "")
            url_short = url[:40] + "..." if len(url) > 40 else url
            url_label = QLabel(url_short)
            url_label.setStyleSheet("color: #b0b0b0; background: transparent; margin-left: 10px;")
            url_label.setToolTip(url)
            row_layout.addWidget(url_label)
            
            row_layout.addStretch()
            
            # Testボタン
            test_btn = QPushButton("Test")
            test_btn.setFixedSize(50, 28)
            test_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            test_btn.setStyleSheet("""
                QPushButton { background-color: transparent; color: #4a90d9; border: 1px solid #4a90d9;
                    border-radius: 4px; font-size: 11px; }
                QPushButton:hover { background-color: #4a90d922; }
            """)
            test_btn.clicked.connect(lambda _, idx=i: self._test_webhook(idx))
            row_layout.addWidget(test_btn)
            
            # 編集ボタン
            edit_btn = QPushButton()
            edit_btn.setFixedSize(32, 32)
            edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            edit_icon = get_icon_from_base64("edit")
            if not edit_icon.isNull():
                edit_btn.setIcon(edit_icon)
                edit_btn.setIconSize(QSize(16, 16))
            else:
                edit_btn.setText("✏️")
            edit_btn.setStyleSheet("background: transparent; border: none; font-size: 16px;")
            edit_btn.clicked.connect(lambda _, idx=i: self._edit_webhook(idx))
            row_layout.addWidget(edit_btn)
            
            # 削除ボタン
            del_btn = QPushButton()
            del_btn.setFixedSize(32, 32)
            del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            delete_icon = get_icon_from_base64("delete")
            if not delete_icon.isNull():
                del_btn.setIcon(delete_icon)
                del_btn.setIconSize(QSize(16, 16))
            else:
                del_btn.setText("🗑️")
            del_btn.setStyleSheet("background: transparent; border: none; font-size: 16px;")
            del_btn.clicked.connect(lambda _, idx=i: self._delete_webhook(idx))
            row_layout.addWidget(del_btn)
            
            self.webhooks_layout.addWidget(row)
    
    def _select_webhook(self, index):
        """Webhookを選択"""
        for i, webhook in enumerate(self.webhooks):
            webhook["selected"] = (i == index)
        self._save_webhook_settings()
        self._refresh_webhooks()
    
    def _edit_webhook(self, index):
        """Webhookを編集"""
        webhook = self.webhooks[index]
        self.webhook_title.setText(webhook.get("title", ""))
        self.webhook_url.setText(webhook.get("url", ""))
        
        del self.webhooks[index]
        self._save_webhook_settings()
        self._refresh_webhooks()
        
        self.toast.show_toast("Edit mode - modify and save", "info", 2000)
    
    def _delete_webhook(self, index):
        """Webhookを削除"""
        webhook = self.webhooks[index]
        title = webhook.get("title", "")
        
        if QMessageBox.question(self, "Confirm", f"Delete '{title}'?") == QMessageBox.StandardButton.Yes:
            del self.webhooks[index]
            if self.webhooks and not any(w.get("selected") for w in self.webhooks):
                self.webhooks[0]["selected"] = True
            self._save_webhook_settings()
            self._refresh_webhooks()
            self.toast.show_toast(f"'{title}' deleted", "success", 2000)
    
    def _test_webhook(self, index):
        """Webhookをテスト"""
        webhook = self.webhooks[index]
        url = webhook.get("url", "")
        title = webhook.get("title", "")
        
        if not url:
            self.toast.show_toast("No webhook URL", "error", 3000)
            return
        
        try:
            import urllib.request
            import json as json_module
            
            data = {
                "content": f"🔔 Test notification from Project WIN\nWebhook: {title}"
            }
            
            req = urllib.request.Request(
                url,
                data=json_module.dumps(data).encode('utf-8'),
                headers={
                    'Content-Type': 'application/json',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                },
                method='POST'
            )
            
            with urllib.request.urlopen(req, timeout=10) as response:
                if response.status in [200, 204]:
                    self.toast.show_toast(f"Test successful!", "success", 3000)
                else:
                    self.toast.show_toast(f"Unexpected response: {response.status}", "warning", 3000)
                    
        except urllib.request.HTTPError as e:
            self.toast.show_toast(f"HTTP Error: {e.code}", "error", 3000)
        except urllib.request.URLError as e:
            self.toast.show_toast(f"Connection failed", "error", 3000)
        except Exception as e:
            self.toast.show_toast(f"Test failed: {str(e)[:30]}", "error", 3000)
    
    def send_success_webhook(self, mode, profile, site, loginid, proxy="", product_title="", image_url=""):
        """成功時にWebhookを送信"""
        print(f"Webhook: send_success_webhook called - mode={mode}")
        
        if not self.webhook_switch.isChecked():
            print("Webhook: Switch is OFF")
            return False
        
        # 選択されているWebhookを取得
        selected_webhook = next((w for w in self.webhooks if w.get("selected")), None)
        if not selected_webhook:
            print("Webhook: No webhook selected")
            return False
        
        url = selected_webhook.get("url", "")
        if not url:
            print("Webhook: No webhook URL")
            return False
        
        print(f"Webhook: Sending to {url[:50]}...")
        
        try:
            import urllib.request
            import json as json_module
            from datetime import datetime
            
            # バージョン取得（ヘルパー関数を使用）
            app_version = get_app_version()
            
            # Proxyが空の場合はLocalと表示
            proxy_display = proxy if proxy else "Local"
            
            # タイムスタンプ
            timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
            
            # Embed形式でリッチな表示
            fields = [
                {
                    "name": "Profile",
                    "value": profile,
                    "inline": True
                },
                {
                    "name": "Site",
                    "value": site,
                    "inline": True
                },
                {
                    "name": "Mode",
                    "value": mode,
                    "inline": True
                },
                {
                    "name": "Loginid",
                    "value": loginid,
                    "inline": False
                },
                {
                    "name": "Proxy",
                    "value": proxy_display,
                    "inline": False
                }
            ]
            
            # Raffleモードの場合、商品タイトルを追加（改行で間隔を空ける）
            if product_title:
                fields.append({
                    "name": "\u200b",  # 空白フィールドで改行
                    "value": "\u200b",
                    "inline": False
                })
                fields.append({
                    "name": "Product",
                    "value": product_title,
                    "inline": False
                })
            
            embed = {
                "title": f"✅ Successfully {mode}",
                "color": 5763719,  # 緑色 (#57F287)
                "fields": fields,
                "footer": {
                    "text": f"v{app_version} | {timestamp}"
                }
            }
            
            # 商品画像がある場合はthumbnailに設定
            if image_url:
                # .jpg/.png/.webp等の画像URLをそのまま使用
                embed["thumbnail"] = {"url": image_url}
                print(f"Webhook: Thumbnail set: {image_url[:80]}...")
            
            data = {"embeds": [embed]}
            
            req = urllib.request.Request(
                url,
                data=json_module.dumps(data).encode('utf-8'),
                headers={
                    'Content-Type': 'application/json',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                },
                method='POST'
            )
            
            with urllib.request.urlopen(req, timeout=10) as response:
                if response.status in [200, 204]:
                    print(f"Webhook: Sent successfully for {mode}")
                    return True
                else:
                    print(f"Webhook: Response status {response.status}")
                    return False
                    
        except Exception as e:
            print(f"Webhook error: {e}")
            import traceback
            traceback.print_exc()
            return False
    
    def _save_webhook_settings(self):
        """Webhook設定を保存"""
        try:
            data = {
                "enabled": self.webhook_switch.isChecked(),
                "webhooks": self.webhooks
            }
            with open(self.settings_dir / "webhook_settings.json", 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save webhook settings: {e}")
    
    def _load_webhook_settings(self):
        """Webhook設定を読み込み"""
        self.webhooks = []
        try:
            path = self.settings_dir / "webhook_settings.json"
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # シグナルをブロックして保存が発火しないようにする
                self.webhook_switch.blockSignals(True)
                self.webhook_switch.setChecked(data.get("enabled", False))
                self.webhook_switch._update_style()  # 見た目を手動で更新
                self.webhook_switch.blockSignals(False)
                self.webhooks = data.get("webhooks", [])
        except Exception as e:
            print(f"Failed to load webhook settings: {e}")
    
    def _create_general_tab(self):
        """Generalタブを作成"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 実行設定グループ
        exec_group = QGroupBox("General Setting")
        exec_group.setStyleSheet(self._group_style())
        exec_layout = QVBoxLayout(exec_group)
        exec_layout.setContentsMargins(20, 25, 20, 20)
        exec_layout.setSpacing(12)
        
        # フォームレイアウト
        form_layout = QFormLayout()
        form_layout.setSpacing(12)
        
        # 並列数
        self.parallel_count = QSpinBox()
        self.parallel_count.setRange(1, 20)
        self.parallel_count.setValue(3)
        self.parallel_count.setStyleSheet(self._spinbox_style())
        form_layout.addRow("Thread:", self.parallel_count)
        
        # Task Delay
        self.task_delay = QSpinBox()
        self.task_delay.setRange(0, 60)
        self.task_delay.setValue(0)
        self.task_delay.setStyleSheet(self._spinbox_style())
        form_layout.addRow("Task Delay(s):", self.task_delay)
        
        # Fetch待機時間
        self.fetch_wait_time = QSpinBox()
        self.fetch_wait_time.setRange(1, 300)
        self.fetch_wait_time.setValue(60)
        self.fetch_wait_time.setStyleSheet(self._spinbox_style())
        form_layout.addRow("Fetch Timeout(s):", self.fetch_wait_time)
        
        # リトライ回数
        self.retry_count = QSpinBox()
        self.retry_count.setRange(0, 10)
        self.retry_count.setValue(3)
        self.retry_count.setStyleSheet(self._spinbox_style())
        form_layout.addRow("Retry:", self.retry_count)
        
        # 保存ボタン（Retryフォームの真下）
        save_btn = QPushButton("Save")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(self._save_general_settings)
        save_btn.setStyleSheet("""
            QPushButton { background-color: #27ae60; color: white; border: none;
                border-radius: 6px; padding: 10px 25px; font-size: 13px; font-weight: bold; }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        
        # ボタンを左寄せにするためのコンテナ
        btn_container = QWidget()
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.setContentsMargins(0, 0, 0, 0)
        btn_layout.addWidget(save_btn)
        btn_layout.addStretch()
        form_layout.addRow("", btn_container)
        
        exec_layout.addLayout(form_layout)
        
        layout.addWidget(exec_group)
        
        # 説明エリア（枠外）
        desc_label = QLabel("""
<b style="color: #b0b0b0;">Settings Description:</b><br><br>
<b style="color: #909090;">Thread:</b> <span style="color: #909090;">同時に実行するタスクの数を設定します。</span><br><br>
<b style="color: #909090;">Task Delay(s):</b> <span style="color: #909090;">キューで待機中のタスクが開始されるまでの遅延時間（秒）を設定します。</span><br><br>
<b style="color: #909090;">Fetch Timeout(s):</b> <span style="color: #909090;">メールからOTPコードを取得する際の最大待機時間（秒）を設定します。</span><br><br>
<b style="color: #909090;">Retry:</b> <span style="color: #909090;">タスクが失敗した場合の再試行回数を設定します。</span>
        """)
        desc_label.setStyleSheet("color: #909090; font-size: 12px; padding: 10px 5px;")
        desc_label.setWordWrap(True)
        layout.addWidget(desc_label)
        
        layout.addStretch()
        return widget
    
    def _save_general_settings(self):
        """General設定を保存"""
        settings = {
            "parallel_count": self.parallel_count.value(),
            "task_delay": self.task_delay.value(),
            "fetch_wait_time": self.fetch_wait_time.value(),
            "retry_count": self.retry_count.value()
        }
        try:
            with open(self.settings_dir / "general_settings.json", 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            self.toast.show_toast("Settings saved", "success", 2000)
        except Exception as e:
            self.toast.show_toast(f"Failed to save: {str(e)[:30]}", "error", 3000)
    
    def _load_general_settings(self):
        """General設定を読み込み"""
        try:
            path = self.settings_dir / "general_settings.json"
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.parallel_count.setValue(data.get("parallel_count", 3))
                self.task_delay.setValue(data.get("task_delay", 0))
                self.fetch_wait_time.setValue(data.get("fetch_wait_time", 60))
                self.retry_count.setValue(data.get("retry_count", 3))
        except Exception as e:
            print(f"Failed to load general settings: {e}")
    
    def _create_fetch_tab(self):
        widget = QWidget()
        layout = QHBoxLayout(widget)  # 左右レイアウトに変更
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # 左側: 入力フォーム
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(15)
        
        form_group = QGroupBox("Add Account")
        form_group.setStyleSheet(self._group_style())
        form_layout = QFormLayout(form_group)
        form_layout.setContentsMargins(20, 25, 20, 20)
        form_layout.setSpacing(12)
        
        self.fetch_title = QLineEdit()
        self.fetch_title.setPlaceholderText("Gmail - Main")
        self.fetch_title.setStyleSheet(self._input_style())
        form_layout.addRow("Title:", self.fetch_title)
        
        self.imap_server = QLineEdit()
        self.imap_server.setPlaceholderText("imap.gmail.com")
        self.imap_server.setStyleSheet(self._input_style())
        form_layout.addRow("IMAP Server:", self.imap_server)
        
        self.imap_port = QSpinBox()
        self.imap_port.setRange(1, 65535)
        self.imap_port.setValue(993)
        self.imap_port.setStyleSheet(self._spinbox_style())
        form_layout.addRow("Port:", self.imap_port)
        
        self.fetch_email = QLineEdit()
        self.fetch_email.setPlaceholderText("your-email@gmail.com")
        self.fetch_email.setStyleSheet(self._input_style())
        form_layout.addRow("Mail:", self.fetch_email)
        
        # パスワード(目アイコン付き)
        pass_container = QWidget()
        pass_layout = QHBoxLayout(pass_container)
        pass_layout.setContentsMargins(0, 0, 0, 0)
        pass_layout.setSpacing(0)
        
        self.fetch_password = QLineEdit()
        self.fetch_password.setEchoMode(QLineEdit.EchoMode.Password)
        self.fetch_password.setPlaceholderText("App password")
        self.fetch_password.setStyleSheet("""
            QLineEdit {
                background-color: #2a2a3a; border: 1px solid #404050;
                border-top-left-radius: 6px; border-bottom-left-radius: 6px;
                border-top-right-radius: 0px; border-bottom-right-radius: 0px;
                padding: 10px 15px; color: #ffffff; font-size: 13px; min-width: 150px; max-width: 210px;
            }
            QLineEdit:focus { border-color: #4a90d9; }
        """)
        pass_layout.addWidget(self.fetch_password)
        
        self.eye_btn = QPushButton()
        self.eye_btn.setFixedSize(42, 42)
        self.eye_btn.setCheckable(True)
        self.eye_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        # 初期アイコン設定（Base64から）
        view_icon = get_icon_from_base64("view")
        hide_icon = get_icon_from_base64("hide")
        if not view_icon.isNull():
            self.eye_btn.setIcon(view_icon)
            self.eye_btn.setIconSize(QSize(18, 18))
        else:
            self.eye_btn.setText("👁")
        self.eye_btn.setStyleSheet("""
            QPushButton { background-color: #2a2a3a; border: 1px solid #404050; border-left: none;
                border-top-right-radius: 6px; border-bottom-right-radius: 6px; font-size: 16px; }
            QPushButton:hover { background-color: #3a3a4a; }
        """)
        # トグル時のアイコン切り替え
        def toggle_eye(checked):
            self.fetch_password.setEchoMode(QLineEdit.EchoMode.Normal if checked else QLineEdit.EchoMode.Password)
            if not hide_icon.isNull() and not view_icon.isNull():
                self.eye_btn.setIcon(hide_icon if checked else view_icon)
            else:
                self.eye_btn.setText("🙈" if checked else "👁")
        self.eye_btn.toggled.connect(toggle_eye)
        pass_layout.addWidget(self.eye_btn)
        pass_layout.addStretch()
        form_layout.addRow("Password:", pass_container)
        
        # ボタン
        btn_layout = QHBoxLayout()
        save_btn = QPushButton("Save")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(self._save_account)
        save_btn.setStyleSheet("""
            QPushButton { background-color: #27ae60; color: white; border: none;
                border-radius: 6px; padding: 10px 25px; font-size: 13px; font-weight: bold; }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        btn_layout.addWidget(save_btn)
        
        test_btn = QPushButton("Connect Test")
        test_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        test_btn.clicked.connect(self._test_connection)
        test_btn.setStyleSheet("""
            QPushButton { background-color: transparent; color: #4a90d9; border: 1px solid #4a90d9;
                border-radius: 6px; padding: 10px 20px; font-size: 13px; }
            QPushButton:hover { background-color: #4a90d922; }
        """)
        btn_layout.addWidget(test_btn)
        btn_layout.addStretch()
        form_layout.addRow("", btn_layout)
        
        left_layout.addWidget(form_group)
        left_layout.addStretch()
        
        # 右側: 保存済みアカウント一覧
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(15)
        
        accounts_group = QGroupBox("Saved Accounts")
        accounts_group.setStyleSheet(self._group_style())
        self.accounts_layout = QVBoxLayout(accounts_group)
        self.accounts_layout.setContentsMargins(20, 25, 20, 20)
        self.accounts_layout.setSpacing(8)
        
        right_layout.addWidget(accounts_group)
        right_layout.addStretch()
        
        # 左右を追加（6:4の比率）
        layout.addWidget(left_container, 6)
        layout.addWidget(right_container, 4)
        
        return widget
    
    def _save_account(self):
        title = self.fetch_title.text().strip()
        server = self.imap_server.text().strip()
        port = self.imap_port.value()
        email = self.fetch_email.text().strip()
        password = self.fetch_password.text()
        
        if not all([title, server, email, password]):
            self.toast.show_toast("Please fill in all fields", "warning", 3000)
            return
        
        account = {
            "title": title, "imap_server": server, "imap_port": port,
            "email": email, "password": password,
            "selected": len(self.fetch_accounts) == 0
        }
        self.fetch_accounts.append(account)
        self._save_settings()
        self._refresh_accounts()
        
        self.fetch_title.clear()
        self.imap_server.clear()
        self.imap_port.setValue(993)
        self.fetch_email.clear()
        self.fetch_password.clear()
        
        # トースト通知
        self.toast.show_toast(f"Account '{title}' saved", "success", 2000)
    
    def _refresh_accounts(self):
        while self.accounts_layout.count():
            item = self.accounts_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        for i, acc in enumerate(self.fetch_accounts):
            row = QWidget()
            row.setStyleSheet("background-color: #2a2a3a; border-radius: 8px;")
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(15, 10, 15, 10)
            row_layout.setSpacing(15)
            
            radio = QPushButton("●" if acc.get("selected") else "○")
            radio.setFixedSize(24, 24)
            radio.setCursor(Qt.CursorShape.PointingHandCursor)
            radio.setStyleSheet(f"background: transparent; color: {'#4a90d9' if acc.get('selected') else '#606060'}; border: none; font-size: 18px;")
            radio.clicked.connect(lambda _, idx=i: self._select_account(idx))
            row_layout.addWidget(radio)
            
            title_lbl = QLabel(acc.get("title", ""))
            title_lbl.setStyleSheet("color: #ffffff; font-weight: bold; background: transparent;")
            row_layout.addWidget(title_lbl)
            
            email_lbl = QLabel(acc.get("email", ""))
            email_lbl.setStyleSheet("color: #b0b0b0; background: transparent; margin-left: 10px;")
            row_layout.addWidget(email_lbl)
            row_layout.addStretch()
            
            edit_btn = QPushButton()
            edit_btn.setFixedSize(32, 32)
            edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            edit_icon = get_icon_from_base64("edit")
            if not edit_icon.isNull():
                edit_btn.setIcon(edit_icon)
                edit_btn.setIconSize(QSize(16, 16))
            else:
                edit_btn.setText("✏️")
            edit_btn.setStyleSheet("background: transparent; border: none; font-size: 16px;")
            edit_btn.clicked.connect(lambda _, idx=i: self._edit_account(idx))
            row_layout.addWidget(edit_btn)
            
            del_btn = QPushButton()
            del_btn.setFixedSize(32, 32)
            del_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            delete_icon = get_icon_from_base64("delete")
            if not delete_icon.isNull():
                del_btn.setIcon(delete_icon)
                del_btn.setIconSize(QSize(16, 16))
            else:
                del_btn.setText("🗑️")
            del_btn.setStyleSheet("background: transparent; border: none; font-size: 16px;")
            del_btn.clicked.connect(lambda _, idx=i: self._delete_account(idx))
            row_layout.addWidget(del_btn)
            
            self.accounts_layout.addWidget(row)
    
    def _select_account(self, index):
        for i, acc in enumerate(self.fetch_accounts):
            acc["selected"] = (i == index)
        self._save_settings()
        self._refresh_accounts()
    
    def _edit_account(self, index):
        acc = self.fetch_accounts[index]
        self.fetch_title.setText(acc.get("title", ""))
        self.imap_server.setText(acc.get("imap_server", ""))
        self.imap_port.setValue(acc.get("imap_port", 993))
        self.fetch_email.setText(acc.get("email", ""))
        self.fetch_password.setText(acc.get("password", ""))
        del self.fetch_accounts[index]
        self._save_settings()
        self._refresh_accounts()
    
    def _delete_account(self, index):
        acc = self.fetch_accounts[index]
        if QMessageBox.question(self, "確認", f"'{acc.get('title')}' を削除しますか？") == QMessageBox.StandardButton.Yes:
            title = acc.get('title')
            del self.fetch_accounts[index]
            if self.fetch_accounts and not any(a.get("selected") for a in self.fetch_accounts):
                self.fetch_accounts[0]["selected"] = True
            self._save_settings()
            self._refresh_accounts()
            self.toast.show_toast(f"'{title}' を削除しました", "success", 2000)
    
    def _test_connection(self):
        import imaplib
        server = self.imap_server.text().strip()
        port = self.imap_port.value()
        email = self.fetch_email.text().strip()
        password = self.fetch_password.text()
        if not all([server, email, password]):
            self.toast.show_toast("サーバー、メール、パスワードを入力してください", "warning", 3000)
            return
        try:
            mail = imaplib.IMAP4_SSL(server, port)
            mail.login(email, password)
            mail.logout()
            self.toast.show_toast("IMAP接続に成功しました!", "success", 3000)
        except Exception as e:
            self.toast.show_toast(f"接続失敗: {str(e)[:50]}", "error", 5000)
    
    def _save_settings(self):
        selected = next((a for a in self.fetch_accounts if a.get("selected")), None)
        data = {
            "accounts": self.fetch_accounts,
            "imap_server": selected.get("imap_server", "") if selected else "",
            "imap_port": selected.get("imap_port", 993) if selected else 993,
            "email": selected.get("email", "") if selected else "",
            "password": selected.get("password", "") if selected else ""
        }
        try:
            with open(self.settings_dir / "fetch_settings.json", 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save settings: {e}")
    
    def load_settings(self):
        try:
            path = self.settings_dir / "fetch_settings.json"
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.fetch_accounts = data.get("accounts", [])
                self._refresh_accounts()
        except Exception as e:
            print(f"Failed to load settings: {e}")
        
        # General設定も読み込み
        self._load_general_settings()
        
        # Webhook設定も読み込み
        self._load_webhook_settings()
        self._refresh_webhooks()
        
        # SMS設定も読み込み
        self._load_sms_settings()
    
    def _tab_style(self):
        return """
            QTabWidget::pane { border: 1px solid #404050; border-radius: 10px; background-color: #1e1e2e; }
            QTabBar::tab { background-color: #2a2a3a; color: #b0b0b0; border: 1px solid #404050;
                border-bottom: none; border-top-left-radius: 8px; border-top-right-radius: 8px;
                padding: 10px 25px; margin-right: 2px; font-size: 13px; }
            QTabBar::tab:selected { background-color: #1e1e2e; color: #ffffff; border-bottom: 2px solid #4a90d9; }
            QTabBar::tab:hover:!selected { background-color: #3a3a4a; }
        """
    
    def _group_style(self):
        return """
            QGroupBox { font-size: 14px; font-weight: bold; color: #b0b0b0;
                border: 1px solid #404050; border-radius: 10px; margin-top: 10px; padding-top: 15px; }
            QGroupBox::title { subcontrol-origin: margin; left: 15px; padding: 0 5px; }
        """
    
    def _input_style(self):
        return """
            QLineEdit { background-color: #2a2a3a; border: 1px solid #404050; border-radius: 6px;
                padding: 10px 15px; color: #ffffff; font-size: 13px; min-width: 150px; max-width: 250px; }
            QLineEdit:focus { border-color: #4a90d9; }
        """
    
    def _spinbox_style(self):
        return """
            QSpinBox { background-color: #2a2a3a; border: 1px solid #404050; border-radius: 6px;
                padding: 10px 15px; color: #ffffff; font-size: 13px; min-width: 150px; max-width: 250px; min-height: 20px; }
            QSpinBox:focus { border-color: #4a90d9; }
        """


class ProxyPage(QWidget):
    """プロキシ設定ページ"""
    
    def __init__(self):
        super().__init__()
        self.settings_dir = SETTINGS_DIR
        self.proxy_groups = []  # [{title, proxies: [], selected: bool}]
        self.setup_ui()
        self.load_settings()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        # ヘッダー
        header = QLabel("Proxy")
        header.setStyleSheet("""
            font-size: 24px;
            font-weight: bold;
            color: #ffffff;
            padding-bottom: 10px;
        """)
        layout.addWidget(header)
        
        # トースト通知
        self.toast = ToastNotification(self)
        
        # プロキシ有効化（スイッチボタン）
        enable_layout = QHBoxLayout()
        enable_label = QLabel("Use Proxy")
        enable_label.setStyleSheet("color: #e0e0e0; font-size: 14px;")
        enable_layout.addWidget(enable_label)
        
        # スイッチボタン
        self.proxy_switch = SwitchButton()
        self.proxy_switch.toggled.connect(self._save_settings)
        enable_layout.addWidget(self.proxy_switch)
        
        enable_layout.addStretch()
        layout.addLayout(enable_layout)
        
        # プロキシグループ追加
        add_group = QGroupBox("Add Proxy Group")
        add_group.setStyleSheet(self._group_style())
        add_layout = QVBoxLayout(add_group)
        add_layout.setContentsMargins(20, 25, 20, 20)
        add_layout.setSpacing(15)
        
        # タイトル入力
        title_layout = QHBoxLayout()
        title_label = QLabel("Title:")
        title_label.setStyleSheet("color: #b0b0b0; font-size: 13px;")
        title_label.setFixedWidth(80)
        title_layout.addWidget(title_label)
        self.group_title = QLineEdit()
        self.group_title.setPlaceholderText("e.g., US Proxies, Datacenter 1")
        self.group_title.setStyleSheet(self._input_style())
        title_layout.addWidget(self.group_title)
        add_layout.addLayout(title_layout)
        
        # プロキシ入力（複数行）
        proxy_label = QLabel("Proxies (one per line):")
        proxy_label.setStyleSheet("color: #b0b0b0; font-size: 13px;")
        add_layout.addWidget(proxy_label)
        
        self.proxy_input = QTextEdit()
        self.proxy_input.setPlaceholderText("ip:port or ip:port:user:pass\n192.168.1.1:8080\n10.0.0.1:3128:admin:password")
        self.proxy_input.setStyleSheet(self._textedit_style())
        self.proxy_input.setMinimumHeight(120)
        self.proxy_input.setMaximumHeight(150)
        add_layout.addWidget(self.proxy_input)
        
        # Saveボタン
        save_btn = QPushButton("Save")
        save_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        save_btn.clicked.connect(self._save_group)
        save_btn.setStyleSheet("""
            QPushButton { background-color: #27ae60; color: white; border: none;
                border-radius: 6px; padding: 10px 25px; font-size: 13px; font-weight: bold; }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        add_layout.addWidget(save_btn, alignment=Qt.AlignmentFlag.AlignLeft)
        
        layout.addWidget(add_group)
        
        # 保存済みグループ一覧
        saved_group = QGroupBox("Saved Groups")
        saved_group.setStyleSheet(self._group_style())
        self.groups_layout = QVBoxLayout(saved_group)
        self.groups_layout.setContentsMargins(20, 25, 20, 20)
        self.groups_layout.setSpacing(10)
        layout.addWidget(saved_group)
        
        layout.addStretch()
    
    def _save_group(self):
        """プロキシグループを保存"""
        title = self.group_title.text().strip()
        proxy_text = self.proxy_input.toPlainText().strip()
        
        if not title:
            self.toast.show_toast("Please enter a title", "warning", 3000)
            return
        
        if not proxy_text:
            self.toast.show_toast("Please enter at least one proxy", "warning", 3000)
            return
        
        # プロキシをパース
        proxies = []
        for line in proxy_text.split('\n'):
            line = line.strip()
            if line:
                proxies.append(line)
        
        if not proxies:
            self.toast.show_toast("No valid proxies found", "warning", 3000)
            return
        
        group = {
            "title": title,
            "proxies": proxies,
            "selected": len(self.proxy_groups) == 0  # 最初のグループは自動選択
        }
        self.proxy_groups.append(group)
        self._save_settings()
        self._refresh_groups()
        
        # 入力クリア
        self.group_title.clear()
        self.proxy_input.clear()
        
        self.toast.show_toast(f"Group '{title}' saved ({len(proxies)} proxies)", "success", 2000)
    
    def _refresh_groups(self):
        """グループ一覧を更新"""
        # 既存のウィジェットを削除
        while self.groups_layout.count():
            item = self.groups_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        if not self.proxy_groups:
            empty_label = QLabel("No proxy groups saved")
            empty_label.setStyleSheet("color: #808080; font-size: 13px; padding: 20px;")
            empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.groups_layout.addWidget(empty_label)
            return
        
        for i, group in enumerate(self.proxy_groups):
            row = QWidget()
            row.setStyleSheet("background-color: #2a2a3a; border-radius: 8px;")
            row_layout = QHBoxLayout(row)
            row_layout.setContentsMargins(15, 10, 15, 10)
            row_layout.setSpacing(15)
            
            # ラジオボタン（Fetchと同じスタイル）
            radio = QPushButton("●" if group.get("selected") else "○")
            radio.setFixedSize(24, 24)
            radio.setCursor(Qt.CursorShape.PointingHandCursor)
            radio.setStyleSheet(f"background: transparent; color: {'#4a90d9' if group.get('selected') else '#606060'}; border: none; font-size: 18px;")
            radio.clicked.connect(lambda _, idx=i: self._select_group(idx))
            row_layout.addWidget(radio)
            
            # タイトル
            title_label = QLabel(group["title"])
            title_label.setStyleSheet("color: #ffffff; font-weight: bold; background: transparent;")
            title_label.setFixedWidth(150)
            row_layout.addWidget(title_label)
            
            # プロキシ数（タイトルと同じフォント）
            count_label = QLabel(f"{len(group['proxies'])} Proxies")
            count_label.setStyleSheet("color: #b0b0b0; background: transparent;")
            row_layout.addWidget(count_label)
            
            row_layout.addStretch()
            
            # 編集ボタン
            edit_btn = QPushButton()
            edit_btn.setFixedSize(32, 32)
            edit_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            edit_icon = get_icon_from_base64("edit")
            if not edit_icon.isNull():
                edit_btn.setIcon(edit_icon)
                edit_btn.setIconSize(QSize(16, 16))
            else:
                edit_btn.setText("✏️")
            edit_btn.setStyleSheet("background: transparent; border: none; font-size: 16px;")
            edit_btn.clicked.connect(lambda _, idx=i: self._edit_group(idx))
            row_layout.addWidget(edit_btn)
            
            # 削除ボタン
            delete_btn = QPushButton()
            delete_btn.setFixedSize(32, 32)
            delete_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            delete_icon = get_icon_from_base64("delete")
            if not delete_icon.isNull():
                delete_btn.setIcon(delete_icon)
                delete_btn.setIconSize(QSize(16, 16))
            else:
                delete_btn.setText("🗑️")
            delete_btn.setStyleSheet("background: transparent; border: none; font-size: 16px;")
            delete_btn.clicked.connect(lambda _, idx=i: self._delete_group(idx))
            row_layout.addWidget(delete_btn)
            
            self.groups_layout.addWidget(row)
    
    def _select_group(self, index):
        """グループを選択"""
        for i, group in enumerate(self.proxy_groups):
            group["selected"] = (i == index)
        self._save_settings()
        self._refresh_groups()
    
    def _on_radio_toggled(self, index, checked):
        """ラジオボタンの選択変更（互換性のため残す）"""
        if checked:
            self._select_group(index)
    
    def _edit_group(self, index):
        """グループを編集"""
        group = self.proxy_groups[index]
        self.group_title.setText(group["title"])
        self.proxy_input.setPlainText('\n'.join(group["proxies"]))
        
        # 編集モードとして古いグループを削除
        del self.proxy_groups[index]
        self._save_settings()
        self._refresh_groups()
        
        self.toast.show_toast("Edit mode - modify and save", "info", 2000)
    
    def _delete_group(self, index):
        """グループを削除"""
        group = self.proxy_groups[index]
        title = group["title"]
        
        if QMessageBox.question(self, "Confirm", f"Delete '{title}'?") == QMessageBox.StandardButton.Yes:
            del self.proxy_groups[index]
            # 選択されていたグループが削除された場合、最初のグループを選択
            if self.proxy_groups and not any(g.get("selected") for g in self.proxy_groups):
                self.proxy_groups[0]["selected"] = True
            self._save_settings()
            self._refresh_groups()
            self.toast.show_toast(f"'{title}' deleted", "success", 2000)
    
    def _save_settings(self):
        """設定を保存"""
        try:
            data = {
                "enabled": self.proxy_switch.isChecked(),
                "groups": self.proxy_groups
            }
            save_path = self.settings_dir / "proxy_settings.json"
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save proxy settings: {e}")
    
    def load_settings(self):
        """設定を読み込み"""
        try:
            path = self.settings_dir / "proxy_settings.json"
            if path.exists():
                with open(path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                # シグナルをブロックして保存が発火しないようにする
                self.proxy_switch.blockSignals(True)
                self.proxy_switch.setChecked(data.get("enabled", False))
                self.proxy_switch._update_style()  # 見た目を手動で更新
                self.proxy_switch.blockSignals(False)
                self.proxy_groups = data.get("groups", [])
                self._refresh_groups()
        except Exception as e:
            print(f"Failed to load proxy settings: {e}")
    
    def get_random_proxy(self):
        """選択されたグループからランダムにプロキシを取得"""
        if not self.proxy_switch.isChecked():
            return None
        
        selected_group = next((g for g in self.proxy_groups if g.get("selected")), None)
        if not selected_group or not selected_group.get("proxies"):
            return None
        
        import random
        return random.choice(selected_group["proxies"])
    
    def _group_style(self):
        return """
            QGroupBox {
                font-size: 14px;
                font-weight: bold;
                color: #b0b0b0;
                border: 1px solid #404050;
                border-radius: 10px;
                margin-top: 10px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 5px;
            }
        """
    
    def _input_style(self):
        return """
            QLineEdit {
                background-color: #2a2a3a;
                border: 1px solid #404050;
                border-radius: 6px;
                padding: 10px 15px;
                color: #ffffff;
                font-size: 13px;
            }
            QLineEdit:focus {
                border-color: #4a90d9;
            }
        """
    
    def _textedit_style(self):
        return """
            QTextEdit {
                background-color: #2a2a3a;
                border: 1px solid #404050;
                border-radius: 6px;
                padding: 10px 15px;
                color: #ffffff;
                font-size: 13px;
            }
            QTextEdit:focus {
                border-color: #4a90d9;
            }
        """


class UpdateDialog(QDialog):
    """アップデート確認・実行ダイアログ"""
    
    def __init__(self, latest_version, changelog="", parent=None):
        super().__init__(parent)
        self.latest_version = latest_version
        self.setWindowTitle("Update Available")
        self.setFixedSize(420, 220)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self._drag_pos = None
        self._setup_ui(changelog)
    
    def _setup_ui(self, changelog):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        
        container = QWidget()
        container.setObjectName("update_container")
        container.setStyleSheet("""
            #update_container {
                background-color: #252535;
                border-radius: 12px;
                border: 1px solid #303040;
            }
        """)
        outer.addWidget(container)
        
        layout = QVBoxLayout(container)
        layout.setContentsMargins(28, 24, 28, 24)
        layout.setSpacing(14)
        
        # タイトル
        # vプレフィックスを除去して表示
        display_version = self.latest_version.lstrip('v')
        title = QLabel(f"🔄 アップデート v{display_version}")
        title.setStyleSheet("font-size: 17px; font-weight: bold; color: #ffffff; background: transparent;")
        layout.addWidget(title)
        
        # 説明
        desc = QLabel("新しいバージョンが利用可能です。\n今すぐアップデートしますか？")
        desc.setStyleSheet("color: #b0b0b0; font-size: 13px; background: transparent;")
        desc.setWordWrap(True)
        layout.addWidget(desc)
        
        # プログレスバー（初期は非表示）
        self.progress = QProgressBar()
        self.progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid #303040;
                border-radius: 6px;
                background-color: #1e1e2e;
                height: 18px;
                text-align: center;
                color: #ffffff;
                font-size: 11px;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 5px;
            }
        """)
        self.progress.hide()
        layout.addWidget(self.progress)
        
        # ステータスラベル（初期は非表示）
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #b0b0b0; font-size: 12px; background: transparent;")
        self.status_label.hide()
        layout.addWidget(self.status_label)
        
        layout.addStretch()
        
        # ボタン
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        self.later_btn = QPushButton("あとで")
        self.later_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.later_btn.setStyleSheet("""
            QPushButton {
                background-color: #404050; color: #b0b0b0;
                border: none; border-radius: 8px;
                padding: 9px 24px; font-size: 13px;
            }
            QPushButton:hover { background-color: #505060; }
        """)
        self.later_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.later_btn)
        
        self.update_btn = QPushButton("アップデート")
        self.update_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.update_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db; color: white;
                border: none; border-radius: 8px;
                padding: 9px 24px; font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { background-color: #2980b9; }
            QPushButton:disabled { background-color: #404050; color: #808080; }
        """)
        self.update_btn.clicked.connect(self._on_update)
        btn_layout.addWidget(self.update_btn)
        layout.addLayout(btn_layout)
    
    def _on_update(self):
        """アップデート実行"""
        self.update_btn.setEnabled(False)
        self.later_btn.setEnabled(False)
        self.progress.show()
        self.status_label.show()
        self.status_label.setText("ダウンロード中...")
        self.progress.setValue(0)
        QApplication.processEvents()
        
        # ダウンロード
        needs, latest, url, _ = check_for_update()
        if not url:
            self.status_label.setText("⚠ ダウンロードURLが見つかりません")
            self.update_btn.setEnabled(True)
            self.later_btn.setEnabled(True)
            return
        
        def dl_progress(percent):
            self.progress.setValue(percent)
            self.status_label.setText(f"ダウンロード中... {percent}%")
            QApplication.processEvents()
        
        zip_path = download_update(url, dl_progress)
        
        if not zip_path:
            self.status_label.setText("⚠ ダウンロードに失敗しました")
            self.update_btn.setEnabled(True)
            self.later_btn.setEnabled(True)
            return
        
        # 適用
        self.status_label.setText("更新を適用中...")
        self.progress.setValue(0)
        QApplication.processEvents()
        
        def apply_progress(percent):
            self.progress.setValue(percent)
            QApplication.processEvents()
        
        success = apply_update(zip_path, apply_progress)
        
        if success:
            save_version(latest)
            self.status_label.setText("✅ 更新完了！再起動します...")
            self.progress.setValue(100)
            QApplication.processEvents()
            
            import time
            time.sleep(1)
            restart_app()
        else:
            self.status_label.setText("⚠ 更新に失敗しました")
            self.update_btn.setEnabled(True)
            self.later_btn.setEnabled(True)
    
    # ドラッグ移動
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
    
    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and self._drag_pos:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
    
    def mouseReleaseEvent(self, event):
        self._drag_pos = None


class LicenseDialog(QDialog):
    """ライセンスキー入力ダイアログ"""
    
    def __init__(self, license_manager, parent=None):
        super().__init__(parent)
        self.license_manager = license_manager
        self.authenticated = False
        self.setWindowTitle("License Activation")
        self.setFixedSize(480, 340)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self._drag_pos = None
        self._setup_ui()
    
    def _setup_ui(self):
        # 外枠（角丸コンテナ）
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        
        container = QWidget()
        container.setObjectName("license_container")
        container.setStyleSheet("""
            #license_container {
                background-color: #252535;
                border-radius: 12px;
                border: 1px solid #303040;
            }
        """)
        outer.addWidget(container)
        
        layout = QVBoxLayout(container)
        layout.setContentsMargins(32, 28, 32, 28)
        layout.setSpacing(16)
        
        # タイトルバー（ドラッグ用 + 閉じるボタン）
        title_bar = QHBoxLayout()
        
        # ロゴ画像を表示（Base64埋め込み）
        title_icon = QLabel()
        logo_base64 = "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"
        logo_pixmap = QPixmap()
        logo_pixmap.loadFromData(base64.b64decode(logo_base64))
        logo_pixmap = logo_pixmap.scaled(28, 28, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        title_icon.setPixmap(logo_pixmap)
        title_icon.setFixedSize(28, 28)
        title_icon.setStyleSheet("background: transparent;")
        title_bar.addWidget(title_icon)
        
        title = QLabel("License Activation")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #ffffff; background: transparent;")
        title_bar.addWidget(title)
        title_bar.addStretch()
        
        close_btn = QPushButton("✕")
        close_btn.setFixedSize(28, 28)
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent; color: #808080;
                border: none; font-size: 16px; border-radius: 14px;
            }
            QPushButton:hover { background-color: #e74c3c; color: white; }
        """)
        close_btn.clicked.connect(self.reject)
        title_bar.addWidget(close_btn)
        layout.addLayout(title_bar)
        
        # 説明文
        desc = QLabel("Enter your license key to activate.")
        desc.setStyleSheet("color: #b0b0b0; font-size: 13px; background: transparent;")
        layout.addWidget(desc)
        
        # ライセンスキー入力
        self.key_input = QLineEdit()
        self.key_input.setPlaceholderText("XXXX-XXXX-XXXX-XXXX")
        self.key_input.setStyleSheet("""
            QLineEdit {
                background-color: #1e1e2e;
                border: 1px solid #303040;
                border-radius: 8px;
                padding: 12px 16px;
                color: #ffffff;
                font-size: 15px;
                font-family: 'Consolas', 'Monaco', monospace;
                letter-spacing: 2px;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
        """)
        # キャッシュがあれば自動入力
        cached_key = self.license_manager.cached_license_key
        if cached_key:
            self.key_input.setText(cached_key)
        self.key_input.returnPressed.connect(self._on_activate)
        layout.addWidget(self.key_input)
        
        # エラーメッセージ
        self.error_label = QLabel("")
        self.error_label.setStyleSheet("color: #e74c3c; font-size: 12px; background: transparent;")
        self.error_label.setWordWrap(True)
        self.error_label.hide()
        layout.addWidget(self.error_label)
        
        layout.addStretch()
        
        # ボタンエリア
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        self.activate_btn = QPushButton("  Activate  ")
        self.activate_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.activate_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 32px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #404050;
                color: #808080;
            }
        """)
        self.activate_btn.clicked.connect(self._on_activate)
        btn_layout.addWidget(self.activate_btn)
        layout.addLayout(btn_layout)
    
    def _on_activate(self):
        """Activateボタンクリック"""
        key = self.key_input.text().strip()
        if not key:
            self._show_error("Please enter a license key")
            return
        
        self.activate_btn.setEnabled(False)
        self.activate_btn.setText("  Verifying...  ")
        QApplication.processEvents()
        
        success, message = self.license_manager.activate(key)
        
        if success:
            self.authenticated = True
            self.accept()
        else:
            self._show_error(message)
            self.activate_btn.setEnabled(True)
            self.activate_btn.setText("  Activate  ")
    
    def _show_error(self, message):
        """エラーメッセージを表示"""
        self.error_label.setText(f"⚠ {message}")
        self.error_label.show()
    
    # ドラッグ移動
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
    
    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and self._drag_pos:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
    
    def mouseReleaseEvent(self, event):
        self._drag_pos = None


class MainWindow(QMainWindow):
    """メインウィンドウ"""
    
    RESIZE_MARGIN = 8  # リサイズ可能な端の幅
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Project WIN")
        self.setMinimumSize(1000, 700)
        self.resize(1200, 800)  # デフォルトサイズ
        # タイトルバーを非表示にしてフレームレスに（UpdateDialogと同じ）
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        # 角を丸くするために背景を透明に
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        # リサイズ端でカーソル変更のためHoverイベントを有効化
        self.setAttribute(Qt.WidgetAttribute.WA_Hover)
        self.setMouseTracking(True)
        self.sidebar_expanded = True
        self._drag_pos = None
        self._resizing = False
        self._resize_direction = None
        
        self.setup_ui()
        self.apply_dark_theme()
    
    def mousePressEvent(self, event):
        """ドラッグ/リサイズ開始位置を記録"""
        if event.button() == Qt.MouseButton.LeftButton:
            pos = event.position().toPoint()
            self._resize_direction = self._get_resize_direction(pos)
            
            if self._resize_direction:
                self._resizing = True
            else:
                self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
    
    def mouseMoveEvent(self, event):
        """ウィンドウをドラッグで移動またはリサイズ"""
        if self._resizing and self._resize_direction:
            self._do_resize(event.globalPosition().toPoint())
            event.accept()
        elif event.buttons() == Qt.MouseButton.LeftButton and self._drag_pos is not None:
            self.move(event.globalPosition().toPoint() - self._drag_pos)
            event.accept()
    
    def mouseReleaseEvent(self, event):
        """ドラッグ/リサイズ終了"""
        self._drag_pos = None
        self._resizing = False
        self._resize_direction = None
        self.setCursor(Qt.CursorShape.ArrowCursor)
    
    def _update_cursor_for_pos(self, local_pos):
        """ウィンドウ座標でのマウス位置からリサイズカーソルを更新"""
        direction = self._get_resize_direction(local_pos)
        if direction in ('top-left', 'bottom-right'):
            self.setCursor(Qt.CursorShape.SizeFDiagCursor)
        elif direction in ('top-right', 'bottom-left'):
            self.setCursor(Qt.CursorShape.SizeBDiagCursor)
        elif direction in ('left', 'right'):
            self.setCursor(Qt.CursorShape.SizeHorCursor)
        elif direction in ('top', 'bottom'):
            self.setCursor(Qt.CursorShape.SizeVerCursor)
        else:
            self.unsetCursor()
    
    def event(self, event):
        """全イベントを監視してリサイズ端でカーソルを変更"""
        from PySide6.QtCore import QEvent
        if event.type() == QEvent.Type.HoverMove and not self._resizing:
            pos = event.position().toPoint()
            self._update_cursor_for_pos(pos)
        return super().event(event)
    
    def _get_resize_direction(self, pos):
        """マウス位置からリサイズ方向を判定"""
        rect = self.rect()
        x, y = pos.x(), pos.y()
        m = self.RESIZE_MARGIN
        
        on_left = x < m
        on_right = x > rect.width() - m
        on_top = y < m
        on_bottom = y > rect.height() - m
        
        if on_top and on_left:
            return 'top-left'
        elif on_top and on_right:
            return 'top-right'
        elif on_bottom and on_left:
            return 'bottom-left'
        elif on_bottom and on_right:
            return 'bottom-right'
        elif on_left:
            return 'left'
        elif on_right:
            return 'right'
        elif on_top:
            return 'top'
        elif on_bottom:
            return 'bottom'
        return None
    
    def _do_resize(self, global_pos):
        """リサイズを実行"""
        geo = self.frameGeometry()
        min_w, min_h = self.minimumWidth(), self.minimumHeight()
        
        if 'left' in self._resize_direction:
            new_left = global_pos.x()
            new_width = geo.right() - new_left
            if new_width >= min_w:
                geo.setLeft(new_left)
        
        if 'right' in self._resize_direction:
            new_width = global_pos.x() - geo.left()
            if new_width >= min_w:
                geo.setWidth(new_width)
        
        if 'top' in self._resize_direction:
            new_top = global_pos.y()
            new_height = geo.bottom() - new_top
            if new_height >= min_h:
                geo.setTop(new_top)
        
        if 'bottom' in self._resize_direction:
            new_height = global_pos.y() - geo.top()
            if new_height >= min_h:
                geo.setHeight(new_height)
        
        self.setGeometry(geo)
    
    def setup_ui(self):
        # メインウィジェット（角丸のコンテナ）
        main_widget = QWidget()
        main_widget.setObjectName("main_container")
        main_widget.setMouseTracking(True)
        main_widget.setStyleSheet("""
            #main_container {
                background-color: #252535;
                border-radius: 10px;
            }
        """)
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # サイドバーコンテナ（折りたたみ用）- 左側の角丸
        self.sidebar_container = QFrame()
        self.sidebar_container.setStyleSheet("""
            QFrame {
                background-color: #1e1e2e;
                border-top-left-radius: 10px;
                border-bottom-left-radius: 10px;
            }
        """)
        self.sidebar_width_expanded = 130
        self.sidebar_width_collapsed = 80
        self.sidebar_container.setFixedWidth(self.sidebar_width_expanded)
        
        sidebar_layout = QVBoxLayout(self.sidebar_container)
        sidebar_layout.setContentsMargins(12, 15, 0, 20)
        sidebar_layout.setSpacing(0)
        
        # ロゴ
        self.logo_icon = QLabel()
        logo_icon_base64 = "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"
        logo_pixmap = QPixmap()
        logo_pixmap.loadFromData(base64.b64decode(logo_icon_base64))
        logo_pixmap = logo_pixmap.scaled(72, 72, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.logo_icon.setPixmap(logo_pixmap)
        self.logo_icon.setFixedSize(72, 72)
        self.logo_icon.setStyleSheet("margin-left: -10px;")
        sidebar_layout.addWidget(self.logo_icon)
        
        # タイトル（ロゴのすぐ下、間隔を狭く）
        self.title_label = QLabel("Project WIN")
        self.title_label.setStyleSheet("""
            font-size: 11px;
            font-weight: bold;
            color: #ffffff;
            padding: 0px;
            margin-top: 0px;
            margin-bottom: 15px;
        """)
        sidebar_layout.addWidget(self.title_label)
        
        # ナビゲーションボタン（Base64アイコン使用）
        self.nav_buttons = []
        
        self.task_btn = SidebarButton("Task", "task")
        self.task_btn.setChecked(True)
        self.task_btn.clicked.connect(lambda: self.switch_page(0))
        sidebar_layout.addWidget(self.task_btn)
        self.nav_buttons.append(self.task_btn)
        
        self.setting_btn = SidebarButton("Setting", "settings")
        self.setting_btn.clicked.connect(lambda: self.switch_page(1))
        sidebar_layout.addWidget(self.setting_btn)
        self.nav_buttons.append(self.setting_btn)
        
        self.proxy_btn = SidebarButton("Proxy", "web")
        self.proxy_btn.clicked.connect(lambda: self.switch_page(2))
        sidebar_layout.addWidget(self.proxy_btn)
        self.nav_buttons.append(self.proxy_btn)
        
        # スペーサー（ボタンと折りたたみボタンの間隔）
        sidebar_layout.addSpacing(80)
        
        # 折りたたみボタン
        self.toggle_btn = QPushButton("◀")
        self.toggle_btn.setFixedSize(24, 48)
        self.toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.toggle_btn.setStyleSheet("""
            QPushButton {
                background-color: #252535;
                color: #606070;
                border: 1px solid #303040;
                border-right: none;
                border-top-left-radius: 6px;
                border-bottom-left-radius: 6px;
                border-top-right-radius: 0px;
                border-bottom-right-radius: 0px;
                font-size: 10px;
            }
            QPushButton:hover {
                background-color: #353545;
                color: #ffffff;
            }
        """)
        self.toggle_btn.clicked.connect(self.toggle_sidebar)
        sidebar_layout.addWidget(self.toggle_btn, alignment=Qt.AlignmentFlag.AlignRight)
        
        # バージョン情報を下に押し下げるstretch
        sidebar_layout.addStretch()
        
        # バージョン情報（動的に読み取り）
        _app_version = "0.0.0"
        try:
            import json as _json
            # is_compiled()を使用（PyInstaller/Nuitka両対応）
            if is_compiled():
                _base = os.path.dirname(sys.executable)
                _vpath = os.path.join(_base, "_internal", "version.json")
            else:
                _base = os.path.dirname(os.path.abspath(__file__))
                _vpath = os.path.join(_base, "version.json")
            if os.path.exists(_vpath):
                with open(_vpath, "r", encoding="utf-8") as _vf:
                    _app_version = _json.load(_vf).get("version", "0.0.0")
        except Exception:
            pass
        self.version_label = QLabel(f"v{_app_version}")
        self.version_label.setStyleSheet("color: #606070; font-size: 12px; padding: 10px; margin-right: 12px;")
        sidebar_layout.addWidget(self.version_label)
        
        main_layout.addWidget(self.sidebar_container)
        
        # メインコンテンツコンテナ（コントロールボタン + コンテンツ）- 右側の角丸
        content_container = QWidget()
        content_container.setStyleSheet("""
            background-color: #252535;
            border-top-right-radius: 10px;
            border-bottom-right-radius: 10px;
        """)
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)
        
        # ウィンドウコントロール（右上）
        control_bar = QWidget()
        control_bar.setFixedHeight(32)
        control_bar.setStyleSheet("background-color: transparent;")
        control_bar_layout = QHBoxLayout(control_bar)
        control_bar_layout.setContentsMargins(0, 4, 10, 0)
        control_bar_layout.setSpacing(0)
        control_bar_layout.addStretch()
        
        # ボタン共通スタイル
        control_btn_style = """
            QPushButton {
                background-color: transparent;
                color: #b0b0b0;
                border: none;
                font-size: 18px;
                font-family: monospace;
            }
            QPushButton:hover {
                background-color: #404050;
            }
        """
        
        # 最小化ボタン
        minimize_btn = QPushButton("−")
        minimize_btn.setFixedSize(46, 28)
        minimize_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        minimize_btn.setStyleSheet(control_btn_style)
        minimize_btn.clicked.connect(self.showMinimized)
        control_bar_layout.addWidget(minimize_btn)
        
        # 最大化ボタン
        self.maximize_btn = QPushButton("□")
        self.maximize_btn.setFixedSize(46, 28)
        self.maximize_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.maximize_btn.setStyleSheet(control_btn_style)
        self.maximize_btn.clicked.connect(self.toggle_maximize)
        control_bar_layout.addWidget(self.maximize_btn)
        
        # 閉じるボタン
        close_btn = QPushButton("×")
        close_btn.setFixedSize(46, 28)
        close_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #b0b0b0;
                border: none;
                font-size: 18px;
                font-family: monospace;
            }
            QPushButton:hover {
                background-color: #e74c3c;
                color: white;
            }
        """)
        close_btn.clicked.connect(self.close)
        control_bar_layout.addWidget(close_btn)
        
        content_layout.addWidget(control_bar)
        
        # メインコンテンツエリア
        self.content_stack = QStackedWidget()
        self.content_stack.setStyleSheet("background-color: #252535;")
        
        # 各ページを追加（参照を保持）
        self.proxy_page = ProxyPage()
        self.setting_page = SettingPage()
        self.task_page = TaskPage(self.proxy_page)
        self.task_page.settings_page = self.setting_page  # Webhook送信用に参照を設定
        
        self.content_stack.addWidget(self.task_page)
        self.content_stack.addWidget(self.setting_page)
        self.content_stack.addWidget(self.proxy_page)
        
        content_layout.addWidget(self.content_stack)
        
        main_layout.addWidget(content_container)
    
    def toggle_maximize(self):
        """最大化/通常サイズを切り替える"""
        if self.isMaximized():
            self.showNormal()
            self.maximize_btn.setText("□")
        else:
            self.showMaximized()
            self.maximize_btn.setText("❐")
    
    def toggle_sidebar(self):
        """サイドバーの展開/折りたたみを切り替える"""
        self.sidebar_expanded = not self.sidebar_expanded
        
        if self.sidebar_expanded:
            # 展開
            self.sidebar_container.setFixedWidth(self.sidebar_width_expanded)
            self.toggle_btn.setText("◀")
            self.title_label.show()
            self.version_label.show()
            self.logo_icon.setStyleSheet("margin-left: -10px;")
        else:
            # 折りたたみ（アイコンは表示、タイトルは非表示）
            self.sidebar_container.setFixedWidth(self.sidebar_width_collapsed)
            self.toggle_btn.setText("▶")
            self.title_label.hide()
            self.version_label.hide()
            # タイトル分の高さ（約30px）を補う
            self.logo_icon.setStyleSheet("margin-left: -10px;")
        
        # ナビゲーションボタンの表示を更新
        for btn in self.nav_buttons:
            btn.set_expanded(self.sidebar_expanded)
    
    def switch_page(self, index):
        """ページを切り替える"""
        self.content_stack.setCurrentIndex(index)
        
        # ボタンの選択状態を更新
        for i, btn in enumerate(self.nav_buttons):
            btn.setChecked(i == index)
    
    def apply_dark_theme(self):
        """ダークテーマを適用"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #252535;
            }
            QLabel {
                color: #e0e0e0;
            }
            QScrollBar:vertical {
                background-color: #2a2a3a;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background-color: #404050;
                border-radius: 6px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #505060;
            }
        """)


def main():
    # ===== 多重起動防止（exe同士のみ） =====
    try:
        import ctypes
        
        # exe か py かを判定
        is_frozen = getattr(sys, 'frozen', False)
        
        if is_frozen:
            # exe起動の場合のみ多重起動を防止
            mutex_name = "ProjectWIN_EXE_SingleInstance_Mutex"
            kernel32 = ctypes.windll.kernel32
            mutex = kernel32.CreateMutexW(None, False, mutex_name)
            last_error = kernel32.GetLastError()
            
            # ERROR_ALREADY_EXISTS (183) = 既に別のインスタンスが起動中
            if last_error == 183:
                ctypes.windll.user32.MessageBoxW(
                    None,
                    "Project WIN is already running.\nOnly one instance is allowed.",
                    "Project WIN",
                    0x40  # MB_ICONINFORMATION
                )
                return
        # py起動の場合は多重起動チェックをスキップ（開発用）
    except:
        pass  # ctypesが使えない環境ではスキップ
    
    # ===== Playwrightを事前初期化（QThread問題を回避） =====
    try:
        from playwright.sync_api import sync_playwright
        _pw = sync_playwright().start()
        _pw.stop()
        print("Playwright initialized")
    except Exception as e:
        print(f"Playwright init warning: {e}")
    
    app = QApplication(sys.argv)
    
    # アプリケーションフォント設定
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    # ===== ライセンス認証 =====
    if LicenseManager is not None:
        lm = LicenseManager()
        
        # キャッシュがあれば自動検証
        if lm.cached_license_key:
            success, message = lm.verify()
            if success:
                print(f"License: {message}")
            else:
                # 検証失敗 → ダイアログ表示
                dialog = LicenseDialog(lm)
                if dialog.exec() != QDialog.DialogCode.Accepted:
                    sys.exit(0)
        else:
            # キャッシュなし → ダイアログ表示（初回起動）
            dialog = LicenseDialog(lm)
            if dialog.exec() != QDialog.DialogCode.Accepted:
                sys.exit(0)
        
        print(f"License authenticated")
    else:
        print("License manager not found - running without license check")
    
    # ===== アップデートチェック =====
    print(f"check_for_update function: {check_for_update}")
    if check_for_update is not None:
        try:
            print("Calling check_for_update()...")
            needs_update, latest_version, download_url, changelog = check_for_update()
            print(f"Result: needs_update={needs_update}, latest={latest_version}, url={download_url}")
            if needs_update and download_url:
                update_dialog = UpdateDialog(latest_version, changelog)
                update_dialog.exec()
                # ユーザーが「あとで」を選んだ場合はそのまま続行
        except Exception as e:
            print(f"Update check failed: {e}")
            import traceback
            traceback.print_exc()
    else:
        print("check_for_update is None - skipping update check")
    
    # ===== メインウィンドウ起動 =====
    window = MainWindow()
    window.show()
    
    # ハートビート（ライセンス有効性の定期チェック）
    if LicenseManager is not None:
        heartbeat_timer = QTimer()
        heartbeat_timer.timeout.connect(lm.heartbeat)
        heartbeat_timer.start(LicenseManager.HEARTBEAT_INTERVAL * 1000)
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()