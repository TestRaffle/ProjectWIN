"""
Project WIN アップデーター
GitHub Releasesから最新版を自動ダウンロード＆更新
"""

import os
import sys
import json
import shutil
import tempfile
import zipfile
import subprocess
from pathlib import Path
from urllib import request, error

# 現在のバージョン
CURRENT_VERSION = "1.0.0"

# GitHub リポジトリ情報
GITHUB_OWNER = "TestRaffle"
GITHUB_REPO = "ProjectWIN"

# GitHub API URL
RELEASES_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"

# バージョン保存ファイル
VERSION_FILE = None


def get_app_dir():
    """アプリケーションのディレクトリを取得"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


def get_version_file():
    """バージョンファイルのパスを取得"""
    global VERSION_FILE
    if VERSION_FILE is None:
        VERSION_FILE = get_app_dir() / "_internal" / "version.json"
    return VERSION_FILE


def parse_version(version_str):
    """バージョン文字列をタプルに変換（比較用）"""
    if not version_str:
        return (0, 0, 0)
    clean = version_str.lstrip('v').strip()
    try:
        parts = clean.split('.')
        return tuple(int(p) for p in parts[:3])
    except:
        return (0, 0, 0)


def get_current_version():
    """現在のバージョンを取得"""
    try:
        version_file = get_version_file()
        if version_file.exists():
            content = version_file.read_text().strip()
            # JSON形式の場合
            if content.startswith('{'):
                import json
                data = json.loads(content)
                return data.get('version', CURRENT_VERSION)
            # プレーンテキストの場合
            return content
    except:
        pass
    return CURRENT_VERSION


def save_version(version):
    """バージョンを保存"""
    try:
        version_file = get_version_file()
        version_file.parent.mkdir(parents=True, exist_ok=True)
        version_file.write_text(version)
        return True
    except:
        return False


def check_for_update():
    """
    最新バージョンをチェック
    
    Returns:
        tuple: (needs_update, latest_version, download_url, changelog)
    """
    try:
        print("Checking for updates...")
        current_ver = get_current_version()
        print(f"Current version: {current_ver}")
        
        req = request.Request(
            RELEASES_API_URL,
            headers={
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': 'ProjectWIN-Updater'
            }
        )
        
        with request.urlopen(req, timeout=10) as response:
            data = json.loads(response.read().decode('utf-8'))
        
        latest_version = data.get('tag_name', '')
        changelog = data.get('body', '')
        print(f"Latest version: {latest_version}")
        
        # ダウンロードURLを探す（.zipファイル）
        download_url = None
        assets = data.get('assets', [])
        for asset in assets:
            if asset['name'].endswith('.zip'):
                download_url = asset['browser_download_url']
                break
        
        print(f"Download URL: {download_url}")
        
        # バージョン比較
        current = parse_version(current_ver)
        latest = parse_version(latest_version)
        
        print(f"Parsed current: {current}, latest: {latest}")
        
        needs_update = latest > current
        print(f"Needs update: {needs_update}")
        
        return needs_update, latest_version, download_url, changelog
        
    except Exception as e:
        print(f"Update check failed: {e}")
        import traceback
        traceback.print_exc()
        return False, None, None, None


def download_update(download_url, progress_callback=None):
    """
    アップデートをダウンロード
    
    Args:
        download_url: ダウンロードURL
        progress_callback: 進捗コールバック (percent: int) -> None
    
    Returns:
        str: ダウンロードしたファイルのパス、失敗時はNone
    """
    try:
        req = request.Request(
            download_url,
            headers={'User-Agent': 'ProjectWIN-Updater'}
        )
        
        with request.urlopen(req, timeout=300) as response:
            total_size = int(response.headers.get('Content-Length', 0))
            
            # 一時ファイルにダウンロード
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, 'projectwin_update.zip')
            
            with open(temp_path, 'wb') as f:
                downloaded = 0
                block_size = 8192
                
                while True:
                    chunk = response.read(block_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    
                    if progress_callback and total_size > 0:
                        percent = int(downloaded * 100 / total_size)
                        progress_callback(percent)
            
            return temp_path
            
    except Exception as e:
        print(f"Download failed: {e}")
        return None


def apply_update(zip_path, progress_callback=None):
    """
    ダウンロードしたアップデートを適用
    
    Args:
        zip_path: ダウンロードしたzipファイルのパス
        progress_callback: 進捗コールバック (percent: int) -> None
    
    Returns:
        bool: 成功したらTrue
    """
    try:
        app_dir = get_app_dir()
        
        # 保護するフォルダ（ユーザーデータ）
        protected_folders = ['settings', 'cookies', 'exports']
        
        # 一時フォルダに展開
        temp_dir = tempfile.mkdtemp(prefix='projectwin_update_')
        
        with zipfile.ZipFile(zip_path, 'r') as zf:
            file_list = zf.namelist()
            total_files = len(file_list)
            
            for i, file_name in enumerate(file_list):
                # 保護フォルダはスキップ
                skip = False
                for protected in protected_folders:
                    if f'_internal/{protected}/' in file_name or f'_internal\\{protected}\\' in file_name:
                        skip = True
                        break
                    if file_name.startswith(f'{protected}/') or file_name.startswith(f'{protected}\\'):
                        skip = True
                        break
                
                if not skip:
                    zf.extract(file_name, temp_dir)
                
                if progress_callback:
                    percent = int((i + 1) * 50 / total_files)  # 展開は50%まで
                    progress_callback(percent)
        
        # 更新スクリプトを作成
        update_script = app_dir / '_update.bat'
        current_exe = Path(sys.executable).name if getattr(sys, 'frozen', False) else 'ProjectWIN.exe'
        
        # 展開したフォルダの中身を確認（releaseフォルダがある場合）
        extracted_content = temp_dir
        for item in os.listdir(temp_dir):
            item_path = os.path.join(temp_dir, item)
            if os.path.isdir(item_path):
                # サブフォルダがある場合はそこを使う
                extracted_content = item_path
                break
        
        script_content = f'''@echo off
chcp 65001 > nul
echo Updating Project WIN...
timeout /t 2 /nobreak > nul

rem 古いexeを削除（リトライあり）
:del_retry
del /f /q "{app_dir}\\{current_exe}" 2>nul
if exist "{app_dir}\\{current_exe}" (
    timeout /t 1 /nobreak > nul
    goto del_retry
)

rem 新しいファイルをコピー
xcopy /E /Y /I "{extracted_content}\\*" "{app_dir}\\" > nul 2>&1

rem 一時ファイルを削除
rmdir /S /Q "{temp_dir}" 2>nul
del /f /q "{zip_path}" 2>nul

rem アプリを再起動
echo Update complete! Restarting...
timeout /t 1 /nobreak > nul
start "" "{app_dir}\\{current_exe}"

rem このスクリプト自身を削除
(goto) 2>nul & del /f /q "%~f0"
'''
        
        with open(update_script, 'w', encoding='utf-8') as f:
            f.write(script_content)
        
        if progress_callback:
            progress_callback(100)
        
        return True
        
    except Exception as e:
        print(f"Apply update failed: {e}")
        import traceback
        traceback.print_exc()
        return False


def restart_app():
    """アプリを再起動（更新スクリプト経由）"""
    try:
        app_dir = get_app_dir()
        update_script = app_dir / '_update.bat'
        
        if update_script.exists():
            # 更新スクリプトを実行
            if sys.platform == 'win32':
                subprocess.Popen(
                    ['cmd', '/c', str(update_script)],
                    creationflags=subprocess.CREATE_NO_WINDOW | subprocess.DETACHED_PROCESS,
                    close_fds=True,
                    cwd=str(app_dir)
                )
            # アプリを終了
            sys.exit(0)
        else:
            # 更新スクリプトがない場合は通常再起動
            if getattr(sys, 'frozen', False):
                subprocess.Popen([sys.executable], cwd=str(app_dir))
                sys.exit(0)
    except Exception as e:
        print(f"Restart failed: {e}")


if __name__ == "__main__":
    # テスト用
    print(f"Current version: {get_current_version()}")
    needs_update, version, url, changelog = check_for_update()
    print(f"Needs update: {needs_update}")
    print(f"Latest version: {version}")
    print(f"Download URL: {url}")