"""
exe化 + 配布用zip作成スクリプト
PysideGUI フォルダ内で実行してください:
    python build.py
"""

import subprocess
import sys
import os
import shutil
import py_compile
import json
from pathlib import Path

# ===== 設定 =====
APP_NAME = "ProjectWIN"
MAIN_SCRIPT = "GUI.py"
ICON_FILE = None  # アイコンがあれば "icon.ico" を指定
VERSION_FILE = "version.json"

# コンパイル対象の.pyファイル（テスターに見せたくないもの）
# ※ すべてサーバーから取得するためビルドに含めない
PYC_TARGETS = [
    # コアモジュールはサーバーからダウンロードするため不要
    # "license_manager.py",
    # "updater.py",
    # botファイルもサーバーからダウンロードするため不要
    # "bots/amazon/amazon_card.py",
    # "bots/amazon/amazon_raffle.py",
    # "bots/amazon/amazon_addy.py",
    # "bots/amazon/amazon_browser.py",
    # "bots/amazon/amazon_signup.py",
]

# 配布に含めるフォルダ（空でも作成）- 自動作成されるため不要
DIST_FOLDERS = [
    # "cookies",   # amazon_signup.py で自動作成される
    # "settings",  # GUI.py で自動作成される
]

# 配布に含める個別ファイル（_internal内に配置）
DIST_FILES_INTERNAL = [
    "version.json",
]

# 配布に含める個別ファイル（exeと同じ場所に配置）
DIST_FILES = [
    # "version.json",  # _internalに移動
]


def get_version():
    """version.json からバージョンを取得"""
    try:
        with open(VERSION_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("version", "1.0.0")
    except (FileNotFoundError, json.JSONDecodeError):
        return "1.0.0"


def step(msg):
    print(f"\n{'='*50}")
    print(f"  {msg}")
    print(f"{'='*50}")


def main():
    base_dir = Path(__file__).parent
    os.chdir(base_dir)
    
    # バージョン取得
    version = get_version()
    print(f"\n*** Building {APP_NAME} v{version} ***\n")
    
    # ----- Step 1: PyInstaller で exe 化 -----
    step("Step 1: GUI.py を exe 化中...")
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",          # 1フォルダにまとめる
        "--console",        # コンソール表示
        "--name", APP_NAME,
        "--hidden-import", "PySide6.QtWidgets",
        "--hidden-import", "PySide6.QtCore",
        "--hidden-import", "PySide6.QtGui",
        "--hidden-import", "openpyxl",
        "--hidden-import", "requests",
        "--hidden-import", "license_manager",
        "--hidden-import", "updater",
        # botが使うライブラリ
        "--hidden-import", "playwright",
        "--hidden-import", "playwright.sync_api",
        "--hidden-import", "bs4",
        "--hidden-import", "imaplib",
        "--hidden-import", "email",
        "--collect-all", "playwright",
    ]
    
    if ICON_FILE and Path(ICON_FILE).exists():
        cmd.extend(["--icon", ICON_FILE])
    
    cmd.append(MAIN_SCRIPT)
    
    result = subprocess.run(cmd, capture_output=False)
    if result.returncode != 0:
        print("ERROR: PyInstaller failed!")
        sys.exit(1)
    
    print("exe 化完了!")
    
    # ----- Step 2: .py → .pyc コンパイル -----
    step("Step 2: bot ファイルをコンパイル中...")
    
    dist_dir = base_dir / "dist" / APP_NAME
    internal_dir = dist_dir / "_internal"
    
    for target in PYC_TARGETS:
        src = base_dir / target
        if not src.exists():
            print(f"  SKIP: {target} (not found)")
            continue
        
        # botsフォルダのファイルは_internal/botsに配置
        if target.startswith("bots/"):
            dst_dir = internal_dir / Path(target).parent
        else:
            dst_dir = dist_dir
        
        dst_dir.mkdir(parents=True, exist_ok=True)
        
        pyc_path = dst_dir / (Path(target).stem + ".pyc")
        py_compile.compile(str(src), cfile=str(pyc_path), doraise=True)
        print(f"  OK: {target} -> {pyc_path.relative_to(dist_dir)}")
    
    # ----- Step 3: フォルダ構成を作成 -----
    step("Step 3: 配布用フォルダ構成を作成中...")
    
    for folder in DIST_FOLDERS:
        folder_path = dist_dir / folder
        folder_path.mkdir(parents=True, exist_ok=True)
        # 空フォルダをgitで管理できるように.gitkeepを置く
        gitkeep = folder_path / ".gitkeep"
        if not any(folder_path.iterdir()):
            gitkeep.touch()
        print(f"  OK: {folder}/")
    
    # 個別ファイルをコピー（exeと同じ場所）
    for file_name in DIST_FILES:
        src = base_dir / file_name
        if src.exists():
            shutil.copy2(str(src), str(dist_dir / file_name))
            print(f"  OK: {file_name}")
    
    # 個別ファイルをコピー（_internal内）
    for file_name in DIST_FILES_INTERNAL:
        src = base_dir / file_name
        if src.exists():
            shutil.copy2(str(src), str(internal_dir / file_name))
            print(f"  OK: _internal/{file_name}")
    
    # ----- Step 4: 不要な .py ファイルを削除 -----
    step("Step 4: 配布フォルダから .py ソースを除去中...")
    
    # dist内にコピーされた.pyファイルを削除（exeに含まれているので不要）
    for py_file in dist_dir.glob("*.py"):
        py_file.unlink()
        print(f"  REMOVED: {py_file.name}")
    
    # _internal/bots/ 内の .py も削除（.pyc だけ残す）
    bots_dir = internal_dir / "bots"
    if bots_dir.exists():
        for py_file in bots_dir.rglob("*.py"):
            py_file.unlink()
            print(f"  REMOVED: {py_file.relative_to(dist_dir)}")
    
    # ----- Step 5: zip 作成 -----
    step("Step 5: 配布用 zip を作成中...")
    
    # バージョン付きのzip名を自動生成
    zip_name = f"{APP_NAME}_v{version}"
    zip_path = base_dir / zip_name
    shutil.make_archive(str(zip_path), "zip", base_dir / "dist", APP_NAME)
    print(f"\n配布用 zip: {zip_path}.zip")
    
    # ----- 完了 -----
    step("ビルド完了!")
    print(f"""
配布物: {zip_path}.zip

【GitHub Release アップロード手順】
  1. GitHub の Releases → Create new release
  2. Tag: v{version}
  3. {zip_name}.zip をアップロード
  4. Publish release

【Railway 環境変数の更新】
  APP_VERSION = {version}
  UPDATE_DOWNLOAD_URL = https://github.com/TestRaffle/app-releases/releases/download/v{version}/{zip_name}.zip

フォルダ構成:
  {APP_NAME}/
  ├── {APP_NAME}.exe     ← 起動ファイル
  ├── _internal/          ← ランタイム + bots（ユーザーは触らない）
  │   └── bots/amazon/    ← botファイル（.pyc）
  ├── cookies/            ← ユーザーデータ（移行可能）
  ├── settings/           ← ユーザーデータ（移行可能）
  └── version.json
""")


if __name__ == "__main__":
    main()