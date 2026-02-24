"""
Nuitka ビルドスクリプト

PyInstallerの代わりにNuitkaでビルドする
- .pycファイルが生成されない（Cにコンパイル済み）
- 逆コンパイルが非常に困難
"""
import subprocess
import sys
import os
import shutil
import json
from pathlib import Path

# ----- 設定 -----
MAIN_SCRIPT = "GUI.py"
APP_NAME = "ProjectWIN"
VERSION_FILE = "version.json"

def step(msg):
    print(f"\n{'='*50}\n{msg}\n{'='*50}")

def main():
    base_dir = Path(__file__).parent
    
    # バージョンを読み込む
    version = "1.0.0"
    version_path = base_dir / VERSION_FILE
    if version_path.exists():
        with open(version_path, 'r') as vf:
            version = json.load(vf).get("version", "1.0.0")
    
    step(f"Nuitka Build: {APP_NAME} v{version}")
    
    # dist_nuitkaフォルダを作成
    dist_dir = base_dir / "dist_nuitka"
    dist_dir.mkdir(exist_ok=True)
    
    # ----- Step 1: Nuitka でビルド -----
    step("Step 1: Nuitka でビルド中...")
    
    cmd = [
        sys.executable, "-m", "nuitka",
        
        # 基本設定
        "--standalone",                    # 単独実行可能な形式（フォルダ形式）
        # "--onefile",                     # ← 削除：一時フォルダに展開されてパスがおかしくなる
        f"--output-dir={dist_dir}",
        f"--output-filename={APP_NAME}.exe",
        
        # Windows設定
        "--windows-console-mode=attach",   # コンソールをアタッチ（デバッグ用）
        "--windows-icon-from-ico=icon.ico" if (base_dir / "icon.ico").exists() else "",
        
        # 最適化
        "--follow-imports",
        "--assume-yes-for-downloads",
        
        # 必要なモジュールを含める
        "--include-module=PySide6",
        "--include-module=requests",
        "--include-module=PIL",
        "--include-module=numpy",
        "--include-module=cryptography",
        "--include-module=playwright",
        "--include-module=playwright.sync_api",
        "--include-module=playwright._impl",
        "--include-module=greenlet",
        "--include-module=bs4",
        "--include-module=soupsieve",
        "--include-module=lxml",
        "--include-module=imaplib",
        "--include-module=email",
        "--include-module=urllib3",
        "--include-module=certifi",
        "--include-module=charset_normalizer",
        "--include-module=idna",
        
        # プラグイン
        "--enable-plugin=pyside6",
        
        # メインスクリプト
        str(base_dir / MAIN_SCRIPT),
    ]
    
    # 空文字列を除去
    cmd = [c for c in cmd if c]
    
    print("Running:", " ".join(cmd[:5]), "...")
    result = subprocess.run(cmd, cwd=base_dir)
    
    if result.returncode != 0:
        print("\nERROR: Nuitka build failed!")
        print("エラーが発生した場合は、以下を確認してください:")
        print("  1. Visual Studio Build Tools がインストールされているか")
        print("  2. pip install nuitka が完了しているか")
        print("  3. 必要なライブラリがインストールされているか")
        sys.exit(1)
    
    # ----- Step 2: version.json をコピー -----
    step("Step 2: 追加ファイルをコピー中...")
    
    # standaloneモードの場合: GUI.dist フォルダが作成される
    gui_dist_dir = base_dir / "dist_nuitka" / f"{MAIN_SCRIPT.replace('.py', '.dist')}"
    
    if not gui_dist_dir.exists():
        # フォルダ名が違う可能性があるので探す
        for item in (base_dir / "dist_nuitka").iterdir():
            if item.is_dir() and item.name.endswith('.dist'):
                gui_dist_dir = item
                break
    
    if gui_dist_dir.exists():
        # _internalフォルダを作成してversion.jsonを配置
        internal_dir = gui_dist_dir / "_internal"
        internal_dir.mkdir(exist_ok=True)
        
        if (base_dir / VERSION_FILE).exists():
            shutil.copy2(base_dir / VERSION_FILE, internal_dir / VERSION_FILE)
            print(f"  OK: {VERSION_FILE} -> {gui_dist_dir.name}/_internal/")
        
        # exeの名前を変更（GUI.exe → ProjectWIN.exe）
        old_exe = gui_dist_dir / "GUI.exe"
        new_exe = gui_dist_dir / f"{APP_NAME}.exe"
        if old_exe.exists() and not new_exe.exists():
            old_exe.rename(new_exe)
            print(f"  OK: GUI.exe -> {APP_NAME}.exe")
    else:
        print("  WARNING: dist folder not found!")
    
    # ----- Step 3: ProjectWINフォルダにコピー -----
    step("Step 3: 最終フォルダを作成中...")
    
    final_dir = base_dir / "dist_nuitka" / APP_NAME
    if final_dir.exists():
        shutil.rmtree(final_dir)
    
    if gui_dist_dir.exists():
        shutil.copytree(gui_dist_dir, final_dir)
        print(f"  OK: {gui_dist_dir.name} -> {APP_NAME}/")
        
        # GUI.distフォルダを削除
        shutil.rmtree(gui_dist_dir)
        print(f"  OK: {gui_dist_dir.name} を削除しました")
    
    # ----- Step 4: 不要なフォルダを削除 -----
    step("Step 4: 不要なフォルダを削除中...")
    
    # GUI.buildフォルダを削除（Nuitkaのキャッシュ）
    gui_build_dir = base_dir / "GUI.build"
    if gui_build_dir.exists():
        shutil.rmtree(gui_build_dir)
        print(f"  OK: GUI.build を削除しました")
    
    # dist_nuitka内の不要なフォルダを削除
    for item in (base_dir / "dist_nuitka").iterdir():
        if item.is_dir() and item.name != APP_NAME:
            shutil.rmtree(item)
            print(f"  OK: dist_nuitka/{item.name} を削除しました")
    
    # ----- 完了 -----
    step("ビルド完了!")
    print(f"""
出力先: dist_nuitka/{APP_NAME}/

【PyInstallerとの違い】
  - .pycファイルが存在しない（Cにコンパイル済み）
  - 逆コンパイルが非常に困難
  - 実行速度が向上する可能性あり

【注意】
  - 初回起動時、Windowsセキュリティの警告が出る場合があります
  - 動作確認を十分に行ってください

【配布用zipの作成方法】
  cd dist_nuitka
  Compress-Archive -Path "ProjectWIN" -DestinationPath "ProjectWIN_v{version}.zip" -Force
""")


if __name__ == "__main__":
    main()