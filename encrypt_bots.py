"""
Botファイル＆コアモジュール暗号化スクリプト
GitHubにアップロードする前にこのスクリプトで暗号化します

使い方:
    python encrypt_bots.py

これにより:
- _internal/bots/フォルダ内の.pyファイルが暗号化され、encrypted_bots/bots/に出力
- license_manager.pyとupdater.pyが暗号化され、encrypted_bots/core/に出力
"""

import os
from pathlib import Path

# GUI.pyと同じ暗号化キー（難読化前の生キー）
# GUI.py側では難読化されているが、暗号化時は元のキーを使用
BOT_ENCRYPTION_KEY = b'ProjectWIN_Bot_Key_2024_Secure!!'  # 32 bytes
CORE_ENCRYPTION_KEY = b'ProjectWIN_Core_Key_2024_Secure!'  # 31 bytes

# 暗号化対象のBotファイル（新しいパス構造）
BOT_FILES = [
    "_internal/bots/amazon/amazon_card.py",
    "_internal/bots/amazon/amazon_raffle.py",
    "_internal/bots/amazon/amazon_addy.py",
    "_internal/bots/amazon/amazon_browser.py",
    "_internal/bots/amazon/amazon_signup.py",
    "_internal/bots/icloud/icloud_collect.py",
    "_internal/bots/icloud/icloud_generate.py",
]

# 暗号化対象のコアモジュール
CORE_FILES = [
    "license_manager.py",
    "updater.py",
]


def encrypt_data(data: bytes, key: bytes) -> bytes:
    """XOR暗号化"""
    encrypted = bytearray()
    for i, byte in enumerate(data):
        encrypted.append(byte ^ key[i % len(key)])
    return bytes(encrypted)


def main():
    base_dir = Path(__file__).parent
    output_dir = base_dir / "encrypted_bots"
    
    print("=" * 50)
    print("Bot & コアモジュール暗号化スクリプト")
    print("=" * 50)
    
    # ----- Botファイルの暗号化 -----
    print("\n【Botファイル】")
    for bot_path in BOT_FILES:
        src = base_dir / bot_path
        
        if not src.exists():
            print(f"  SKIP: {bot_path} (not found)")
            continue
        
        # 出力先: _internal/bots/amazon/xxx.py -> encrypted_bots/bots/amazon/xxx.enc
        # _internal/を除去して出力
        relative_path = Path(bot_path)
        # "_internal/bots/amazon/xxx.py" -> "bots/amazon/xxx.py"
        output_relative = Path(*relative_path.parts[1:])  # _internalを除去
        
        dst_dir = output_dir / output_relative.parent
        dst_dir.mkdir(parents=True, exist_ok=True)
        
        # 暗号化
        with open(src, 'rb') as f:
            original_data = f.read()
        
        encrypted_data = encrypt_data(original_data, BOT_ENCRYPTION_KEY)
        
        # .encとして保存
        dst = dst_dir / (output_relative.stem + ".enc")
        with open(dst, 'wb') as f:
            f.write(encrypted_data)
        
        print(f"  OK: {bot_path} -> {dst.relative_to(base_dir)}")
    
    # ----- コアモジュールの暗号化 -----
    print("\n【コアモジュール】")
    core_dir = output_dir / "core"
    core_dir.mkdir(parents=True, exist_ok=True)
    
    for core_file in CORE_FILES:
        src = base_dir / core_file
        
        if not src.exists():
            print(f"  SKIP: {core_file} (not found)")
            continue
        
        # 暗号化
        with open(src, 'rb') as f:
            original_data = f.read()
        
        encrypted_data = encrypt_data(original_data, CORE_ENCRYPTION_KEY)
        
        # .encとして保存
        dst = core_dir / (Path(core_file).stem + ".enc")
        with open(dst, 'wb') as f:
            f.write(encrypted_data)
        
        print(f"  OK: {core_file} -> {dst.relative_to(base_dir)}")
    
    print()
    print("=" * 50)
    print("暗号化完了!")
    print("=" * 50)
    print()
    print(f"出力先: {output_dir}")
    print()
    print("GitHubへのアップロード:")
    print("  projectwin-assets リポジトリに以下の構造でアップロード:")
    print()
    print("  projectwin-assets/")
    print("  ├── bots/")
    print("  │   ├── amazon/")
    print("  │   │   ├── amazon_signup.enc")
    print("  │   │   ├── amazon_raffle.enc")
    print("  │   │   └── ...")
    print("  │   └── icloud/")
    print("  │       ├── icloud_collect.enc")
    print("  │       └── icloud_generate.enc")
    print("  ├── core/")
    print("  │   ├── license_manager.enc")
    print("  │   └── updater.enc")
    print("  ├── task.png")
    print("  └── ...")


if __name__ == "__main__":
    main()