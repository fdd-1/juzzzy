#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CRM-SKU 首次配置助手

用法：
    python setup.py

会引导用户：
1. 复制 config.example.json -> config.local.json
2. 输入账号密码
3. 检查 Playwright 是否安装
"""

import sys
import io
import json
import shutil
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = Path(__file__).parent
EXAMPLE_FILE = SCRIPT_DIR / "config.example.json"
CONFIG_FILE = SCRIPT_DIR / "config.local.json"


def main():
    print("=" * 60)
    print("CRM-SKU 首次配置")
    print("=" * 60)

    # 1. 复制 config
    if CONFIG_FILE.exists():
        print(f"✓ {CONFIG_FILE.name} 已存在")
        with open(CONFIG_FILE, encoding="utf-8") as f:
            config = json.load(f)
    else:
        if not EXAMPLE_FILE.exists():
            print(f"✗ 找不到 {EXAMPLE_FILE.name}")
            sys.exit(1)
        shutil.copy(EXAMPLE_FILE, CONFIG_FILE)
        print(f"✓ 已创建 {CONFIG_FILE.name}")
        with open(CONFIG_FILE, encoding="utf-8") as f:
            config = json.load(f)

    # 2. 输入账号
    auth = config.setdefault("auth", {})
    if not auth.get("username") or auth.get("username", "").startswith("<"):
        username = input("CRM 账号（手机号）: ").strip()
        auth["username"] = username
    else:
        print(f"✓ 已配置账号: {auth['username']}")

    if not auth.get("password") or auth.get("password", "").startswith("<"):
        password = input("CRM 密码: ").strip()
        auth["password"] = password
    else:
        print(f"✓ 已配置密码: ***")

    # 3. 删除 _comment（不需要）
    config.pop("_comment", None)

    # 4. 保存
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    print(f"✓ 配置已保存: {CONFIG_FILE}")

    # 5. 检查 Playwright
    try:
        import playwright
        print("✓ Playwright 已安装")
    except ImportError:
        print("✗ Playwright 未安装")
        print("  请运行: pip install playwright openpyxl")
        print("  然后: playwright install chromium")
        sys.exit(1)

    print()
    print("=" * 60)
    print("配置完成！下一步：")
    print("=" * 60)
    print("1. 准备 Excel 配置表（参考 ../templates/课包配置模板-美澳.xlsx）")
    print("2. 跑课时包：")
    print('   PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \\')
    print('       --xlsx "<Excel 路径>" --skip-existing --use-password')
    print("3. 跑套餐：")
    print('   PYTHONIOENCODING=utf-8 python crm_batch_create_packages.py \\')
    print('       --xlsx "<Excel 路径>" --use-password')


if __name__ == "__main__":
    main()
