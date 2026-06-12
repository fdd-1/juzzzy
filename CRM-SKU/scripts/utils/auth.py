# -*- coding: utf-8 -*-
"""
凭据获取：环境变量优先，其次交互式输入。
不再从 config.local.json 读取密码，避免明文落盘。
"""

import os
import sys
import getpass


ENV_USERNAME = "CRM_USERNAME"
ENV_PASSWORD = "CRM_PASSWORD"


def get_credentials(config_auth=None):
    """按优先级获取账号密码。

    1. 环境变量 CRM_USERNAME / CRM_PASSWORD
    2. config.local.json 的 auth.username（**仅账号**，密码不再读取）
    3. 交互式输入（密码用 getpass 隐藏）

    Returns: (username, password)
    Raises : RuntimeError 在非交互式且无凭据时
    """
    config_auth = config_auth or {}

    username = os.environ.get(ENV_USERNAME) or config_auth.get("username") or ""
    password = os.environ.get(ENV_PASSWORD) or ""

    if username:
        print(f"账号：{_mask(username)}（来源：{'env' if os.environ.get(ENV_USERNAME) else 'config'}）")
    else:
        if not sys.stdin.isatty():
            raise RuntimeError(
                f"未提供账号。请设置环境变量 {ENV_USERNAME} / {ENV_PASSWORD}，"
                f"或在交互式终端运行后输入。"
            )
        username = input("CRM 账号: ").strip()
        if not username:
            raise RuntimeError("账号不能为空")

    if not password:
        if not sys.stdin.isatty():
            raise RuntimeError(
                f"未提供密码。请设置环境变量 {ENV_PASSWORD}，"
                f"或在交互式终端运行后输入。"
            )
        password = getpass.getpass("CRM 密码（输入不回显）: ")
        if not password:
            raise RuntimeError("密码不能为空")
    else:
        print("密码：来源 env（已隐藏）")

    return username, password


def _mask(s):
    if not s:
        return ""
    if len(s) <= 4:
        return "*" * len(s)
    return s[:2] + "*" * (len(s) - 4) + s[-2:]
