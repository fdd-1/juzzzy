#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""北极星外呼平台 扫码登录

流程：
  1) 打开 LOGIN_URL → 截 before_login.png
  2) 用户扫码确认
  3) 通过 URL 跳转（不再含 passport.vipthink.cn）判断登录成功
  4) 截 after_login.png → 保存 auth_state.json
"""
import sys, io, time
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
AUTH_PATH = SCRIPT_DIR / "auth_state.json"

LOGIN_URL = "https://passport.vipthink.cn/#/account/login?redirectUrl=https%3A%2F%2Fsh-center.vipthink.cn%2F%23%2F"
HOME_URL = "https://sh-center.vipthink.cn/#/"
LOGIN_HOST_KEY = "passport.vipthink.cn"


with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome")
    ctx = browser.new_context()
    page = ctx.new_page()

    print(f"[STEP 1] 打开登录页: {LOGIN_URL}", flush=True)
    page.goto(LOGIN_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(2000)
    print(f"  [URL] {page.url}", flush=True)

    before_shot = SCRIPT_DIR / "before_login.png"
    page.screenshot(path=str(before_shot), full_page=True)
    print(f"[SHOT] 登录前截图: {before_shot}", flush=True)

    print("=" * 60, flush=True)
    print("[ACTION] 请用钉钉扫码登录北极星", flush=True)
    print("[POLL]   通过 URL 离开 passport.vipthink.cn 判断登录成功，最多等 10 分钟", flush=True)
    print("=" * 60, flush=True)

    deadline = time.time() + 600
    logged_in = False
    last_url = page.url
    while time.time() < deadline:
        cur_url = page.url
        if cur_url != last_url:
            print(f"  [POLL] {cur_url}", flush=True)
            last_url = cur_url
        if LOGIN_HOST_KEY not in cur_url.lower():
            logged_in = True
            print("[DETECT] URL 已离开登录域，判定登录成功", flush=True)
            break
        page.wait_for_timeout(2000)

    if not logged_in:
        print(f"[ERROR] 10 分钟内未检测到登录成功，当前 URL: {page.url}", flush=True)
        page.screenshot(path=str(SCRIPT_DIR / "login_timeout.png"), full_page=True)
        browser.close()
        sys.exit(1)

    print(f"[STEP 2] 登录成功，当前 URL: {page.url}", flush=True)
    page.wait_for_timeout(3000)

    # 强制再访问 home 确认登录态有效
    print(f"[STEP 3] 强制跳转 home: {HOME_URL}", flush=True)
    page.goto(HOME_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(3000)
    print(f"  [URL] {page.url}", flush=True)

    if LOGIN_HOST_KEY in page.url.lower():
        print(f"[ERROR] 跳到 home 后又被踢回登录页: {page.url}", flush=True)
        browser.close()
        sys.exit(2)

    after_shot = SCRIPT_DIR / "after_login.png"
    page.screenshot(path=str(after_shot), full_page=True)
    print(f"[SHOT] 登录后截图: {after_shot}", flush=True)

    home_html = SCRIPT_DIR / "home.html"
    home_html.write_text(page.content(), encoding="utf-8")
    print(f"[DUMP] home HTML: {home_html}", flush=True)

    ctx.storage_state(path=str(AUTH_PATH))
    print(f"[OK]   登录态已保存: {AUTH_PATH}", flush=True)
    print("[DONE] 后续可用 auth_state.json 直接复用", flush=True)

    print("[STAY] 浏览器保持 20 秒方便确认", flush=True)
    page.wait_for_timeout(20000)
    browser.close()
