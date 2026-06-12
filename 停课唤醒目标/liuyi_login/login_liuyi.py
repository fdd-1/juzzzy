#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""六一工作台 扫码登录

流程：
  1) 打开 portal → 截 before_login.png（此时是登录页 / 二维码）
  2) 用户扫码确认
  3) 脚本通过两个信号判断登录成功：
       a. URL 不再含 login.61info.cn
       b. 页面出现「六一工作台 / 选择进入系统」等瓷砖文字
  4) 截 after_login.png 留证 → 保存 auth_state.json
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
AUTH_PATH  = SCRIPT_DIR / "auth_state.json"
LOGIN_URL  = "https://login.61info.cn/"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

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
    print("[ACTION] 请用钉钉扫码登录（页面右侧二维码）", flush=True)
    print("[POLL]   脚本会通过 URL 跳转 + 页面出现「六一工作台」瓷砖来判断登录，最多等 10 分钟", flush=True)
    print("=" * 60, flush=True)

    deadline = time.time() + 600
    logged_in = False
    last_url = page.url
    while time.time() < deadline:
        cur_url = page.url
        if cur_url != last_url:
            print(f"  [POLL] {cur_url}", flush=True)
            last_url = cur_url

        try:
            has_select_text = page.locator("text=选择进入系统").count() > 0
            has_liuyi_tile  = page.locator("p:has-text('六一工作台')").count() > 0
            has_cms_tile    = page.locator("p:has-text('CMS系统')").count() > 0
            if has_select_text or has_liuyi_tile or has_cms_tile:
                logged_in = True
                print("[DETECT] 页面已出现 portal 瓷砖，判定登录成功", flush=True)
                break
        except Exception:
            pass

        if "login.61info.cn" not in cur_url.lower():
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

    if page.locator("p:has-text('六一工作台')").count() == 0:
        print(f"[STEP 3] 强制跳转 portal: {PORTAL_URL}", flush=True)
        page.goto(PORTAL_URL, timeout=30000)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(3000)
        print(f"  [URL] {page.url}", flush=True)

        if "login.61info.cn" in page.url.lower():
            print(f"[ERROR] 跳到 portal 后又回到登录页: {page.url}", flush=True)
            browser.close()
            sys.exit(2)

    after_shot = SCRIPT_DIR / "after_login.png"
    page.screenshot(path=str(after_shot), full_page=True)
    print(f"[SHOT] 登录后截图: {after_shot}", flush=True)

    portal_html = SCRIPT_DIR / "portal_home.html"
    portal_html.write_text(page.content(), encoding="utf-8")
    print(f"[DUMP] portal 首页 HTML: {portal_html}", flush=True)

    has_liuyi = page.locator("p:has-text('六一工作台')").count()
    print(f"[CHECK] 六一工作台 瓷砖数量: {has_liuyi}", flush=True)
    if has_liuyi == 0:
        print("[WARN] 没看到「六一工作台」瓷砖，请人工核对 after_login.png 是否真的登录成功", flush=True)

    ctx.storage_state(path=str(AUTH_PATH))
    print(f"[OK]   登录态已保存: {AUTH_PATH}", flush=True)
    print("[DONE] 后续可用 enter_liuyi.py 直接进六一工作台。", flush=True)

    print("[STAY] 浏览器保持 20 秒方便确认。", flush=True)
    page.wait_for_timeout(20000)
    browser.close()
