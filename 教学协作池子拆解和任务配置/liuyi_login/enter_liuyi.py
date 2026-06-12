#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""复用 auth_state.json 进入六一 portal，自动点击「六一工作台」瓷砖。
   入口确认：
     - portal 加载成功
     - portal_home.png 截图保存
     - 点击「六一工作台」→ 新 tab 加载完成
     - liuyi_home.png 截图保存
"""
import sys, io
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
AUTH_PATH  = SCRIPT_DIR / "auth_state.json"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到 {AUTH_PATH}，请先跑 login_liuyi.py 完成登录。", flush=True)
    sys.exit(1)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=120)
    ctx = browser.new_context(storage_state=str(AUTH_PATH))
    page = ctx.new_page()

    print(f"[STEP 1] 打开 portal: {PORTAL_URL}", flush=True)
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")

    # 登录态判定：等「六一工作台」瓷砖出现，最多等 30 秒
    # 注意 portal 入口会先经过 login.61info.cn 的中转跳转，URL 不可靠，元素出现才是真信号
    print("[STEP 1.5] 等待 portal 瓷砖渲染（最多 30s）", flush=True)
    try:
        page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)
    except Exception:
        print(f"[ERROR] 30s 内未渲染出「六一工作台」瓷砖，当前 URL: {page.url}", flush=True)
        page.screenshot(path=str(SCRIPT_DIR / "portal_expired.png"), full_page=True)
        (SCRIPT_DIR / "portal_expired.html").write_text(page.content(), encoding="utf-8")
        if "login.61info.cn" in page.url.lower():
            print("[HINT] 当前停在登录页，登录态可能失效，重跑 login_liuyi.py", flush=True)
        browser.close()
        sys.exit(2)
    print(f"  [URL] {page.url}", flush=True)

    portal_shot = SCRIPT_DIR / "portal_after_reuse.png"
    page.screenshot(path=str(portal_shot), full_page=True)
    print(f"[SHOT] portal 截图: {portal_shot}", flush=True)
    print("[CHECK] 六一工作台 瓷砖已渲染", flush=True)

    print("[STEP 2] 点击「六一工作台」瓷砖", flush=True)
    with ctx.expect_page(timeout=30000) as new_page_info:
        page.locator("p:has-text('六一工作台')").first.click()
    liuyi_page = new_page_info.value

    print("[STEP 3] 六一工作台新 tab 已打开，等待加载", flush=True)
    liuyi_page.wait_for_load_state("domcontentloaded", timeout=30000)
    liuyi_page.wait_for_timeout(5000)
    print(f"  [LIUYI URL] {liuyi_page.url}", flush=True)

    liuyi_shot = SCRIPT_DIR / "liuyi_home.png"
    liuyi_html = SCRIPT_DIR / "liuyi_home.html"
    liuyi_page.screenshot(path=str(liuyi_shot), full_page=True)
    liuyi_html.write_text(liuyi_page.content(), encoding="utf-8")
    print(f"[SHOT] 六一工作台首页截图: {liuyi_shot}", flush=True)
    print(f"[DUMP] 六一工作台首页 HTML: {liuyi_html}", flush=True)

    ctx.storage_state(path=str(AUTH_PATH))
    print(f"[OK] 已更新登录态 (含六一工作台域 cookie): {AUTH_PATH}", flush=True)

    print("[STAY] 浏览器保持 5 秒确认子域 cookie 同步完成。", flush=True)
    liuyi_page.wait_for_timeout(5000)
    # 再次保存 storage_state，确保六一工作台所有子域 cookie 都落盘
    ctx.storage_state(path=str(AUTH_PATH))
    print(f"[OK] 已二次保存 auth_state.json", flush=True)
    browser.close()
