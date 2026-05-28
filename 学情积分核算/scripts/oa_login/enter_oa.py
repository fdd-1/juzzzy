#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""复用 auth_state.json 进入六一 portal，自动点击 OA 系统瓷砖，进入 OA 后截图。"""
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
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到 {AUTH_PATH}，请先跑 login_oa_screenshot.py 完成登录。", flush=True)
    sys.exit(1)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=120)
    ctx = browser.new_context(storage_state=str(AUTH_PATH))
    page = ctx.new_page()

    print(f"[STEP 1] 打开 portal: {PORTAL_URL}", flush=True)
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(3000)
    print(f"  [URL] {page.url}", flush=True)

    # 先截图当前 portal
    portal_shot = SCRIPT_DIR / "portal_after_reuse.png"
    page.screenshot(path=str(portal_shot), full_page=True)
    print(f"[SHOT] portal 截图: {portal_shot}", flush=True)

    # 检查页面是否显示「选择进入系统」
    has_oa = page.locator("p:has-text('OA系统')").count()
    print(f"[CHECK] OA系统 瓷砖数量: {has_oa}", flush=True)
    if has_oa == 0:
        # 可能登录态过期了，dump html 排查
        (SCRIPT_DIR / "portal_after_reuse.html").write_text(page.content(), encoding="utf-8")
        print("[ERROR] 没看到 OA系统 瓷砖，登录态可能失效。已 dump portal_after_reuse.html", flush=True)
        page.wait_for_timeout(20000)
        browser.close()
        sys.exit(1)

    print("[STEP 2] 点击 OA系统 瓷砖", flush=True)
    # 监听新页面（OA 多半会在新 tab 打开）
    with ctx.expect_page(timeout=30000) as new_page_info:
        page.locator("p:has-text('OA系统')").first.click()
    oa_page = new_page_info.value

    print("[STEP 3] OA 新 tab 已打开，等待加载", flush=True)
    oa_page.wait_for_load_state("domcontentloaded", timeout=30000)
    oa_page.wait_for_timeout(5000)
    print(f"  [OA URL] {oa_page.url}", flush=True)

    # 截图 + dump
    oa_shot = SCRIPT_DIR / "oa_home.png"
    oa_html = SCRIPT_DIR / "oa_home.html"
    oa_page.screenshot(path=str(oa_shot), full_page=True)
    oa_html.write_text(oa_page.content(), encoding="utf-8")
    print(f"[SHOT] OA 首页截图: {oa_shot}", flush=True)
    print(f"[DUMP] OA 首页 HTML : {oa_html}", flush=True)

    # 顺便保存最新 storage state（包含 OA 域 cookie）
    ctx.storage_state(path=str(AUTH_PATH))
    print(f"[OK] 已更新登录态 (含 OA cookie): {AUTH_PATH}", flush=True)

    print("[STAY] 浏览器保持 90 秒方便观察 OA 首页结构。", flush=True)
    oa_page.wait_for_timeout(90000)
    browser.close()
