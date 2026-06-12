#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""六一 OA 扫码登录：打开页面 → 用户扫码 → 检测登录成功 → 保存 auth_state.json"""
import sys, io, time
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    # 关闭 stdout 缓冲，方便后台模式实时观察
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
AUTH_PATH  = SCRIPT_DIR / "auth_state.json"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome")
    ctx = browser.new_context()
    page = ctx.new_page()

    print(f"[STEP 1] 打开 portal: {PORTAL_URL}", flush=True)
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(1500)
    print(f"  [URL] {page.url}", flush=True)

    print("=" * 60, flush=True)
    print("[ACTION] 请在右侧用钉钉扫码登录（或左侧账号密码 + 短信也行）", flush=True)
    print("[POLL]   登录成功后脚本会自动检测并保存登录态，最多等 10 分钟", flush=True)
    print("=" * 60, flush=True)

    deadline = time.time() + 600
    logged_in = False
    last_url = page.url
    while time.time() < deadline:
        cur_url = page.url
        if cur_url != last_url:
            print(f"  [POLL] {cur_url}", flush=True)
            last_url = cur_url
        # 检测页面内容判断登录成功（出现"选择进入系统"或系统瓷砖）
        try:
            has_system_select = page.locator("text=选择进入系统").count() > 0
            has_oa_tile = page.locator("text=OA系统").count() > 0
            has_portal_content = page.locator("text=CMS系统").count() > 0
            if has_system_select or has_oa_tile or has_portal_content:
                logged_in = True
                break
        except Exception:
            pass
        # 也保留 URL 判断作为兜底
        if "login.61info.cn" not in cur_url.lower():
            logged_in = True
            break
        page.wait_for_timeout(2000)

    if not logged_in:
        print(f"[ERROR] 10 分钟内未检测到登录跳转，当前 URL: {page.url}", flush=True)
        page.wait_for_timeout(15000)
        browser.close()
        sys.exit(1)

    print(f"[STEP 2] 登录成功，当前 URL: {page.url}", flush=True)
    page.wait_for_timeout(3000)

    # 检查当前页面是否已经是系统选择页（有 OA系统 瓷砖）
    if page.locator("text=OA系统").count() > 0:
        print(f"[STEP 3] 当前已在系统选择页，直接保存登录态", flush=True)
    else:
        # 如果不是，尝试访问 portal 主页
        print(f"[STEP 3] 访问 portal 验证: {PORTAL_URL}", flush=True)
        page.goto(PORTAL_URL, timeout=30000)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(3000)
        print(f"  [URL] {page.url}", flush=True)

        if "login.61info.cn" in page.url.lower():
            print(f"[ERROR] 跳到 portal 后又被重定向到登录页: {page.url}", flush=True)
            browser.close()
            sys.exit(1)

    # 截图 + dump portal 首页
    shot = SCRIPT_DIR / "portal_home.png"
    page.screenshot(path=str(shot), full_page=True)
    print(f"[SHOT] portal 首页截图: {shot}", flush=True)

    portal_html = SCRIPT_DIR / "portal_home.html"
    portal_html.write_text(page.content(), encoding="utf-8")
    print(f"[DUMP] portal 首页 HTML: {portal_html}", flush=True)

    ctx.storage_state(path=str(AUTH_PATH))
    print(f"[OK]   登录态已保存到: {AUTH_PATH}", flush=True)
    print("[DONE] 后续操作可直接复用，无需重新登录。", flush=True)

    print("[STAY] 浏览器再保持 60 秒，方便你确认 / 我观察 portal 结构。", flush=True)
    page.wait_for_timeout(60000)
    browser.close()
