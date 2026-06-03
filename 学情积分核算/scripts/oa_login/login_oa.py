#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""六一 OA 登录：
   1) 脚本启动 → 自动填账号密码
   2) 用户在浏览器里点击「获取验证码」（等待短信）
   3) 用户把 6 位验证码写到 scripts/oa_login/code.txt （或我代写）
   4) 脚本检测到 code.txt → 自动填验证码 → 点登录 → 保存 auth_state.json
"""
import sys, io, time
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = Path(__file__).parent
AUTH_PATH  = SCRIPT_DIR / "auth_state.json"
CODE_FILE  = SCRIPT_DIR / "code.txt"

PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"
USERNAME   = "15113986797"
PASSWORD   = "Qiou76$y"

# 进入前先清空旧验证码文件
if CODE_FILE.exists():
    CODE_FILE.unlink()
    print(f"[INIT] 已清理旧的 {CODE_FILE.name}")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=80)
    ctx = browser.new_context()
    page = ctx.new_page()

    print(f"[STEP 1] 打开 portal: {PORTAL_URL}")
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(1500)
    print(f"  [URL] {page.url}")

    print("[STEP 2] 自动填入账号、密码")
    page.fill("input[name='account']", USERNAME)
    page.wait_for_timeout(300)
    page.fill("input[name='password']", PASSWORD)
    page.wait_for_timeout(300)

    print("=" * 60)
    print(f"[ACTION] 请在浏览器里点击「获取验证码」按钮，等待 {USERNAME} 短信")
    print(f"[WAIT]   收到验证码后，把 6 位数字写进文件：{CODE_FILE}")
    print(f"         （用户可以告诉我验证码，由我代写）")
    print("[POLL]   脚本每 2 秒轮询一次该文件，最多等 10 分钟")
    print("=" * 60)

    # 轮询等待验证码文件
    code = None
    deadline = time.time() + 600
    while time.time() < deadline:
        if CODE_FILE.exists():
            txt = CODE_FILE.read_text(encoding="utf-8").strip()
            # 提取里面的纯数字
            digits = "".join(ch for ch in txt if ch.isdigit())
            if len(digits) >= 4:
                code = digits[:6]
                print(f"[GOT] 检测到验证码: {code}")
                break
        page.wait_for_timeout(2000)

    if not code:
        print("[ERROR] 10 分钟内没有收到验证码，退出。")
        browser.close()
        sys.exit(1)

    print("[STEP 3] 自动填入验证码并点击登录")
    page.fill("input[name='code']", code)
    page.wait_for_timeout(300)
    page.click("button.login-btn")
    print("[CLICK] 已点击登录按钮，等待跳转")

    # 轮询登录成功
    logged_in = False
    last_url  = page.url
    deadline  = time.time() + 60
    while time.time() < deadline:
        cur_url = page.url
        if cur_url != last_url:
            print(f"  [POLL] {cur_url}")
            last_url = cur_url
        if "login.61info.cn" not in cur_url.lower():
            logged_in = True
            break
        page.wait_for_timeout(1500)

    if not logged_in:
        print(f"[ERROR] 60s 内未跳出登录页，当前 URL: {page.url}")
        # dump 错误页结构
        err_html = SCRIPT_DIR / "login_error.html"
        err_html.write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(SCRIPT_DIR / "login_error.png"), full_page=True)
        print(f"[DUMP] 错误页已保存: login_error.html / login_error.png")
        page.wait_for_timeout(15000)
        browser.close()
        sys.exit(1)

    print(f"[STEP 4] 登录成功，当前 URL: {page.url}")
    page.wait_for_timeout(3000)

    print(f"[STEP 5] 强制跳转到 portal 验证: {PORTAL_URL}")
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(3000)
    print(f"  [URL] {page.url}")

    if "login.61info.cn" in page.url.lower():
        print(f"[ERROR] 跳到 portal 后又回到登录页: {page.url}")
        browser.close()
        sys.exit(1)

    # 截图 + dump portal 首页
    shot = SCRIPT_DIR / "portal_home.png"
    page.screenshot(path=str(shot), full_page=True)
    print(f"[SHOT] portal 首页截图: {shot}")

    portal_html = SCRIPT_DIR / "portal_home.html"
    portal_html.write_text(page.content(), encoding="utf-8")
    print(f"[DUMP] portal 首页 HTML: {portal_html}")

    ctx.storage_state(path=str(AUTH_PATH))
    print(f"[OK]   登录态已保存到: {AUTH_PATH}")
    print("[DONE] 后续可直接复用，无需重新登录。")

    print("[STAY] 浏览器保持 30 秒方便确认。")
    page.wait_for_timeout(30000)
    browser.close()
