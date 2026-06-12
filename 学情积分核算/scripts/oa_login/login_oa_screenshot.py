#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""六一 OA 扫码登录（截图驱动版）：
   - 打开页面后，每 3 秒截图 + dump HTML 一次（覆盖式），文件名固定方便外部观察
   - 通过文件信号 done.txt 触发：用户/外部告诉脚本「我已经登录好了」，脚本立即抓取登录态并保存
   - 之所以这么做，是因为登录前后 URL 可能不变，需要靠人工 / 视觉判断
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
DONE_FILE  = SCRIPT_DIR / "done.txt"
SHOT_FILE  = SCRIPT_DIR / "current_page.png"
HTML_FILE  = SCRIPT_DIR / "current_page.html"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

# 进入前先清空旧信号
for f in (DONE_FILE, AUTH_PATH):
    if f.exists():
        f.unlink()
        print(f"[INIT] 清理旧文件 {f.name}", flush=True)

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
    print("[ACTION] 请在浏览器里完成登录（钉钉扫码 / 账号密码）", flush=True)
    print(f"[SHOT]   每 3 秒覆盖式截图: {SHOT_FILE.name}", flush=True)
    print(f"[DONE]   登录完成后，往 {DONE_FILE.name} 写任意内容，脚本立即保存登录态", flush=True)
    print("=" * 60, flush=True)

    deadline = time.time() + 900  # 15 分钟
    tick = 0
    while time.time() < deadline:
        # 截图 + dump
        try:
            page.screenshot(path=str(SHOT_FILE), full_page=False)
            HTML_FILE.write_text(page.content(), encoding="utf-8")
        except Exception as e:
            print(f"  [WARN] 截图失败: {e}", flush=True)
        tick += 1
        if tick % 3 == 1:
            print(f"  [TICK {tick}] url={page.url}  (已截图)", flush=True)

        # 完成信号
        if DONE_FILE.exists():
            print("[GOT] 检测到 done.txt，准备保存登录态", flush=True)
            break

        page.wait_for_timeout(3000)
    else:
        print("[ERROR] 15 分钟内没收到 done.txt，退出。", flush=True)
        browser.close()
        sys.exit(1)

    # 再访问 portal 确认
    print(f"[STEP 2] 访问 portal 二次确认: {PORTAL_URL}", flush=True)
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(3500)
    print(f"  [URL] {page.url}", flush=True)

    final_shot = SCRIPT_DIR / "portal_home.png"
    final_html = SCRIPT_DIR / "portal_home.html"
    page.screenshot(path=str(final_shot), full_page=True)
    final_html.write_text(page.content(), encoding="utf-8")
    print(f"[SHOT] portal 截图: {final_shot}", flush=True)
    print(f"[DUMP] portal HTML: {final_html}", flush=True)

    ctx.storage_state(path=str(AUTH_PATH))
    print(f"[OK]   登录态已保存到: {AUTH_PATH}", flush=True)
    print("[DONE] 后续可复用，无需重新登录。", flush=True)

    print("[STAY] 浏览器再保持 60 秒方便确认。", flush=True)
    page.wait_for_timeout(60000)
    browser.close()
