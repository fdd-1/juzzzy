#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""录屏：用户演示 OA 申请流程。
   - 复用 auth_state.json 登录
   - 打开 OA 系统
   - 开启视频录制
   - 用户手动操作，演示完整 OA 申请流程
   - 录屏保存到 oa_demo_video.webm
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
VIDEO_DIR  = SCRIPT_DIR / "videos"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到 {AUTH_PATH}，请先完成登录。", flush=True)
    sys.exit(1)

VIDEO_DIR.mkdir(exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome")
    ctx = browser.new_context(
        storage_state=str(AUTH_PATH),
        record_video_dir=str(VIDEO_DIR),
        record_video_size={"width": 1920, "height": 1080}
    )
    page = ctx.new_page()

    print(f"[STEP 1] 打开 portal: {PORTAL_URL}", flush=True)
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(3000)
    print(f"  [URL] {page.url}", flush=True)

    # 检查是否到了系统选择页
    has_oa = page.locator("p:has-text('OA系统')").count()
    if has_oa > 0:
        print("[STEP 2] 点击 OA系统 瓷砖", flush=True)
        with ctx.expect_page(timeout=30000) as new_page_info:
            page.locator("p:has-text('OA系统')").first.click()
        oa_page = new_page_info.value
        print("[STEP 3] OA 新 tab 已打开", flush=True)
        oa_page.wait_for_load_state("domcontentloaded", timeout=30000)
        oa_page.wait_for_timeout(3000)
        print(f"  [OA URL] {oa_page.url}", flush=True)
        active_page = oa_page
    else:
        print("[INFO] 未检测到系统选择页，可能已在 OA 内", flush=True)
        active_page = page

    print("=" * 60, flush=True)
    print("[录屏中] 请在浏览器里演示完整的 OA 申请流程：", flush=True)
    print("        1. 从 OA 首页开始", flush=True)
    print("        2. 点击对应的申请入口", flush=True)
    print("        3. 填写表单、上传附件", flush=True)
    print("        4. 提交申请", flush=True)
    print("        5. 演示完成后，关闭浏览器窗口（或等 10 分钟自动结束）", flush=True)
    print("=" * 60, flush=True)

    # 等待用户操作，最多 10 分钟
    try:
        active_page.wait_for_timeout(600000)
    except Exception as e:
        print(f"[INFO] 页面关闭或超时: {e}", flush=True)

    print("[STEP 4] 关闭浏览器，保存录屏", flush=True)
    ctx.close()
    browser.close()

    import time, shutil
    time.sleep(3)

    # 找到录屏文件
    videos = list(VIDEO_DIR.glob("*.webm"))
    if videos:
        latest = max(videos, key=lambda p: p.stat().st_size)
        final = SCRIPT_DIR / "oa_demo_video.webm"
        try:
            shutil.copy2(str(latest), str(final))
        except Exception as e:
            print(f"[WARN] 复制失败: {e}，文件仍在: {latest}", flush=True)
            final = latest
        print(f"[OK] 录屏已保存: {final}", flush=True)
        print(f"     大小: {final.stat().st_size / 1024 / 1024:.1f} MB", flush=True)
    else:
        print("[WARN] 未找到录屏文件", flush=True)
