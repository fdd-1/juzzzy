#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""录屏 + 抓包：用户演示「平台中心 → 用户管理 → 批量用户查询」的全流程

复用六一登录态 (../liuyi_login/auth_state.json)
  → 进 portal
  → 点「平台中心」瓷砖（新 tab）
  → 同时开启视频录制 + HAR 抓包（接口、请求体、响应体）
  → 用户手动演示：新建批量查询 → 上传学员 ID → 选「豌豆大账号」导出类型 → 等待处理成功 → 导出
  → 关闭浏览器后，落盘 demo_video.webm + demo_network.har 供分析
"""
import sys, io, time, shutil
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
AUTH_PATH  = SCRIPT_DIR.parent / "liuyi_login" / "auth_state.json"
VIDEO_DIR  = SCRIPT_DIR / "videos"
HAR_PATH   = SCRIPT_DIR / "demo_network.har"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到 {AUTH_PATH}，请先跑 liuyi_login/login_liuyi.py 登录。", flush=True)
    sys.exit(1)

VIDEO_DIR.mkdir(exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome")
    ctx = browser.new_context(
        storage_state=str(AUTH_PATH),
        record_video_dir=str(VIDEO_DIR),
        record_video_size={"width": 1920, "height": 1080},
        record_har_path=str(HAR_PATH),
        record_har_content="embed",
    )
    page = ctx.new_page()

    print(f"[STEP 1] 打开 portal: {PORTAL_URL}", flush=True)
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")

    print("[STEP 1.5] 等 portal 瓷砖渲染（最多 30s）", flush=True)
    try:
        page.locator("p:has-text('平台中心')").first.wait_for(timeout=30000)
    except Exception:
        print(f"[ERROR] 30s 内未渲染出「平台中心」瓷砖，当前 URL: {page.url}", flush=True)
        page.screenshot(path=str(SCRIPT_DIR / "portal_expired.png"), full_page=True)
        ctx.close(); browser.close(); sys.exit(2)
    print(f"  [URL] {page.url}", flush=True)

    print("[STEP 2] 点击「平台中心」瓷砖", flush=True)
    with ctx.expect_page(timeout=30000) as new_page_info:
        page.locator("p:has-text('平台中心')").first.click()
    pt_page = new_page_info.value

    print("[STEP 3] 平台中心新 tab 已打开，等待加载", flush=True)
    pt_page.wait_for_load_state("domcontentloaded", timeout=30000)
    pt_page.wait_for_timeout(4000)
    print(f"  [PT URL] {pt_page.url}", flush=True)

    landing_shot = SCRIPT_DIR / "pingtai_landing.png"
    pt_page.screenshot(path=str(landing_shot), full_page=True)
    print(f"[SHOT] 平台中心首页截图: {landing_shot}", flush=True)

    print("=" * 60, flush=True)
    print("[录屏 + 抓包中] 请在浏览器里演示「批量用户查询」完整流程：", flush=True)
    print("  1. 左侧「用户管理 → 批量用户查询」", flush=True)
    print("  2. 右上角「新建」", flush=True)
    print("  3. 选择「导入类型 = 平台用户/学员ID」、「导出类型 = 豌豆大账号」", flush=True)
    print("  4. 上传一份小样的学员 ID 列表（5~10 条即可）", flush=True)
    print("  5. 提交 → 等列表里出现「处理成功」→ 点「导出」下载结果", flush=True)
    print("演示完成后，直接关闭浏览器窗口（或等 15 分钟自动结束）。", flush=True)
    print("=" * 60, flush=True)

    try:
        pt_page.wait_for_event("close", timeout=900_000)
    except Exception as e:
        print(f"[INFO] 等到超时或异常: {e}", flush=True)

    print("[STEP 4] 收尾：关闭 ctx，落盘 HAR 与视频", flush=True)
    try:
        ctx.close()
    except Exception as e:
        print(f"[WARN] ctx.close 异常: {e}", flush=True)
    try:
        browser.close()
    except Exception:
        pass

    time.sleep(3)
    videos = list(VIDEO_DIR.glob("*.webm"))
    if videos:
        latest = max(videos, key=lambda p: p.stat().st_mtime)
        final = SCRIPT_DIR / "demo_video.webm"
        try:
            shutil.copy2(str(latest), str(final))
            print(f"[OK] 录屏已保存: {final} ({final.stat().st_size/1024/1024:.1f} MB)", flush=True)
        except Exception as e:
            print(f"[WARN] 复制录屏失败: {e}，原文件: {latest}", flush=True)
    else:
        print("[WARN] 未找到录屏文件", flush=True)

    if HAR_PATH.exists():
        print(f"[OK] HAR 抓包已保存: {HAR_PATH} ({HAR_PATH.stat().st_size/1024/1024:.1f} MB)", flush=True)
    else:
        print("[WARN] 没找到 HAR 文件", flush=True)
