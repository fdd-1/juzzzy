#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""录屏 + 抓包：用户演示「修改既有停课唤醒外呼任务」全流程

复用 ../polaris_login/auth_state.json
  → 进 sh-center.vipthink.cn
  → 同时开启视频录制 + HAR 抓包（接口、请求体、响应体）
  → 用户手动演示：
      A. 进入外呼任务列表
      B. 找到既有「停课唤醒」任务 → 编辑
      C. 修改名称（按当前月份）+ 必要的人群关联
      D. 保存提交（务必，否则 HAR 抓不到 update 接口）
  → 关闭浏览器后，落盘 demo_video.webm + demo_network.har
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
ROOT = SCRIPT_DIR.parent
AUTH_PATH = ROOT / "polaris_login" / "auth_state.json"
VIDEO_DIR = SCRIPT_DIR / "videos"
HAR_PATH = SCRIPT_DIR / "demo_network.har"
HOME_URL = "https://sh-center.vipthink.cn/#/"

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到 {AUTH_PATH}，先跑 polaris_login/login_polaris.py", flush=True)
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

    print(f"[STEP 1] 打开北极星 home: {HOME_URL}", flush=True)
    page.goto(HOME_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(4000)
    print(f"  [URL] {page.url}", flush=True)

    landing_shot = SCRIPT_DIR / "polaris_landing.png"
    page.screenshot(path=str(landing_shot), full_page=True)
    print(f"[SHOT] 北极星首页截图: {landing_shot}", flush=True)

    print("=" * 60, flush=True)
    print("[录屏 + 抓包中] 请在浏览器里演示「修改既有停课唤醒外呼任务」：", flush=True)
    print("  1. 左侧菜单进入外呼任务列表", flush=True)
    print("  2. 找到既有的「停课唤醒」任务 → 点编辑", flush=True)
    print("  3. 修改任务名称（改成本月版本）", flush=True)
    print("  4. 如有人群/标签关联也演示出来", flush=True)
    print("  5. 务必点保存提交（不点保存 HAR 抓不到接口！）", flush=True)
    print("演示完后直接关闭浏览器窗口（或等 20 分钟自动结束）。", flush=True)
    print("=" * 60, flush=True)

    try:
        page.wait_for_event("close", timeout=1_200_000)  # 20min
    except Exception as e:
        print(f"[INFO] 等到超时或异常: {e}", flush=True)

    print("[STEP 2] 收尾：关闭 ctx，落盘 HAR 与视频", flush=True)
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
