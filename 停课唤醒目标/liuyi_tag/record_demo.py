#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""录屏 + 抓包：用户演示六一工作台「新建标签」+「新建用户群」全流程

复用六一登录态 (../liuyi_login/auth_state.json)
  → 进 portal
  → 点「六一工作台」瓷砖（新 tab）
  → 同时开启视频录制 + HAR 抓包（接口、请求体、响应体）
  → 用户手动演示：
      A. 新建标签
      B. 新建用户群
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
ROOT = SCRIPT_DIR.parent
AUTH_PATH = ROOT / "liuyi_login" / "auth_state.json"
VIDEO_DIR = SCRIPT_DIR / "videos"
HAR_PATH = SCRIPT_DIR / "demo_network.har"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到 {AUTH_PATH}，先跑 liuyi_login/login_liuyi.py", flush=True)
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
        page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)
    except Exception:
        print(f"[ERROR] 30s 内未渲染出「六一工作台」瓷砖，当前 URL: {page.url}", flush=True)
        page.screenshot(path=str(SCRIPT_DIR / "portal_expired.png"), full_page=True)
        ctx.close(); browser.close(); sys.exit(2)

    print("[STEP 2] 点击「六一工作台」瓷砖", flush=True)
    with ctx.expect_page(timeout=30000) as new_page_info:
        page.locator("p:has-text('六一工作台')").first.click()
    liuyi = new_page_info.value

    print("[STEP 3] 六一工作台新 tab 已打开，等待加载", flush=True)
    liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
    liuyi.wait_for_timeout(4000)
    print(f"  [LIUYI URL] {liuyi.url}", flush=True)

    landing_shot = SCRIPT_DIR / "liuyi_landing.png"
    liuyi.screenshot(path=str(landing_shot), full_page=True)
    print(f"[SHOT] 六一工作台首页截图: {landing_shot}", flush=True)

    print("=" * 60, flush=True)
    print("[录屏 + 抓包中] 请在浏览器里依次演示：", flush=True)
    print("  A. 新建标签", flush=True)
    print("     1) 左侧菜单进入「标签管理」", flush=True)
    print("     2) 点「新建」", flush=True)
    print("     3) 填命名规则（建议：停课唤醒_YYYYMMDD）", flush=True)
    print("     4) 上传/选择目标人群（用户ID 或 豌豆大账号 ID）", flush=True)
    print("     5) 提交保存", flush=True)
    print("  B. 新建用户群", flush=True)
    print("     1) 进入用户群入口（具体菜单等你演示）", flush=True)
    print("     2) 点「新建」", flush=True)
    print("     3) 命名 + 关联标签 / 上传 ID 列表", flush=True)
    print("     4) 提交保存", flush=True)
    print("两步都演示完后，直接关闭浏览器窗口（或等 20 分钟自动结束）。", flush=True)
    print("=" * 60, flush=True)

    try:
        liuyi.wait_for_event("close", timeout=1_200_000)  # 20min
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
