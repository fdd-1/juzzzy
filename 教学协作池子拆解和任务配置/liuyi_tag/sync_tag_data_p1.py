#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""P1服务池 - 标签数据同步（同步益智群到豌豆数合表）"""
import sys, io, time, argparse, datetime as dt, json
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
OUTPUT_DIR = ROOT / "output" / "p1"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

def log(m): print(m, flush=True)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", help="月份 YYYY-MM")
    args = ap.parse_args()

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到 {AUTH_PATH}")
        sys.exit(1)

    # 读取益智群名称
    today_tag = dt.date.today().strftime("%Y%m%d")
    group_path = OUTPUT_DIR / f"p1_group_ids_{today_tag}.json"
    if not group_path.exists():
        log(f"[ERROR] 找不到 {group_path}")
        sys.exit(1)
    group_data = json.loads(group_path.read_text(encoding="utf-8"))
    user_group_name = group_data["yizhi_group"]["name"]

    log(f"[INFO] 用户群名称（益智群）: {user_group_name}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=100)
        ctx = browser.new_context(storage_state=str(AUTH_PATH))
        page = ctx.new_page()

        log("[STEP 1] 打开 portal")
        page.goto(PORTAL_URL, timeout=30000)
        page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)

        log("[STEP 2] 点击「六一工作台」")
        with ctx.expect_page(timeout=30000) as new_page_info:
            page.locator("p:has-text('六一工作台')").first.click()
        liuyi = new_page_info.value
        liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
        liuyi.wait_for_timeout(5000)

        log("[STEP 3] 访问用户标签页面刷新")
        liuyi.goto("https://home.61info.cn/#/userTag", timeout=30000)
        liuyi.wait_for_load_state("networkidle", timeout=30000)
        liuyi.wait_for_timeout(2000)

        log("[STEP 4] 访问用户群页面刷新")
        liuyi.goto("https://home.61info.cn/#/userGroup", timeout=30000)
        liuyi.wait_for_load_state("networkidle", timeout=30000)
        liuyi.wait_for_timeout(2000)

        log("[STEP 5] 回到标签数据同步页面")
        liuyi.goto("https://home.61info.cn/#/tagDataSync", timeout=30000)
        liuyi.wait_for_load_state("networkidle", timeout=30000)
        liuyi.wait_for_timeout(5000)

        log("[STEP 6] 点击「新增」按钮")
        # 多重备选 selector，应对不同 UI 文案
        add_btn_selectors = [
            "button:has-text('新增')",
            "button:has-text('新 增')",
            "button:has-text('添加')",
            ".el-button:has-text('新增')",
            "button.el-button--primary:has-text('新增')",
        ]
        add_btn = None
        for sel in add_btn_selectors:
            cand = liuyi.locator(sel).first
            try:
                cand.wait_for(state="visible", timeout=3000)
                add_btn = cand
                log(f"  -> 命中 selector: {sel}")
                break
            except Exception:
                continue

        if add_btn is None:
            # 调试：保存当前页面截图 + HTML
            dbg_shot = OUTPUT_DIR / f"p1_sync_debug_{today_tag}.png"
            dbg_html = OUTPUT_DIR / f"p1_sync_debug_{today_tag}.html"
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            liuyi.screenshot(path=str(dbg_shot), full_page=True)
            dbg_html.write_text(liuyi.content(), encoding="utf-8")
            log(f"[ERROR] 找不到「新增」按钮，调试输出：{dbg_shot} / {dbg_html}")
            log(f"  当前 URL: {liuyi.url}")
            # 尝试列出所有 button 文本
            try:
                btn_texts = liuyi.locator("button").all_inner_texts()
                log(f"  页面上所有 button 文本: {btn_texts[:30]}")
            except Exception as e:
                log(f"  列举 button 失败: {e}")
            sys.exit(2)

        add_btn.click()
        liuyi.wait_for_timeout(2000)

        log("[STEP 7] 填写表单")
        # 业务类型：豌豆
        liuyi.locator("label:has-text('业务类型') + div .el-select").click()
        liuyi.wait_for_timeout(500)
        liuyi.locator(".el-select-dropdown__item:has-text('豌豆')").click()
        liuyi.wait_for_timeout(500)

        # 同步业务系统：豌豆数合表
        liuyi.locator("label:has-text('同步业务系统') + div .el-select").click()
        liuyi.wait_for_timeout(500)
        liuyi.locator(".el-select-dropdown__item:has-text('豌豆数合表')").click()
        liuyi.wait_for_timeout(500)

        # 同步用户群
        liuyi.locator("label:has-text('同步用户群') + div .el-select").click()
        liuyi.wait_for_timeout(500)
        liuyi.locator(f".el-select-dropdown__item:has-text('{user_group_name}')").click()
        liuyi.wait_for_timeout(500)

        # 同步数据频率：每天
        liuyi.locator("label:has-text('同步数据频率') + div .el-radio-group .el-radio:has-text('每天')").click()
        liuyi.wait_for_timeout(500)

        # 状态：启用
        liuyi.locator("label:has-text('状态') + div .el-radio-group .el-radio:has-text('启用')").click()
        liuyi.wait_for_timeout(500)

        log("[STEP 8] 点击「确认」")
        liuyi.locator("button:has-text('确认')").click()
        liuyi.wait_for_timeout(3000)

        log("[STEP 9] 点击「手动同步」")
        liuyi.locator("button:has-text('手动同步')").first.click()
        liuyi.wait_for_timeout(3000)

        # 截图
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        screenshot_path = OUTPUT_DIR / f"p1_sync_tag_final_{today_tag}.png"
        liuyi.screenshot(path=str(screenshot_path), full_page=True)
        log(f"[OK] 已截图保存到 {screenshot_path}")

        browser.close()

if __name__ == "__main__":
    main()
