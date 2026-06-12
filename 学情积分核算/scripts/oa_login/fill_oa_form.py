#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""OA 豌豆币添加申请 - 表单填写脚本
   前提：用户已手动点击「豌豆币添加申请」进入表单页面
   功能：自动填写表单、上传附件、提交
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
READY_FILE = SCRIPT_DIR / "form_ready.txt"  # 信号文件

# 数据来源（一次性调试脚本：把期次目录写在这里）
PROJECT_DIR = Path(__file__).resolve().parent.parent.parent
WORK_DIR = PROJECT_DIR / "03_output" / "20260501-20260515学情积分发放明细_v2"
ATTACHMENT = WORK_DIR / "发放豌豆币文档填写模板.xlsx"
TOTAL_AMOUNT = 1257000  # 从积分汇总计算得出

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到 {AUTH_PATH}", flush=True)
    sys.exit(1)

if not ATTACHMENT.exists():
    print(f"[ERROR] 找不到附件: {ATTACHMENT}", flush=True)
    sys.exit(1)

# 清理旧信号
if READY_FILE.exists():
    READY_FILE.unlink()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=200)
    ctx = browser.new_context(storage_state=str(AUTH_PATH))
    page = ctx.new_page()

    # 打开 OA portal
    page.goto("https://dingding.61info.cn/sys/portal/page.jsp", timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(3000)

    print("=" * 60, flush=True)
    print("[等待] 请在浏览器里：", flush=True)
    print("       1. 点击「豌豆币添加申请」", flush=True)
    print("       2. 进入表单页面后，创建文件: form_ready.txt", flush=True)
    print(f"          (或者我来监测 URL 变化)", flush=True)
    print("=" * 60, flush=True)

    # 轮询等待：URL 变化 或 信号文件
    import time
    start_url = page.url
    deadline = time.time() + 300
    while time.time() < deadline:
        if READY_FILE.exists() or page.url != start_url:
            print(f"[检测] 页面已变化或收到信号", flush=True)
            break
        time.sleep(2)
    else:
        print("[TIMEOUT] 5 分钟超时", flush=True)
        browser.close()
        sys.exit(1)

    print(f"[当前 URL] {page.url}", flush=True)
    page.screenshot(path=str(SCRIPT_DIR / "form_before_fill.png"), full_page=True)

    print(f"[STEP 1] 填写「需要添加的豌豆币/魔力币数总共」= {TOTAL_AMOUNT}", flush=True)
    # 根据截图，这个字段的值是 1164500，输入框可能在"需要添加"这个 label 附近
    amount_input = page.locator("input").filter(has=page.locator("text=/需要添加.*豌豆币.*魔力币.*总共/"))
    if amount_input.count() == 0:
        # 尝试通过附近文本定位
        amount_input = page.locator("text=/需要添加.*总共/").locator("..").locator("input").first
    if amount_input.count() == 0:
        # 最后尝试：找值为 1164500 的输入框（如果已经填过）
        amount_input = page.locator("input[value='1164500']").first

    if amount_input.count() > 0:
        amount_input.clear()
        amount_input.fill(str(TOTAL_AMOUNT))
        print(f"  [已填写] {TOTAL_AMOUNT}", flush=True)
    else:
        print("[WARN] 找不到金额输入框，请手动填写", flush=True)

    page.wait_for_timeout(1000)

    print(f"[STEP 2] 上传附件: {ATTACHMENT.name}", flush=True)
    file_input = page.locator("input[type=file]").first
    if file_input.count() > 0:
        file_input.set_input_files(str(ATTACHMENT))
        print(f"  [已上传] {ATTACHMENT.name}", flush=True)
        page.wait_for_timeout(3000)
    else:
        print("[WARN] 找不到文件上传控件，请手动上传", flush=True)

    page.screenshot(path=str(SCRIPT_DIR / "form_after_fill.png"), full_page=True)

    print("[STEP 3] 查找「流程处理」或「提交」按钮", flush=True)
    # 根据截图底部有「流程处理」按钮
    submit_btn = page.locator("button:has-text('流程处理'), button:has-text('提交')").first
    if submit_btn.count() == 0:
        submit_btn = page.locator("input[type=submit], input[value*='提交']").first

    if submit_btn.count() > 0:
        print("  [找到提交按钮，准备点击]", flush=True)
        print("  [等待 5 秒，请确认表单填写无误...]", flush=True)
        page.wait_for_timeout(5000)
        submit_btn.click()
        print("  [已点击提交]", flush=True)
        page.wait_for_timeout(5000)
        page.screenshot(path=str(SCRIPT_DIR / "form_submitted.png"), full_page=True)
        print(f"  [提交后 URL] {page.url}", flush=True)
    else:
        print("[WARN] 找不到提交按钮，请手动提交", flush=True)

    print("[OK] 脚本执行完成，浏览器保持 30 秒", flush=True)
    page.wait_for_timeout(30000)
    browser.close()
