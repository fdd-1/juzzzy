#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys, io
from pathlib import Path
from playwright.sync_api import sync_playwright

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stdout.reconfigure(line_buffering=True)

AUTH = Path(r"C:\Users\fengjianyi\Desktop\六一标签\auth_state.json")

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=200)
    ctx = browser.new_context(storage_state=str(AUTH))
    liuyi = ctx.new_page()

    liuyi.goto("https://home.61info.cn/#/tagDataSync", timeout=30000)
    liuyi.wait_for_timeout(5000)
    liuyi.screenshot(path="logs/debug_1_before_reload.png", full_page=True)

    liuyi.reload(timeout=30000)
    liuyi.wait_for_timeout(5000)
    liuyi.screenshot(path="logs/debug_2_after_reload.png", full_page=True)

    btns = liuyi.evaluate("() => Array.from(document.querySelectorAll('button')).map(b => b.textContent.trim())")
    print("按钮:", btns)
    print("URL:", liuyi.url)

    browser.close()
