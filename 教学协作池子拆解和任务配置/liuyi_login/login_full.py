#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""一键登录 + 进入六一工作台，保证 auth_state.json 含全部子域 cookie

流程：
  1) 打开登录页，等用户扫码
  2) 登录成功后跳 portal（写入 dingding.61info.cn cookie）
  3) 点击「六一工作台」瓷砖（写入 home.61info.cn / gw-mg.61info.cn cookie）
  4) 等子域加载完成后保存 storage_state
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
LOGIN_URL  = "https://login.61info.cn/"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome")
        ctx = browser.new_context()
        page = ctx.new_page()

        print(f"[STEP 1] 打开登录页: {LOGIN_URL}", flush=True)
        page.goto(LOGIN_URL, timeout=30000)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(2000)

        print("=" * 60, flush=True)
        print("[ACTION] 请用钉钉扫码登录（页面右侧二维码）", flush=True)
        print("=" * 60, flush=True)

        # 等扫码成功
        deadline = time.time() + 600
        logged_in = False
        last_url = page.url
        while time.time() < deadline:
            cur_url = page.url
            if cur_url != last_url:
                print(f"  [POLL] {cur_url}", flush=True)
                last_url = cur_url
            try:
                if page.locator("p:has-text('六一工作台')").count() > 0:
                    logged_in = True
                    break
                if "login.61info.cn" not in cur_url.lower():
                    logged_in = True
                    break
            except Exception:
                pass
            page.wait_for_timeout(2000)

        if not logged_in:
            print("[ERROR] 10 分钟内未检测到登录成功", flush=True)
            browser.close()
            sys.exit(1)

        print("[STEP 2] 登录成功，强制访问 portal 写入 dingding.61info.cn cookie", flush=True)
        page.goto(PORTAL_URL, timeout=30000)
        page.wait_for_load_state("domcontentloaded")
        try:
            page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)
        except Exception:
            print(f"[ERROR] 30s 未渲染出六一瓷砖，当前 URL: {page.url}", flush=True)
            browser.close()
            sys.exit(2)
        print(f"  [URL] {page.url}", flush=True)

        print("[STEP 3] 点击「六一工作台」瓷砖进入子域", flush=True)
        with ctx.expect_page(timeout=30000) as info:
            page.locator("p:has-text('六一工作台')").first.click()
        liuyi = info.value
        liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
        # 等子域 JS / cookie 完全初始化
        liuyi.wait_for_timeout(8000)
        print(f"  [LIUYI URL] {liuyi.url}", flush=True)

        # 触发一次后端 API 调用确保 token 加到 localStorage（list 接口最轻量）
        try:
            res = liuyi.evaluate("""
            async () => {
              let token = '';
              for (let i = 0; i < localStorage.length; i++) {
                const v = localStorage.getItem(localStorage.key(i));
                if (v && typeof v === 'string' && v.startsWith('eyJ')) { token = v; break; }
              }
              if (!token) return { ok: false, reason: 'no token in localStorage' };
              const r = await fetch('https://gw-mg.61info.cn/bizcenter-usertag/o/v1/userTag/list', {
                method: 'POST',
                headers: {'content-type':'application/json','authorization':token},
                body: JSON.stringify({bizChannelCode: 'WANDOU', name: '_warmup_', pageSize: 1, pageNum: 1}),
                credentials: 'include'
              });
              const j = await r.json();
              return { ok: r.status === 200, status: r.status, code: j.code, msg: j.msg };
            }
            """)
            print(f"[STEP 4] 接口预热: {res}", flush=True)
        except Exception as e:
            print(f"[WARN] 接口预热失败（不影响 cookie 保存）: {e}", flush=True)

        ctx.storage_state(path=str(AUTH_PATH))
        print(f"[OK] auth_state.json 已保存（含全部子域 cookie）→ {AUTH_PATH}", flush=True)

        liuyi.wait_for_timeout(2000)
        browser.close()


if __name__ == "__main__":
    main()
