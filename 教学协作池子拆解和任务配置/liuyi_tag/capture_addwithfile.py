#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""打开六一工作台，监听 userTag/addwithfile 请求与响应，
   方便手动创建一次标签时拓包对比 create_tag.py 的参数差异。

用法：
    python liuyi_tag/capture_addwithfile.py
   会保持浏览器打开 10 分钟（可 Ctrl+C 提前退出），手动完成一次创建标签。
   抓到的请求会写到 liuyi_tag/capture_addwithfile.json
"""
import sys, io, json, time
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

ROOT = Path(__file__).parent.parent
AUTH = ROOT / "liuyi_login" / "auth_state.json"
OUT  = Path(__file__).parent / "capture_addwithfile.json"
PORTAL = "https://dingding.61info.cn/sys/portal/page.jsp"

captured = []

def on_request(req):
    if "userTag/add" not in req.url:
        return
    item = {
        "type": "request",
        "url": req.url,
        "method": req.method,
        "headers": dict(req.headers),
    }
    try:
        item["post_data"] = req.post_data
    except Exception:
        item["post_data"] = "<unreadable>"
    try:
        # multipart 解析（粗略：只取边界后字段名 / Content-Disposition）
        body_raw = req.post_data_buffer
        if body_raw:
            try:
                text = body_raw.decode("utf-8", errors="replace")
            except Exception:
                text = "<binary>"
            # 截取前 8KB 看字段
            item["post_data_text_8k"] = text[:8000]
    except Exception:
        pass
    captured.append(item)
    print(f"[REQ] {req.method} {req.url}")
    print(f"  headers.authorization={req.headers.get('authorization','')[:40]}...")
    print(f"  headers.content-type={req.headers.get('content-type','')}")

def on_response(resp):
    if "userTag/add" not in resp.url:
        return
    item = {
        "type": "response",
        "url": resp.url,
        "status": resp.status,
    }
    try:
        item["body"] = resp.text()
    except Exception:
        item["body"] = "<unreadable>"
    captured.append(item)
    print(f"[RESP {resp.status}] {resp.url}")
    print(f"  body: {item['body'][:1000]}")


def main():
    if not AUTH.exists():
        print(f"[ERROR] 找不到 {AUTH}，先跑 liuyi_login/login_liuyi.py")
        sys.exit(1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=80)
        ctx = browser.new_context(storage_state=str(AUTH))
        page = ctx.new_page()
        page.on("request", on_request)
        page.on("response", on_response)

        page.goto(PORTAL, timeout=30000)
        page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)
        with ctx.expect_page(timeout=30000) as info:
            page.locator("p:has-text('六一工作台')").first.click()
        liuyi = info.value
        liuyi.on("request", on_request)
        liuyi.on("response", on_response)
        liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
        liuyi.wait_for_timeout(3000)

        print("=" * 70)
        print("[ACTION] 浏览器已打开六一工作台。")
        print("[ACTION] 请手动操作：标签管理 -> 新建标签 -> 上传 xlsx -> 提交")
        print("[ACTION] 创建一个测试标签即可（建议直接用 P0 文件创建）")
        print("[ACTION] 创建完成后回到此终端，按 Ctrl+C 退出，会自动保存抓包")
        print("=" * 70)

        try:
            # 等待最多 30 分钟
            for _ in range(180):
                liuyi.wait_for_timeout(10000)
        except KeyboardInterrupt:
            pass

        OUT.write_text(json.dumps(captured, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"[OK] 抓包写入 {OUT}（共 {len(captured)} 条事件）")
        ctx.storage_state(path=str(AUTH))
        browser.close()


if __name__ == "__main__":
    main()
