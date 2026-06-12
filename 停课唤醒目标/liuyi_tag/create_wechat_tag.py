#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""调企微标签创建接口：把六一的「大账号用户群」关联到企微「【益智】长期标签」标签组下。

工作流：
  读 output/group_ids_{today}.json -> 拿 dadou_group.{groupId, name}
  POST /corporate-wechat-backend/o/v1/tagGroup/create
"""
import sys, io, json, time, argparse, datetime as dt
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
OUTPUT_DIR = ROOT / "output"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

CREATE_URL = "https://gw-mg.61info.cn/corporate-wechat-backend/o/v1/tagGroup/create"
LIST_URL = "https://gw-mg.61info.cn/corporate-wechat-backend/o/v1/tagGroup/getRelationList"

# 企微标签组「【益智】长期标签」的固定 id（来自 5/28 演示请求）
CORP_TAG_GROUP_ID = "etN7IECgAAkW39vv9E__scZlJAnXZFzw"
BIZ_CODE = "WANDOU"


def log(m): print(m, flush=True)


JS_FETCH = """
async ({url, method, body}) => {
  let token = '';
  for (let i = 0; i < localStorage.length; i++) {
    const v = localStorage.getItem(localStorage.key(i));
    if (v && typeof v === 'string' && v.startsWith('eyJ')) { token = v; break; }
  }
  const init = { method, headers: {'authorization': token}, credentials: 'include' };
  if (body !== null && body !== undefined) {
    init.headers['content-type'] = 'application/json';
    init.body = JSON.stringify(body);
  }
  const r = await fetch(url, init);
  let resp;
  try { resp = await r.json(); } catch (e) { resp = await r.text(); }
  return { status: r.status, body: resp, tokenFound: !!token };
}
"""


def call(pt, url, method="GET", body=None):
    return pt.evaluate(JS_FETCH, {"url": url, "method": method, "body": body})


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--group-ids-json", help="group_ids json，默认 output/group_ids_{today}.json")
    ap.add_argument("--corp-tag-group-id", default=CORP_TAG_GROUP_ID,
                    help="企微标签组 id（默认「【益智】长期标签」）")
    ap.add_argument("--dry-run", action="store_true", help="只构造请求体不调 create")
    args = ap.parse_args()

    today_tag = dt.date.today().strftime("%Y%m%d")
    gj_path = Path(args.group_ids_json) if args.group_ids_json else OUTPUT_DIR / f"group_ids_{today_tag}.json"
    if not gj_path.exists():
        log(f"[ERROR] 找不到 {gj_path}（先跑 liuyi-group）")
        sys.exit(1)
    gj = json.loads(gj_path.read_text(encoding="utf-8"))
    dadou = gj.get("dadou_group") or {}
    group_id = dadou.get("groupId")
    group_name = dadou.get("name")
    if not group_id or not group_name:
        log(f"[ERROR] {gj_path} 缺 dadou_group.groupId/name")
        sys.exit(2)

    log(f"[INFO] 关联企微标签：")
    log(f"  - 用户群: id={group_id}, name={group_name!r}")
    log(f"  - 企微标签组 id: {args.corp_tag_group_id}")

    payload = {
        "bizCode": BIZ_CODE,
        "corpTagGroupId": args.corp_tag_group_id,
        "userTagGroupList": [{"id": group_id, "name": group_name}],
    }
    log(f"  - payload: {json.dumps(payload, ensure_ascii=False)}")

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到 {AUTH_PATH}")
        sys.exit(1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=80)
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
        liuyi.wait_for_timeout(3000)

        if args.dry_run:
            log("[dry-run] 不调 create")
            browser.close()
            return

        log("[STEP 3] 调 tagGroup/create")
        r = call(liuyi, CREATE_URL, "POST", payload)
        log(f"  -> tokenFound={r.get('tokenFound')} status={r.get('status')} body={r.get('body')}")
        if r.get("status") != 200 or r["body"].get("code") != 0:
            log(f"[ERROR] create 失败")
            browser.close()
            sys.exit(3)

        # 验证：在 getRelationList 里找新创建的（按 createTime 倒序）
        log("[STEP 4] 校验：getRelationList 找新建条目")
        time.sleep(2)
        lr = call(liuyi, LIST_URL, "POST", {
            "bizCode": BIZ_CODE,
            "corpGroupName": "",
            "corpTagName": group_name,
            "orderByField": "",
            "pageNum": 1, "pageSize": 5, "total": 0,
        })
        items = lr.get("body", {}).get("data", {}).get("list", []) or []
        match = next((x for x in items if x.get("corpTagName") == group_name and x.get("groupId") == group_id), None)
        if match:
            log(f"  -> id={match['id']} corpGroupName={match['corpGroupName']!r} status={match.get('statusName')}")
            wechat_tag_id = match["id"]
        else:
            log(f"  [WARN] 未在列表里找到新建条目，列表前几条: {items[:2]}")
            wechat_tag_id = None

        out = OUTPUT_DIR / f"wechat_tag_{today_tag}.json"
        out.write_text(json.dumps({
            "wechat_tag_id": wechat_tag_id,
            "corp_tag_group_id": args.corp_tag_group_id,
            "corp_tag_name": group_name,
            "user_group_id": group_id,
            "payload": payload,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        log("=" * 60)
        log(f"[OK] 企微标签已创建 → {out}")
        log(f"  wechat_tag_id = {wechat_tag_id}")

        ctx.storage_state(path=str(AUTH_PATH))
        liuyi.wait_for_timeout(2000)
        browser.close()


if __name__ == "__main__":
    main()
