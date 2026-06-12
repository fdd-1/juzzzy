#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""P1服务池协作 - 创建用户群（参考P0的正确实现）
   1. 「【海外】海外学管2026年X月协作池-豌豆-X+1月服务池」      ← 复制模板16811，挂普通标签
   2. 「（益智）海外学管2026年X月协作池-豌豆-X+1月服务池」      ← 复制模板16812，挂益智标签
   命名规则严格执行：X = 任务执行当月，X+1 = 执行当月+1月
"""
import sys, io, json, time, random, argparse, datetime as dt
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

GW = "https://gw-mg.61info.cn/bizcenter-usertag/o/v1"
TEMPLATE_NORMAL = 16811   # 豌豆大账号模板（普通群）
TEMPLATE_YIZHI  = 16812   # 学员ID小账号模板（益智群）


def log(m): print(m, flush=True)


JS_FETCH = """
async ({url, method, body}) => {
  let token = '';
  for (let i = 0; i < localStorage.length; i++) {
    const v = localStorage.getItem(localStorage.key(i));
    if (v && typeof v === 'string' && v.startsWith('eyJ')) { token = v; break; }
  }
  const init = {
    method,
    headers: {'authorization': token},
    credentials: 'include'
  };
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


def gen_code(year, month, is_yizhi):
    suffix = "（yz）" if is_yizhi else ""
    rand = random.randint(10**8, 10**9)
    return f"{year}n{int(month)}yfwch{suffix}_{rand}"


def build_payload(template, new_name, new_main_tag_id, new_code):
    """从模板用户群构造新建请求体：替换 name/code/tagsConfig[0].tagIds，删除 id"""
    tags_config = []
    for i, item in enumerate(template["tagsConfig"]):
        new_item = dict(item)
        if i == 0:
            new_item["tagIds"] = [new_main_tag_id]
        new_item.setdefault("status", "")
        new_item.setdefault("showOperate", "")
        tags_config.append(new_item)
    payload = {
        "name": new_name,
        "code": new_code,
        "tagsConfig": tags_config,
        "bizChannelCode": template.get("bizChannelCode", "WANDOU"),
        "bizSysCode": template.get("bizSysCode", "all"),
        "groupType": template.get("groupType", 1),
        "status": template.get("status", 1),
        "operateUser": template.get("operateUser", ""),
        "createTime": template.get("createTime", ""),
        "updateTime": template.get("updateTime", ""),
        "cached": template.get("cached", 0),
    }
    return payload


def find_group_id(pt, name):
    """搜索用户群获取groupId"""
    list_r = call(pt, f"{GW}/userGroup/list", "POST",
                  {"bizChannelCode": "WANDOU", "name": name, "pageSize": 5, "pageNum": 1})
    items = list_r.get("body", {}).get("data", {}).get("list", []) or []
    matches = [x for x in items if x.get("name") == name]
    if not matches:
        return None
    return max(matches, key=lambda x: x.get("createTime", ""))["id"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--tag-ids-json", help="p1_tag_ids_*.json，默认 output/p1_tag_ids_{today}.json")
    ap.add_argument("--month", required=True, help="用户群命名月份，YYYY-MM")
    ap.add_argument("--suffix", default="", help="用户群名追加后缀")
    ap.add_argument("--dry-run", action="store_true", help="只构造请求体，不调 add")
    args = ap.parse_args()

    today_tag = dt.date.today().strftime("%Y%m%d")
    tag_ids_path = Path(args.tag_ids_json) if args.tag_ids_json else OUTPUT_DIR / f"p1_tag_ids_{today_tag}.json"
    if not tag_ids_path.exists():
        log(f"[ERROR] 找不到 {tag_ids_path}（先跑 create_tag_p1.py）")
        sys.exit(1)
    tag_ids = json.loads(tag_ids_path.read_text(encoding="utf-8"))
    normal_tag_id = tag_ids["normal_tag"]["tagId"]
    yizhi_tag_id  = tag_ids["yizhi_tag"]["tagId"]

    y, m = args.month.split("-")
    iy, im = int(y), int(m)
    # 传入的是内部月份 X+1，命名前缀用用户月份 X = X+1 - 1
    if im == 1:
        uy, um = iy - 1, 12
    else:
        uy, um = iy, im - 1
    log(f"[INFO] 用户月份 X={uy}-{um}，内部月份 X+1={iy}-{im}")

    # P1 命名规则（严格执行）：
    #   普通: 【海外】海外学管{uy}年{um}月协作池-豌豆-{im}月服务池
    #   益智: （益智）海外学管{uy}年{um}月协作池-豌豆-{im}月服务池
    name_normal = f"【海外】海外学管{uy}年{um}月协作池-豌豆-{im}月服务池{args.suffix}"
    name_yizhi  = f"（益智）海外学管{uy}年{um}月协作池-豌豆-{im}月服务池{args.suffix}"

    log(f"[INFO] 用户群名:")
    log(f"  - 普通: {name_normal}  (tagId={normal_tag_id}) <- 模板 {TEMPLATE_NORMAL}")
    log(f"  - 益智: {name_yizhi} (tagId={yizhi_tag_id}) <- 模板 {TEMPLATE_YIZHI}")

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到 {AUTH_PATH}")
        sys.exit(1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=80)
        ctx = browser.new_context(storage_state=str(AUTH_PATH))
        page = ctx.new_page()

        log(f"[STEP 1] 打开 portal")
        page.goto(PORTAL_URL, timeout=30000)
        page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)
        log("[STEP 2] 点击「六一工作台」")
        with ctx.expect_page(timeout=30000) as new_page_info:
            page.locator("p:has-text('六一工作台')").first.click()
        liuyi = new_page_info.value
        liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
        liuyi.wait_for_timeout(3000)

        results = {}
        for label, name, tagid, tmpl_id, is_yizhi in (
            ("normal_group", name_normal, normal_tag_id, TEMPLATE_NORMAL, False),
            ("yizhi_group",  name_yizhi,  yizhi_tag_id,  TEMPLATE_YIZHI,  True),
        ):
            log("=" * 60)
            log(f"[STEP {label}] 拉模板 {tmpl_id}")
            r = call(liuyi, f"{GW}/userGroup/query/{tmpl_id}", "GET")
            if r.get("status") != 200 or r["body"].get("code") != 0:
                log(f"[ERROR] 拉模板失败: {r}")
                browser.close(); sys.exit(2)
            template = r["body"]["data"]
            log(f"  -> 模板 name={template.get('name')!r} tagsConfig.len={len(template.get('tagsConfig', []))}")

            new_code = gen_code(y, m, is_yizhi)
            payload = build_payload(template, name, tagid, new_code)
            log(f"  -> 新 code = {new_code}")
            log(f"  -> tagsConfig: {json.dumps(payload['tagsConfig'], ensure_ascii=False)}")

            # checkData 预检
            log(f"[checkData]")
            chk_payload = dict(payload)
            chk_payload["id"] = tmpl_id
            chk = call(liuyi, f"{GW}/userGroup/checkData", "POST", chk_payload)
            log(f"  -> {chk.get('body')}")

            if args.dry_run:
                log("[dry-run] 跳过 add")
                results[label] = {"name": name, "code": new_code, "payload": payload}
                continue

            log(f"[add]")
            ar = call(liuyi, f"{GW}/userGroup/add", "POST", payload)
            log(f"  -> {ar.get('body')}")

            ar_body = ar.get("body") if isinstance(ar.get("body"), dict) else {}
            ar_code = ar_body.get("code")

            if ar_code == 0:
                # 创建成功，等异步落库后获取groupId
                time.sleep(2)
                group_id = find_group_id(liuyi, name)
                if not group_id:
                    log(f"[ERROR] 创建后未找到用户群 name={name}")
                    browser.close(); sys.exit(3)
                log(f"  -> 创建成功，groupId = {group_id}")
            elif ar_code == 101 and "已存在" in ar_body.get("msg", ""):
                # 用户群已存在，搜索获取groupId
                log(f"  [WARN] 用户群名已存在，搜索获取已有groupId")
                group_id = find_group_id(liuyi, name)
                if not group_id:
                    log(f"[ERROR] 用户群已存在但搜索不到 name={name}")
                    browser.close(); sys.exit(3)
                log(f"  -> 复用已存在的 groupId = {group_id}")
            else:
                log(f"[ERROR] add 失败: {ar_body}")
                browser.close(); sys.exit(3)

            results[label] = {"name": name, "code": new_code, "groupId": group_id}

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        groups_path = OUTPUT_DIR / f"p1_group_ids_{today_tag}.json"
        groups_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        log("=" * 60)
        log(f"[OK] 用户群已创建 → {groups_path}")
        log(json.dumps(results, ensure_ascii=False, indent=2))

        ctx.storage_state(path=str(AUTH_PATH))
        liuyi.wait_for_timeout(2000)
        browser.close()


if __name__ == "__main__":
    main()
