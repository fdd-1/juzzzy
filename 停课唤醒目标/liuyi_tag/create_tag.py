#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""调六一工作台 userTag/addwithfile 接口创建两个标签：
   1. 「2026年{月}月海外益智停课学员」          ← user_ids_*.xlsx
   2. 「2026年{月}月海外益智停课学员（大账号）」 ← dadou_ids_*.xlsx

复用 ../liuyi_login/auth_state.json
浏览器里 page.evaluate 调 fetch（multipart 通过 base64 在 JS 端还原 File）
"""
import sys, io, json, base64, argparse, datetime as dt
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

# 接口
ADD_TAG_URL = "https://gw-mg.61info.cn/bizcenter-usertag/o/v1/userTag/addwithfile"
LIST_TAG_URL = "https://gw-mg.61info.cn/bizcenter-usertag/o/v1/userTag/list"

# 业务常量（来自前端 JS 反编译 + HAR 抓包）
BIZ_CODE = "WANDOU"
TAG_TYPE = 531  # 关键行为/活跃行为
TAG_TYPES_STRING = "[53,531]"
DATA_FROM = 2  # USERID
TAG_UPDATE_TYPE = 1
ONCE_UPDATE = 1
STATUS = 1


def log(m): print(m, flush=True)


def call_eval(pt, js, arg):
    return pt.evaluate(js, arg)


JS_ADD_TAG = """
async ({name, fileB64, fileName, addUrl, bizCode, tagType, tagTypesString, dataFrom, tagUpdateType, onceUpdate, status}) => {
  let token = '';
  for (let i = 0; i < localStorage.length; i++) {
    const v = localStorage.getItem(localStorage.key(i));
    if (v && typeof v === 'string' && v.startsWith('eyJ')) { token = v; break; }
  }
  const bin = atob(fileB64);
  const buf = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
  const blob = new Blob([buf], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const file = new File([blob], fileName, {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const fd = new FormData();
  fd.append('bizCode', bizCode);
  fd.append('tagType', String(tagType));
  fd.append('tagTypesString', tagTypesString);
  fd.append('name', name);
  fd.append('dataFrom', String(dataFrom));
  fd.append('tagUpdateType', String(tagUpdateType));
  fd.append('sensorsGroupId', '');
  fd.append('importFileName', fileName);
  fd.append('status', String(status));
  fd.append('onceUpdate', String(onceUpdate));
  fd.append('file', file);
  const r = await fetch(addUrl, {
    method: 'POST',
    headers: {'authorization': token},
    body: fd,
    credentials: 'include'
  });
  let body;
  try { body = await r.json(); } catch (e) { body = await r.text(); }
  return { status: r.status, body, tokenFound: !!token };
}
"""

JS_LIST_TAG = """
async ({name, listUrl, bizCode}) => {
  let token = '';
  for (let i = 0; i < localStorage.length; i++) {
    const v = localStorage.getItem(localStorage.key(i));
    if (v && typeof v === 'string' && v.startsWith('eyJ')) { token = v; break; }
  }
  const r = await fetch(listUrl, {
    method: 'POST',
    headers: {'content-type':'application/json','authorization':token},
    body: JSON.stringify({bizChannelCode: bizCode, name: name, pageSize: 10, pageNum: 1}),
    credentials: 'include'
  });
  return await r.json();
}
"""


def add_tag(pt, name, file_path):
    file_bytes = Path(file_path).read_bytes()
    file_b64 = base64.b64encode(file_bytes).decode("ascii")
    file_name = Path(file_path).name
    arg = {
        "name": name,
        "fileB64": file_b64,
        "fileName": file_name,
        "addUrl": ADD_TAG_URL,
        "bizCode": BIZ_CODE,
        "tagType": TAG_TYPE,
        "tagTypesString": TAG_TYPES_STRING,
        "dataFrom": DATA_FROM,
        "tagUpdateType": TAG_UPDATE_TYPE,
        "onceUpdate": ONCE_UPDATE,
        "status": STATUS,
    }
    return call_eval(pt, JS_ADD_TAG, arg)


def find_tag_id(pt, name):
    res = call_eval(pt, JS_LIST_TAG, {"name": name, "listUrl": LIST_TAG_URL, "bizCode": BIZ_CODE})
    if res.get("code") != 0:
        return None, res
    items = res.get("data", {}).get("list", []) or []
    matches = [x for x in items if x.get("name") == name]
    if not matches:
        return None, items
    return max(matches, key=lambda x: x.get("createTime", ""))["id"], matches[0]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--user-ids-xlsx", help="学员 ID xlsx，默认 latest_inputs.json 里的")
    ap.add_argument("--dadou-ids-xlsx", help="大账号 ID xlsx，默认 latest_inputs.json 里的")
    ap.add_argument("--month", help="标签命名月份，YYYY-MM，默认按当前月")
    args = ap.parse_args()

    # 解析月份：默认按 Windows 当月（运营节奏：当月跑当月数据）
    # 也可以 --month YYYY-MM 显式指定（比如月底测下月，或者补跑历史月）
    if args.month:
        y, m = args.month.split("-")
        log(f"[INFO] 使用 --month 指定月份 {y}-{int(m)}")
    else:
        today = dt.date.today()
        y, m = str(today.year), str(today.month)
    name_user = f"{y}年{int(m)}月海外益智停课学员新（模拟）"
    name_dadou = f"{y}年{int(m)}月海外益智停课学员（大账号）新（模拟）"

    # 文件
    manifest = SCRIPT_DIR / "latest_inputs.json"
    inputs = {}
    if manifest.exists():
        inputs = json.loads(manifest.read_text(encoding="utf-8"))
    user_xlsx = Path(args.user_ids_xlsx) if args.user_ids_xlsx else Path(inputs.get("user_ids_xlsx", ""))
    dadou_xlsx = Path(args.dadou_ids_xlsx) if args.dadou_ids_xlsx else Path(inputs.get("dadou_ids_xlsx", ""))
    if not user_xlsx.exists():
        log(f"[ERROR] 找不到 user_ids xlsx: {user_xlsx}（先跑 prepare_csv.py）")
        sys.exit(1)
    if not dadou_xlsx.exists():
        log(f"[ERROR] 找不到 dadou_ids xlsx: {dadou_xlsx}")
        sys.exit(1)

    log(f"[INFO] 标签名:")
    log(f"  - 小账号: {name_user}  (file: {user_xlsx.name})")
    log(f"  - 大账号: {name_dadou} (file: {dadou_xlsx.name})")

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到 {AUTH_PATH}，先跑 liuyi_login/login_liuyi.py")
        sys.exit(1)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=80)
        ctx = browser.new_context(storage_state=str(AUTH_PATH))
        page = ctx.new_page()

        log(f"[STEP 1] 打开 portal: {PORTAL_URL}")
        page.goto(PORTAL_URL, timeout=30000)
        page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)
        log("[STEP 2] 点击「六一工作台」瓷砖")
        with ctx.expect_page(timeout=30000) as new_page_info:
            page.locator("p:has-text('六一工作台')").first.click()
        liuyi = new_page_info.value
        liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
        liuyi.wait_for_timeout(3000)
        log(f"  [LIUYI URL] {liuyi.url}")

        results = {}
        for label, name, fp in (
            ("user_tag",  name_user,  user_xlsx),
            ("dadou_tag", name_dadou, dadou_xlsx),
        ):
            log("=" * 60)
            log(f"[STEP 3-{label}] 创建标签「{name}」")
            r = add_tag(liuyi, name, fp)
            log(f"  -> tokenFound={r.get('tokenFound')} status={r.get('status')} body={r.get('body')}")
            if r.get("status") != 200 or (isinstance(r.get("body"), dict) and r["body"].get("code") != 0):
                log(f"[ERROR] 创建标签失败")
                ctx.storage_state(path=str(AUTH_PATH))
                browser.close()
                sys.exit(2)

            # 等异步落库
            liuyi.wait_for_timeout(2000)
            tag_id, info = find_tag_id(liuyi, name)
            if not tag_id:
                log(f"[ERROR] 创建后未在 list 里找到标签 name={name}, info={info}")
                ctx.storage_state(path=str(AUTH_PATH))
                browser.close()
                sys.exit(3)
            log(f"  -> 拿到 tagId = {tag_id}")
            results[label] = {"tagId": tag_id, "name": name}

        # 落盘
        OUTPUT_DIR.mkdir(exist_ok=True)
        today_tag = dt.date.today().strftime("%Y%m%d")
        tag_ids_path = OUTPUT_DIR / f"tag_ids_{today_tag}.json"
        tag_ids_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        log("=" * 60)
        log(f"[OK] 标签已创建并落盘 → {tag_ids_path}")
        log(json.dumps(results, ensure_ascii=False, indent=2))

        ctx.storage_state(path=str(AUTH_PATH))
        liuyi.wait_for_timeout(3000)
        browser.close()


if __name__ == "__main__":
    main()
