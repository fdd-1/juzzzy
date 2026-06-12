#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""P1服务池协作 - 创建标签（参考P0的正确实现）
   1. 「【海外】26年X月教学协作池-X月服务池」          ← p1_dadou_ids_*.xlsx（豌豆大账号ID）
   2. 「【海外】26年X月教学协作池-X月服务池（益智）」 ← p1_user_ids_*.xlsx（学员ID）
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
OUTPUT_DIR = ROOT / "output" / "p1"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

ADD_TAG_URL = "https://gw-mg.61info.cn/bizcenter-usertag/o/v1/userTag/addwithfile"
LIST_TAG_URL = "https://gw-mg.61info.cn/bizcenter-usertag/o/v1/userTag/list"

BIZ_CODE = "WANDOU"
TAG_TYPE = 531
TAG_TYPES_STRING = "[53,531]"
DATA_FROM = 2
TAG_UPDATE_TYPE = 1
ONCE_UPDATE = 1
STATUS = 1

def log(m): print(m, flush=True)

# 创建标签的JS（与P0完全一致）
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

# 搜索标签获取ID的JS（与P0完全一致）
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

def add_tag(page, name, file_path):
    """创建标签"""
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
    return page.evaluate(JS_ADD_TAG, arg)

def find_tag_id(page, name):
    """通过搜索API获取标签ID"""
    res = page.evaluate(JS_LIST_TAG, {"name": name, "listUrl": LIST_TAG_URL, "bizCode": BIZ_CODE})
    if res.get("code") != 0:
        return None, res
    items = res.get("data", {}).get("list", []) or []
    matches = [x for x in items if x.get("name") == name]
    if not matches:
        return None, items
    return max(matches, key=lambda x: x.get("createTime", ""))["id"], matches[0]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="目标月份 YYYY-MM（例：2026-06）")
    args = ap.parse_args()

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到登录态文件：{AUTH_PATH}")
        sys.exit(1)

    # 解析月份（传入的是内部月份 X+1，用户做的是 X 月任务）
    try:
        iy, im = map(int, args.month.split("-"))
    except:
        log(f"[ERROR] --month 格式错误：{args.month}")
        sys.exit(1)
    # 推导用户月份 X
    if im == 1:
        uy, um = iy - 1, 12
    else:
        uy, um = iy, im - 1
    month_cn = f"{uy-2000}年{um}月"  # 命名前缀用用户月份 X

    # 读取 latest_inputs_p1.json
    manifest_path = SCRIPT_DIR / "latest_inputs_p1.json"
    if not manifest_path.exists():
        log(f"[ERROR] 找不到 {manifest_path}，请先运行 filter_p1.py")
        sys.exit(2)
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    dadou_path = Path(manifest["dadou_ids_xlsx"])
    user_path = Path(manifest["user_ids_xlsx"])

    if not dadou_path.exists() or not user_path.exists():
        log(f"[ERROR] 筛选产物文件不存在")
        sys.exit(2)

    # 标签命名：X月教学协作-(X+1)月服务池
    name_normal = f"【海外】{month_cn}教学协作-{im}月服务池"
    name_yizhi = f"【海外】{month_cn}教学协作-{im}月服务池（益智）"

    log(f"[INFO] 标签名:")
    log(f"  - 普通: {name_normal}  (file: {dadou_path.name})")
    log(f"  - 益智: {name_yizhi} (file: {user_path.name})")

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
            ("normal_tag", name_normal, dadou_path),
            ("yizhi_tag",  name_yizhi,  user_path),
        ):
            log("=" * 60)
            log(f"[STEP 3-{label}] 创建标签「{name}」")
            r = add_tag(liuyi, name, fp)
            log(f"  -> tokenFound={r.get('tokenFound')} status={r.get('status')} body={r.get('body')}")

            body = r.get("body") if isinstance(r.get("body"), dict) else {}
            code = body.get("code")

            if code == 0:
                # 创建成功，等异步落库后获取tagId
                liuyi.wait_for_timeout(2000)
                tag_id, info = find_tag_id(liuyi, name)
                if not tag_id:
                    log(f"[ERROR] 创建后未在 list 里找到标签 name={name}, info={info}")
                    ctx.storage_state(path=str(AUTH_PATH))
                    browser.close()
                    sys.exit(3)
                log(f"  -> 创建成功，tagId = {tag_id}")
            elif code == 101 and "已存在" in body.get("msg", ""):
                # 标签已存在，直接搜索获取tagId
                log(f"  [WARN] 标签名已存在，搜索获取已有tagId")
                tag_id, info = find_tag_id(liuyi, name)
                if not tag_id:
                    log(f"[ERROR] 标签已存在但搜索不到，name={name}, info={info}")
                    ctx.storage_state(path=str(AUTH_PATH))
                    browser.close()
                    sys.exit(3)
                log(f"  -> 复用已存在的tagId = {tag_id}")
            else:
                log(f"[ERROR] 创建标签失败")
                ctx.storage_state(path=str(AUTH_PATH))
                browser.close()
                sys.exit(2)

            results[label] = {"tagId": tag_id, "name": name}

        # 落盘
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        today_tag = dt.date.today().strftime("%Y%m%d")
        tag_ids_path = OUTPUT_DIR / f"p1_tag_ids_{today_tag}.json"
        tag_ids_path.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        log("=" * 60)
        log(f"[OK] 标签已创建并落盘 → {tag_ids_path}")
        log(json.dumps(results, ensure_ascii=False, indent=2))

        ctx.storage_state(path=str(AUTH_PATH))
        liuyi.wait_for_timeout(3000)
        browser.close()

if __name__ == "__main__":
    main()
