#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""北极星外呼平台：克隆既有「【海外】停课120天以内-N月」任务模板，把月份替换为目标月份提交。

虽然用户口径是「修改」，但接口实际是 POST /task/taskTemplate/add（克隆 + 新建），
原任务保留不删。

工作流：
  1. 复用 ../polaris_login/auth_state.json 进 sh-center.vipthink.cn
  2. POST /task/taskTemplate/16/list 搜任务（关键词「停课120天以内」），按 createTime 取最新一条
  3. GET  /task/taskTemplate/getDetail?id={id} 拉详情
  4. 把 taskName 末尾的「-N月」替换为目标月份
  5. 把详情字段加工成 add 接口期望的形状（去掉顶层 id 等只读字段）
  6. POST /task/taskTemplate/add 提交
  7. 再次列表确认新建条目存在
"""
import sys, io, json, time, re, argparse, datetime as dt
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
ROOT = SCRIPT_DIR.parent
AUTH_PATH = ROOT / "polaris_login" / "auth_state.json"
OUTPUT_DIR = ROOT / "output"
HOME_URL = "https://sh-center.vipthink.cn/#/"

GW = "https://eos-gateway.vipthink.cn/task/taskTemplate"
LIST_URL = f"{GW}/16/list"
DETAIL_URL = f"{GW}/getDetail"
ADD_URL = f"{GW}/add"

DEFAULT_KEYWORD = "停课120天以内"
DEFAULT_BUSINESS_ID = "16"


def log(m): print(m, flush=True)


JS_FETCH = """
async ({url, method, body}) => {
  // 找 token：优先 localStorage 里 Bearer 开头或 eyJ 开头的串
  let token = '';
  function findToken(store) {
    for (let i = 0; i < store.length; i++) {
      const k = store.key(i);
      const v = store.getItem(k);
      if (!v || typeof v !== 'string') continue;
      if (v.startsWith('Bearer ')) return v;
      if (v.startsWith('eyJ')) return 'Bearer ' + v;
      // 有些前端把 token 存在 JSON 里
      try {
        const j = JSON.parse(v);
        const candidates = [j.token, j.accessToken, j.access_token, j.Authorization, j.authorization];
        for (const c of candidates) {
          if (typeof c === 'string') {
            if (c.startsWith('Bearer ')) return c;
            if (c.startsWith('eyJ')) return 'Bearer ' + c;
          }
        }
      } catch (e) {}
    }
    return '';
  }
  token = findToken(localStorage) || findToken(sessionStorage);

  const init = { method, headers: {} };
  if (token) init.headers['authorization'] = token;
  if (body !== null && body !== undefined) {
    init.headers['content-type'] = 'application/json;charset=UTF-8';
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


# 详情字段 → add 字段的映射规则
# add 请求体不包含的顶层字段（来自详情但 add 不需要）
DETAIL_DROP = {
    "id", "creatorName", "creatorId", "createTime", "updateTime",
    "subjectName", "departmentList", "departmentNames",
    "cancelOverdueDay",  # 演示请求里没出现
    "autoRuleParam",     # 演示里也去掉了
    "isUnionOrder",      # add 里有保留 false
    "latestServiceTime",
}
# 这些保留在 add payload 中
DETAIL_KEEP_ALSO = ("isUnionOrder",)


def detail_to_add_payload(detail, new_task_name):
    """把 getDetail 返回的 data 加工成 add 接口的请求体。
       核心动作：复制大部分字段，改 taskName，丢掉只读字段，子任务里补一些 UI 字段。
       不保证完美，但尽量贴合演示的 add 请求体。
    """
    p = dict(detail)
    p["taskName"] = new_task_name
    # 统一把 businessId 转字符串（演示请求里是 "16"）
    if "businessId" in p:
        p["businessId"] = str(p["businessId"])
    # 丢字段
    for k in DETAIL_DROP:
        p.pop(k, None)
    # 演示请求里这两个一定存在
    p.setdefault("isUnionOrder", False)
    # 子任务模板加工
    children = p.get("childTaskTemplateList") or []
    new_children = []
    for c in children:
        cc = dict(c)
        complete_list = cc.get("completeList") or []
        new_completes = []
        for item in complete_list:
            ii = dict(item)
            # param 在详情里可能是 "1s" 或一个 JSON 字符串；演示里 add 时拆成 number/对象 顶层字段
            param = ii.get("param", "")
            if isinstance(param, str) and param.startswith("{"):
                try:
                    nested = json.loads(param)
                    for kk, vv in nested.items():
                        ii.setdefault(kk, vv)
                    ii["param"] = nested.get("times", 1)
                except Exception:
                    pass
            elif isinstance(param, str) and param.endswith("s"):
                # "1s" → number 1，unit="s"
                try:
                    ii["param"] = int(param[:-1])
                    ii.setdefault("unit", "s")
                except Exception:
                    pass
            # completeStrategyObj：演示里加了，但只是冗余信息，最小化补
            ii.setdefault("completeStrategyObj", {
                "operatorType": ii.get("operatorType"),
                "completeStrategyName": ii.get("completeStrategyName"),
                "completeStrategy": ii.get("completeStrategy"),
            })
            ii.setdefault("strategyList", [])
            ii.setdefault("times", ii.get("times", 1))
            ii.setdefault("interval", ii.get("interval", 1))
            ii.setdefault("isRemind", ii.get("isRemind", False))
            ii.setdefault("resultIds", ii.get("resultIds", []))
            new_completes.append(ii)
        cc["completeList"] = new_completes
        cc.setdefault("showBody", False)
        new_children.append(cc)
    p["childTaskTemplateList"] = new_children
    # 演示里附加的字段
    p.setdefault("count", detail.get("count", 0))
    p.setdefault("timeCycle", 1)
    p.setdefault("timeCycleChild", ["0"])
    p.setdefault("monthCycleTimeChild", 0)
    return p


def replace_month(name, target_month):
    """把任务名末尾的「-N月」替换为「-{target_month}月新（模拟）」"""
    new, n = re.subn(r"-\d+月\b", f"-{target_month}月新（模拟）", name)
    if n == 0:
        # 退路：没匹配到就追加
        return f"{name}-{target_month}月新（模拟）"
    return new


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keyword", default=DEFAULT_KEYWORD,
                    help=f"任务名关键词（默认 {DEFAULT_KEYWORD}）")
    ap.add_argument("--business-id", default=DEFAULT_BUSINESS_ID, help="业务线 id（默认 16）")
    ap.add_argument("--target-month", type=int,
                    help="目标月份（数字 1-12）。不传则按「下月」（运营节奏是月底为下月准备）")
    ap.add_argument("--source-task-id", type=int,
                    help="基线任务 id；不传则按 keyword 搜最新一条")
    ap.add_argument("--dry-run", action="store_true", help="只构造 payload 不调 add")
    args = ap.parse_args()

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到 {AUTH_PATH}（先跑 polaris_login/login_polaris.py）")
        sys.exit(1)

    if args.target_month is None:
        # 默认下月（演示行为：5/28 把 5月版本克隆到 6月版本）
        cur = dt.date.today()
        target_month = cur.month + 1 if cur.month < 12 else 1
    else:
        target_month = args.target_month
    if target_month < 1 or target_month > 12:
        log(f"[ERROR] target_month 非法: {target_month}")
        sys.exit(1)
    log(f"[INFO] 目标月份 = {target_month}, 关键词 = {args.keyword!r}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=80)
        ctx = browser.new_context(storage_state=str(AUTH_PATH))
        page = ctx.new_page()

        log(f"[STEP 1] 打开 home: {HOME_URL}")
        page.goto(HOME_URL, timeout=30000)
        page.wait_for_load_state("domcontentloaded")
        page.wait_for_timeout(4000)

        # === Step 2: 找基线任务 ===
        if args.source_task_id:
            source_id = args.source_task_id
            log(f"[STEP 2] 用 --source-task-id={source_id}")
        else:
            log(f"[STEP 2] 搜任务列表 keyword={args.keyword!r}")
            r = call(page, LIST_URL, "POST", {
                "isEnable": 1,
                "taskName": args.keyword,
                "businessId": args.business_id,
                "pageNum": 1, "pageSize": 20,
            })
            log(f"  -> tokenFound={r.get('tokenFound')} status={r.get('status')}")
            if r.get("status") != 200 or r["body"].get("code") != 0:
                log(f"[ERROR] list 失败: {r}")
                browser.close(); sys.exit(2)
            items = r["body"].get("data", {}).get("list", []) or []
            log(f"  -> 命中 {len(items)} 条")
            for it in items[:5]:
                log(f"     id={it['id']} name={it['taskName']!r} createTime={it.get('createTime')}")
            if not items:
                log(f"[ERROR] keyword {args.keyword!r} 没匹配任何任务")
                browser.close(); sys.exit(3)
            # 取 createTime 最新的
            best = max(items, key=lambda x: x.get("createTime", ""))
            source_id = best["id"]
            log(f"  -> 选最新: id={source_id} name={best['taskName']!r}")

        # === Step 3: 拉详情 ===
        log(f"[STEP 3] GET getDetail?id={source_id}")
        r = call(page, f"{DETAIL_URL}?id={source_id}", "GET")
        if r.get("status") != 200 or r["body"].get("code") != 0:
            log(f"[ERROR] getDetail 失败: {r}")
            browser.close(); sys.exit(4)
        detail = r["body"]["data"]
        original_name = detail.get("taskName", "")
        log(f"  -> 原任务名: {original_name!r}")

        # === Step 4: 改名 ===
        new_name = replace_month(original_name, target_month)
        log(f"  -> 新任务名: {new_name!r}")
        if new_name == original_name:
            log("[WARN] 新名 == 原名，不会真的克隆，请检查 --target-month 或者 keyword")

        # === Step 5: 构造 add payload ===
        payload = detail_to_add_payload(detail, new_name)
        log(f"[STEP 5] payload 摘要: businessId={payload.get('businessId')!r}, "
            f"taskType={payload.get('taskType')}, childCount={len(payload.get('childTaskTemplateList') or [])}")

        if args.dry_run:
            log("[dry-run] 跳过 add")
            log(json.dumps(payload, ensure_ascii=False, indent=2)[:2000])
            browser.close()
            return

        # === Step 6: 提交 ===
        log(f"[STEP 6] POST add")
        ar = call(page, ADD_URL, "POST", payload)
        log(f"  -> {ar.get('body')}")
        if ar.get("status") != 200 or ar["body"].get("code") != 0:
            log(f"[ERROR] add 失败")
            browser.close(); sys.exit(5)

        # === Step 7: 校验 ===
        log("[STEP 7] 列表校验新建任务")
        time.sleep(2)
        lr = call(page, LIST_URL, "POST", {
            "isEnable": 1,
            "taskName": new_name,
            "businessId": args.business_id,
            "pageNum": 1, "pageSize": 5,
        })
        items = lr.get("body", {}).get("data", {}).get("list", []) or []
        match = next((x for x in items if x.get("taskName") == new_name), None)
        new_id = match["id"] if match else None
        if match:
            log(f"  -> id={new_id} name={match['taskName']!r} createTime={match.get('createTime')}")
        else:
            log(f"  [WARN] 未在列表里找到新建条目，前几条: {[i.get('taskName') for i in items[:3]]}")

        # === 落盘 ===
        OUTPUT_DIR.mkdir(exist_ok=True)
        today_tag = dt.date.today().strftime("%Y%m%d")
        out = OUTPUT_DIR / f"polaris_task_{today_tag}.json"
        out.write_text(json.dumps({
            "source_task_id": source_id,
            "source_task_name": original_name,
            "new_task_id": new_id,
            "new_task_name": new_name,
            "target_month": target_month,
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        log("=" * 60)
        log(f"[OK] 北极星任务已克隆 → {out}")
        log(f"  source: id={source_id} name={original_name!r}")
        log(f"  new:    id={new_id} name={new_name!r}")

        ctx.storage_state(path=str(AUTH_PATH))
        page.wait_for_timeout(2000)
        browser.close()


if __name__ == "__main__":
    main()
