#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""北极星外呼平台：克隆既有「P0（学位预警）」任务模板，把月份替换为目标月份，挂普通用户群，设教师任务完成时间。

工作流：
  1. 复用 ../polaris_login/auth_state.json 进 sh-center.vipthink.cn
  2. POST /task/taskTemplate/16/list 搜任务（关键词「P0（学位预警）」），按 createTime 取最新一条
  3. GET  /task/taskTemplate/getDetail?id={id} 拉详情
  4. 把 taskName 末尾的「-N月」替换为目标月份
  5. 读取 group_ids_*.json 拿普通用户群的 groupId，替换详情里的 userGroupIds
  6. 替换 teacherTaskCompleteTime 为用户传入的时间
  7. 把详情字段加工成 add 接口期望的形状（去掉顶层 id 等只读字段）
  8. POST /task/taskTemplate/add 提交
  9. 再次列表确认新建条目存在
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
OUTPUT_DIR = ROOT / "output" / "p0"
HOME_URL = "https://sh-center.vipthink.cn/#/"

GW = "https://eos-gateway.vipthink.cn/task/taskTemplate"
LIST_URL = f"{GW}/16/list"
DETAIL_URL = f"{GW}/getDetail"
ADD_URL = f"{GW}/add"

DEFAULT_KEYWORD = "P0（学位预警）"
DEFAULT_BUSINESS_ID = "16"


def log(m): print(m, flush=True)


JS_FETCH = """
async ({url, method, body}) => {
  let token = '';
  function findToken(store) {
    for (let i = 0; i < store.length; i++) {
      const k = store.key(i);
      const v = store.getItem(k);
      if (!v || typeof v !== 'string') continue;
      if (v.startsWith('Bearer ')) return v;
      if (v.startsWith('eyJ')) return 'Bearer ' + v;
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
DETAIL_DROP = {
    "id", "creatorName", "creatorId", "createTime", "updateTime",
    "subjectName", "departmentList", "departmentNames",
    "autoRuleParam",
}
# 这些字段必须从原 detail 透传过去，不能丢也不能默认 False
DETAIL_KEEP_FROM_SOURCE = ("isUnionOrder", "cancelOverdueDay", "isEnable", "isAutoCancel", "isInheritable", "isDurable", "isFanout")


def detail_to_add_payload(detail, new_task_name, new_group_id, teacher_complete_time):
    """把 getDetail 返回的 data 加工成 add 接口的请求体。
       核心动作：复制大部分字段，改 taskName、userGroupIds、latestServiceTime（教师任务完成时间），丢掉只读字段。
    """
    p = dict(detail)
    p["taskName"] = new_task_name
    # 挂普通用户群
    p["userGroupIds"] = [new_group_id]
    # 教师任务完成时间 → latestServiceTime（不是 teacherTaskCompleteTime！）
    p["latestServiceTime"] = teacher_complete_time
    # 统一把 businessId 转字符串
    if "businessId" in p:
        p["businessId"] = str(p["businessId"])
    # 透传必须保留的字段
    for k in DETAIL_KEEP_FROM_SOURCE:
        if k in detail and detail[k] is not None:
            p[k] = detail[k]
    # 丢字段
    for k in DETAIL_DROP:
        p.pop(k, None)
    # 子任务模板加工（照搬停课唤醒的逻辑）
    children = p.get("childTaskTemplateList") or []
    new_children = []
    for c in children:
        cc = dict(c)
        complete_list = cc.get("completeList") or []
        new_completes = []
        for item in complete_list:
            ii = dict(item)
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
                try:
                    ii["param"] = int(param[:-1])
                    ii.setdefault("unit", "s")
                except Exception:
                    pass
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
    p.setdefault("count", detail.get("count", 0))
    p.setdefault("timeCycle", 1)
    p.setdefault("timeCycleChild", ["0"])
    p.setdefault("monthCycleTimeChild", 0)
    return p


def replace_month(name, target_month):
    """把任务名末尾的「-N月」替换为「-{target_month}月」"""
    new, n = re.subn(r"-\d+月\b", f"-{target_month}月", name)
    if n == 0:
        return f"{name}-{target_month}月"
    return new


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--keyword", default=DEFAULT_KEYWORD, help=f"任务名关键词（默认 {DEFAULT_KEYWORD}）")
    ap.add_argument("--business-id", default=DEFAULT_BUSINESS_ID, help="业务线 id（默认 16）")
    ap.add_argument("--target-month", type=int, required=True, help="目标月份（数字 1-12）")
    ap.add_argument("--teacher-complete-time", required=True,
                    help="教师任务完成时间（格式 YYYY-MM-DD HH:mm:ss）")
    ap.add_argument("--group-ids-json", help="group_ids_*.json，默认 output/group_ids_{today}.json")
    ap.add_argument("--source-task-id", type=int, help="基线任务 id；不传则按 keyword 搜最新一条")
    ap.add_argument("--suffix", default="", help="新任务名追加后缀，例如 --suffix 测试 用于试跑")
    ap.add_argument("--dry-run", action="store_true", help="只构造 payload 不调 add")
    args = ap.parse_args()

    target_month = args.target_month
    if target_month < 1 or target_month > 12:
        log(f"[ERROR] target_month 非法: {target_month}")
        sys.exit(1)
    log(f"[INFO] 目标月份 = {target_month}, 关键词 = {args.keyword!r}")
    log(f"[INFO] 教师任务完成时间 = {args.teacher_complete_time!r}")

    # 读取 group_ids
    today_tag = dt.date.today().strftime("%Y%m%d")
    group_ids_path = Path(args.group_ids_json) if args.group_ids_json else OUTPUT_DIR / f"group_ids_{today_tag}.json"
    if not group_ids_path.exists():
        log(f"[ERROR] 找不到 {group_ids_path}（先跑 liuyi_tag/create_group.py）")
        sys.exit(1)
    group_ids = json.loads(group_ids_path.read_text(encoding="utf-8"))
    normal_group_id = group_ids["normal_group"]["groupId"]
    log(f"  -> 普通用户群 groupId = {normal_group_id}")

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到 {AUTH_PATH}（先跑 polaris_login/login_polaris.py）")
        sys.exit(1)

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
        new_name = replace_month(original_name, target_month) + (args.suffix or "")
        log(f"  -> 新任务名: {new_name!r}")

        # === Step 5: 构造 add payload ===
        payload = detail_to_add_payload(detail, new_name, normal_group_id, args.teacher_complete_time)
        log(f"[STEP 5] payload 摘要: businessId={payload.get('businessId')!r}, "
            f"taskType={payload.get('taskType')}, userGroupIds={payload.get('userGroupIds')}, "
            f"teacherTaskCompleteTime={payload.get('teacherTaskCompleteTime')!r}")

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
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out = OUTPUT_DIR / f"polaris_task_{today_tag}.json"
        out.write_text(json.dumps({
            "source_task_id": source_id,
            "source_task_name": original_name,
            "new_task_id": new_id,
            "new_task_name": new_name,
            "target_month": target_month,
            "teacher_complete_time": args.teacher_complete_time,
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
