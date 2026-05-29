"""停课唤醒目标 全流程自动化入口

借鉴 x-cli 的统一输出契约：
  成功：{"ok": true, "data": {...}}
  失败：{"ok": false, "error": {"code": "...", "message": "..."}}

子步骤：
  export        从 BI 拉「思维停课学员执行明细」
  filter        基于原始 xlsx 做业务筛选（待实现）
  liuyi-tag     六一工作台新建标签（待实现）
  liuyi-group   六一工作台新建用户群（待实现）
  wechat-tag    新建企微标签（待实现）
  polaris-task  北极星建立外呼任务（待实现）
  all           端到端
"""

import argparse
import calendar
import datetime as dt
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

BI_SKILL = Path.home() / ".workbuddy" / "skills" / "bi_skill" / "bi_skill.py"
REPORT_NAME = "思维停课学员执行明细"
REPORT_PATH = r"海外直播业务线\海外后端\思维-后端\服务\停课唤醒\思维停课学员执行明细"
LIUYI_GROUP_VALUE = "历史停课未唤醒"


def envelope_ok(data):
    return {"ok": True, "data": data}


def envelope_err(code, message, **extra):
    err = {"code": code, "message": message}
    err.update(extra)
    return {"ok": False, "error": err}


def emit(payload):
    print(json.dumps(payload, ensure_ascii=False))


def compute_signin_window(month_str: str | None) -> tuple[str, str, str]:
    """根据导出当月 M，算出签到时间窗口 [M-4 月 1 号, M-1 月最后一天]。

    用户口径：距今前 1~4 个月。5 月跑 → 1 月到 4 月；6 月跑 → 2 月到 5 月。
    入参 month_str 形如 '2026-05'；为空则取今天所在月。
    返回 (start_date, end_date, base_month_label)，日期均为 'YYYY-MM-DD'。
    """
    if month_str:
        y, m = map(int, month_str.split("-"))
    else:
        today = dt.date.today()
        y, m = today.year, today.month

    # 前 4 月 1 号
    start_y, start_m = y, m - 4
    while start_m <= 0:
        start_m += 12
        start_y -= 1
    start_date = dt.date(start_y, start_m, 1)

    # 前 1 月最后一天
    end_y, end_m = y, m - 1
    while end_m <= 0:
        end_m += 12
        end_y -= 1
    last_day = calendar.monthrange(end_y, end_m)[1]
    end_date = dt.date(end_y, end_m, last_day)

    return (
        start_date.strftime("%Y-%m-%d"),
        end_date.strftime("%Y-%m-%d"),
        f"{y}-{m:02d}",
    )


def cmd_export(args):
    if not BI_SKILL.exists():
        emit(envelope_err("bi_skill_missing", f"未找到 bi_skill: {BI_SKILL}"))
        return 2

    start, end, base = compute_signin_window(args.month)
    download_dir = OUTPUT_DIR / "downloads"
    download_dir.mkdir(exist_ok=True)

    # BI 只设停课学员分组，开始/结束日期保持默认不动
    cmd = [
        sys.executable,
        str(BI_SKILL),
        "search",
        "--path",
        REPORT_PATH,
        "--filters",
        f"停课学员分组={LIUYI_GROUP_VALUE}",
        "--output",
        str(download_dir),
    ]

    print(f"[export] 当月={base}  导出后二次筛选签到时间=[{start}, {end}]", file=sys.stderr)
    print(f"[export] 调用 bi_skill: {' '.join(cmd)}", file=sys.stderr)

    if args.dry_run:
        emit(envelope_ok({
            "dry_run": True,
            "month": base,
            "signin_start": start,
            "signin_end": end,
            "filters": {"停课学员分组": LIUYI_GROUP_VALUE},
            "command": cmd,
        }))
        return 0

    try:
        result = subprocess.run(cmd, check=False)
    except Exception as e:
        emit(envelope_err("bi_skill_exec_failed", str(e)))
        return 3

    if result.returncode != 0:
        emit(envelope_err(
            "bi_skill_returned_nonzero",
            f"bi_skill exit code={result.returncode}",
            command=cmd,
        ))
        return result.returncode

    # 找最新 xlsx，规范化拷贝到 raw_{yyyymmdd}.xlsx
    xlsxs = sorted(download_dir.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not xlsxs:
        emit(envelope_err("no_xlsx_downloaded", f"导出后未在 {download_dir} 找到 xlsx"))
        return 4

    latest = xlsxs[0]
    today_tag = dt.date.today().strftime("%Y%m%d")
    target = OUTPUT_DIR / f"raw_{today_tag}.xlsx"
    shutil.copy2(latest, target)

    emit(envelope_ok({
        "month": base,
        "signin_start": start,
        "signin_end": end,
        "raw_xlsx": str(target),
        "downloaded": str(latest),
    }))
    return 0


def cmd_not_implemented(name):
    def _run(_args):
        emit(envelope_err("not_implemented", f"步骤 '{name}' 尚未实现"))
        return 1
    return _run


HEADER_ROW = 9
SIGNIN_COL = "月初前最近一次签到时间"
WAKEUP_COL = "是否停课唤醒"


def cmd_filter(args):
    """二次筛选：月初前最近一次签到时间 ∈ [M-4月1号, M-2月最后一天] 且 是否停课唤醒=0"""
    import pandas as pd

    start, end, base = compute_signin_window(args.month)

    # 定位输入文件
    if args.input:
        raw_path = Path(args.input)
    else:
        today_tag = dt.date.today().strftime("%Y%m%d")
        raw_path = OUTPUT_DIR / f"raw_{today_tag}.xlsx"

    if not raw_path.exists():
        emit(envelope_err("input_not_found", f"找不到输入文件: {raw_path}"))
        return 2

    print(f"[filter] 读取 {raw_path}，header_row={HEADER_ROW}", file=sys.stderr)
    df = pd.read_excel(str(raw_path), header=HEADER_ROW)
    total = len(df)
    print(f"[filter] 原始行数: {total}", file=sys.stderr)

    # 筛选 是否停课唤醒 == 0
    df[WAKEUP_COL] = pd.to_numeric(df[WAKEUP_COL], errors="coerce")
    df = df[df[WAKEUP_COL] == 0]
    after_wakeup = len(df)
    print(f"[filter] 是否停课唤醒=0 后: {after_wakeup} 行", file=sys.stderr)

    # 筛选签到时间窗口
    df[SIGNIN_COL] = pd.to_datetime(df[SIGNIN_COL], errors="coerce")
    start_dt = pd.Timestamp(start)
    end_dt = pd.Timestamp(end)
    df = df[(df[SIGNIN_COL] >= start_dt) & (df[SIGNIN_COL] <= end_dt)]
    after_signin = len(df)
    print(f"[filter] 签到时间 [{start}, {end}] 后: {after_signin} 行", file=sys.stderr)

    # 输出
    today_tag = dt.date.today().strftime("%Y%m%d")
    out_path = OUTPUT_DIR / f"filtered_{today_tag}.xlsx"
    df.to_excel(str(out_path), index=False)

    emit(envelope_ok({
        "month": base,
        "signin_window": [start, end],
        "total_rows": total,
        "after_wakeup_filter": after_wakeup,
        "after_signin_filter": after_signin,
        "output": str(out_path),
    }))
    return 0


def cmd_query_dadou_id(args):
    """步骤 2.5：把 filtered_*.xlsx 里的学员 ID 喂给平台中心「批量用户查询」，
       拿回「学员ID → 豌豆大账号 ID」映射，落盘 output/dadou_mapping_{today}.xlsx。

       两步串行：
         a. pingtai_query/prepare_template.py  → upload_{stamp}.xlsx
         b. pingtai_query/query_dadou.py       → dadou_mapping_{today}.xlsx
    """
    pq_dir = ROOT / "pingtai_query"
    prepare = pq_dir / "prepare_template.py"
    query = pq_dir / "query_dadou.py"
    if not prepare.exists() or not query.exists():
        emit(envelope_err("scripts_missing", f"找不到 {prepare} 或 {query}"))
        return 2

    # a. 生成 upload xlsx
    prep_cmd = [sys.executable, str(prepare)]
    if args.input:
        prep_cmd += ["--input", args.input]
    print(f"[query-dadou-id] step a: {' '.join(prep_cmd)}", file=sys.stderr)
    r = subprocess.run(prep_cmd, check=False)
    if r.returncode != 0:
        emit(envelope_err("prepare_failed", f"prepare_template exit={r.returncode}"))
        return r.returncode

    # b. Playwright 自动化查询
    q_cmd = [sys.executable, str(query)]
    if args.import_type:
        q_cmd += ["--import-type", args.import_type]
    if args.export_type:
        q_cmd += ["--export-type", args.export_type]
    if args.poll_timeout:
        q_cmd += ["--poll-timeout", str(args.poll_timeout)]
    print(f"[query-dadou-id] step b: {' '.join(q_cmd)}", file=sys.stderr)
    r = subprocess.run(q_cmd, check=False)
    if r.returncode != 0:
        emit(envelope_err("query_failed", f"query_dadou exit={r.returncode}"))
        return r.returncode

    today_tag = dt.date.today().strftime("%Y%m%d")
    out = OUTPUT_DIR / f"dadou_mapping_{today_tag}.xlsx"
    if not out.exists():
        emit(envelope_err("output_missing", f"未找到 {out}"))
        return 3
    emit(envelope_ok({
        "mapping_xlsx": str(out),
        "size_kb": round(out.stat().st_size / 1024, 1),
    }))
    return 0


def cmd_liuyi_tag(args):
    """步骤 3：六一工作台 新建标签 ×2（小账号 + 大账号）

       两步串行：
         a. liuyi_tag/prepare_csv.py     → user_ids_*.xlsx + dadou_ids_*.xlsx
         b. liuyi_tag/create_tag.py      → output/tag_ids_{today}.json
    """
    lt_dir = ROOT / "liuyi_tag"
    prep = lt_dir / "prepare_csv.py"
    create = lt_dir / "create_tag.py"
    if not prep.exists() or not create.exists():
        emit(envelope_err("scripts_missing", f"找不到 {prep} 或 {create}"))
        return 2

    # a
    prep_cmd = [sys.executable, str(prep)]
    if args.filtered:
        prep_cmd += ["--filtered", args.filtered]
    if args.dadou:
        prep_cmd += ["--dadou", args.dadou]
    print(f"[liuyi-tag] step a: {' '.join(prep_cmd)}", file=sys.stderr)
    r = subprocess.run(prep_cmd, check=False)
    if r.returncode != 0:
        emit(envelope_err("prepare_failed", f"prepare_csv exit={r.returncode}"))
        return r.returncode

    # b
    c_cmd = [sys.executable, str(create)]
    if args.month:
        c_cmd += ["--month", args.month]
    print(f"[liuyi-tag] step b: {' '.join(c_cmd)}", file=sys.stderr)
    r = subprocess.run(c_cmd, check=False)
    if r.returncode != 0:
        emit(envelope_err("create_tag_failed", f"create_tag exit={r.returncode}"))
        return r.returncode

    today_tag = dt.date.today().strftime("%Y%m%d")
    out = OUTPUT_DIR / f"tag_ids_{today_tag}.json"
    if not out.exists():
        emit(envelope_err("output_missing", f"未找到 {out}"))
        return 3
    import json as _json
    emit(envelope_ok({"tag_ids": _json.loads(out.read_text(encoding="utf-8")), "file": str(out)}))
    return 0


def cmd_liuyi_group(args):
    """步骤 3.5：六一工作台 新建用户群 ×2（复制 16811/16812 模板，替换主 tagId）

       依赖：output/tag_ids_{today}.json（由 liuyi-tag 产出）
    """
    create = ROOT / "liuyi_tag" / "create_group.py"
    if not create.exists():
        emit(envelope_err("scripts_missing", f"找不到 {create}"))
        return 2

    cmd = [sys.executable, str(create)]
    if args.month:
        cmd += ["--month", args.month]
    if args.tag_ids_json:
        cmd += ["--tag-ids-json", args.tag_ids_json]
    if args.dry_run:
        cmd += ["--dry-run"]
    print(f"[liuyi-group] {' '.join(cmd)}", file=sys.stderr)
    r = subprocess.run(cmd, check=False)
    if r.returncode != 0:
        emit(envelope_err("create_group_failed", f"create_group exit={r.returncode}"))
        return r.returncode

    today_tag = dt.date.today().strftime("%Y%m%d")
    out = OUTPUT_DIR / f"group_ids_{today_tag}.json"
    if not out.exists():
        emit(envelope_err("output_missing", f"未找到 {out}"))
        return 3
    import json as _json
    emit(envelope_ok({"group_ids": _json.loads(out.read_text(encoding="utf-8")), "file": str(out)}))
    return 0


def cmd_wechat_tag(args):
    """步骤 4：六一工作台「【益智】长期标签」企微标签组下，关联本批次大账号用户群

       依赖：output/group_ids_{today}.json（由 liuyi-group 产出）
       注：只关联大账号用户群，小账号不关联
    """
    create = ROOT / "liuyi_tag" / "create_wechat_tag.py"
    if not create.exists():
        emit(envelope_err("scripts_missing", f"找不到 {create}"))
        return 2

    cmd = [sys.executable, str(create)]
    if args.group_ids_json:
        cmd += ["--group-ids-json", args.group_ids_json]
    if args.corp_tag_group_id:
        cmd += ["--corp-tag-group-id", args.corp_tag_group_id]
    if args.dry_run:
        cmd += ["--dry-run"]
    print(f"[wechat-tag] {' '.join(cmd)}", file=sys.stderr)
    r = subprocess.run(cmd, check=False)
    if r.returncode != 0:
        emit(envelope_err("create_wechat_tag_failed", f"create_wechat_tag exit={r.returncode}"))
        return r.returncode

    if args.dry_run:
        emit(envelope_ok({"dry_run": True}))
        return 0

    today_tag = dt.date.today().strftime("%Y%m%d")
    out = OUTPUT_DIR / f"wechat_tag_{today_tag}.json"
    if not out.exists():
        emit(envelope_err("output_missing", f"未找到 {out}"))
        return 3
    import json as _json
    emit(envelope_ok({"wechat_tag": _json.loads(out.read_text(encoding="utf-8")), "file": str(out)}))
    return 0


def cmd_polaris_task(args):
    """步骤 5：北极星外呼平台 — 克隆既有「【海外】停课120天以内-N月」任务到目标月份

       本质是 POST /task/taskTemplate/add（克隆 + 新建），原任务保留。
       默认目标月份 = 下月（运营节奏：月底为下月准备）。
    """
    update = ROOT / "polaris_task" / "update_task.py"
    if not update.exists():
        emit(envelope_err("scripts_missing", f"找不到 {update}"))
        return 2

    cmd = [sys.executable, str(update)]
    if args.keyword:
        cmd += ["--keyword", args.keyword]
    if args.target_month is not None:
        cmd += ["--target-month", str(args.target_month)]
    if args.source_task_id:
        cmd += ["--source-task-id", str(args.source_task_id)]
    if args.dry_run:
        cmd += ["--dry-run"]
    print(f"[polaris-task] {' '.join(cmd)}", file=sys.stderr)
    r = subprocess.run(cmd, check=False)
    if r.returncode != 0:
        emit(envelope_err("polaris_task_failed", f"update_task exit={r.returncode}"))
        return r.returncode

    if args.dry_run:
        emit(envelope_ok({"dry_run": True}))
        return 0

    today_tag = dt.date.today().strftime("%Y%m%d")
    out = OUTPUT_DIR / f"polaris_task_{today_tag}.json"
    if not out.exists():
        emit(envelope_err("output_missing", f"未找到 {out}"))
        return 3
    import json as _json
    emit(envelope_ok({"polaris_task": _json.loads(out.read_text(encoding="utf-8")), "file": str(out)}))
    return 0


def cmd_all(args):
    """端到端跑完整套停课唤醒目标流程：
       export → filter → query-dadou-id → liuyi-tag → liuyi-group → wechat-tag → polaris-task

       --month YYYY-MM 同时控制：
         - export/filter 的签到时间窗口基准月
         - liuyi-tag/group 的标签和用户群命名月份
         - polaris-task 的目标月份（数字部分）
    """
    if args.month:
        y, m = args.month.split("-")
        target_month = int(m)
        month_arg = args.month
    else:
        today = dt.date.today()
        target_month = today.month
        month_arg = f"{today.year}-{today.month:02d}"

    print(f"\n{'=' * 70}", file=sys.stderr)
    print(f"  全流程开始 — 目标月份: {month_arg}", file=sys.stderr)
    print(f"{'=' * 70}\n", file=sys.stderr)

    # 每个步骤的描述 + 触发函数 + 触发时构造的 Namespace
    steps = [
        ("export", cmd_export, argparse.Namespace(month=month_arg, dry_run=False)),
        ("filter", cmd_filter, argparse.Namespace(month=month_arg, input=None)),
        ("query-dadou-id", cmd_query_dadou_id, argparse.Namespace(
            input=None,
            import_type="豌豆用户",
            export_type="豌豆大账号",
            poll_timeout=180,
        )),
        ("liuyi-tag", cmd_liuyi_tag, argparse.Namespace(
            filtered=None, dadou=None, month=month_arg,
        )),
        ("liuyi-group", cmd_liuyi_group, argparse.Namespace(
            month=month_arg, tag_ids_json=None, dry_run=False,
        )),
        ("wechat-tag", cmd_wechat_tag, argparse.Namespace(
            group_ids_json=None, corp_tag_group_id=None, dry_run=False,
        )),
        ("polaris-task", cmd_polaris_task, argparse.Namespace(
            keyword=None, target_month=target_month, source_task_id=None, dry_run=False,
        )),
    ]

    failed_at = None
    for name, fn, ns in steps:
        print(f"\n>>> [{name}] 开始", file=sys.stderr)
        rc = fn(ns)
        if rc != 0:
            failed_at = name
            print(f"<<< [{name}] 失败，退出码 {rc}", file=sys.stderr)
            break
        print(f"<<< [{name}] 完成", file=sys.stderr)

    print(f"\n{'=' * 70}", file=sys.stderr)
    if failed_at:
        print(f"  ❌ 全流程失败于步骤 [{failed_at}]", file=sys.stderr)
        emit(envelope_err("all_failed", f"failed at step {failed_at}"))
        return 1
    print(f"  ✅ 全流程完成 — month={month_arg}", file=sys.stderr)
    print(f"{'=' * 70}\n", file=sys.stderr)
    today_tag = dt.date.today().strftime("%Y%m%d")
    summary = {
        "month": month_arg,
        "outputs": {
            "raw":     str(OUTPUT_DIR / f"raw_{today_tag}.xlsx"),
            "filtered":str(OUTPUT_DIR / f"filtered_{today_tag}.xlsx"),
            "dadou_mapping": str(OUTPUT_DIR / f"dadou_mapping_{today_tag}.xlsx"),
            "tag_ids": str(OUTPUT_DIR / f"tag_ids_{today_tag}.json"),
            "group_ids": str(OUTPUT_DIR / f"group_ids_{today_tag}.json"),
            "wechat_tag": str(OUTPUT_DIR / f"wechat_tag_{today_tag}.json"),
            "polaris_task": str(OUTPUT_DIR / f"polaris_task_{today_tag}.json"),
        }
    }
    emit(envelope_ok(summary))
    return 0


def main():
    p = argparse.ArgumentParser(prog="tingke_wakeup")
    sub = p.add_subparsers(dest="step", required=True)

    p_export = sub.add_parser("export", help="从 BI 拉思维停课学员执行明细")
    p_export.add_argument("--month", help="导出基准月 YYYY-MM，默认当月")
    p_export.add_argument("--dry-run", action="store_true", help="只打印参数不真跑")
    p_export.set_defaults(func=cmd_export)

    p_filter = sub.add_parser("filter", help="二次筛选：签到时间窗口 + 是否停课唤醒=0")
    p_filter.add_argument("--month", help="导出基准月 YYYY-MM，默认当月")
    p_filter.add_argument("--input", help="输入 xlsx 路径，默认 output/raw_{today}.xlsx")
    p_filter.set_defaults(func=cmd_filter)

    p_qd = sub.add_parser("query-dadou-id",
                          help="平台中心批量用户查询：学员ID → 豌豆大账号ID 映射")
    p_qd.add_argument("--input", help="filtered xlsx 路径，默认 output/filtered_{today}.xlsx")
    p_qd.add_argument("--import-type", default="豌豆用户", help="导入类型（默认 豌豆用户）")
    p_qd.add_argument("--export-type", default="豌豆大账号", help="导出类型（默认 豌豆大账号）")
    p_qd.add_argument("--poll-timeout", type=int, default=120, help="轮询处理结果秒数（默认 120）")
    p_qd.set_defaults(func=cmd_query_dadou_id)

    p_lt = sub.add_parser("liuyi-tag", help="六一工作台 新建标签 ×2（小账号 + 大账号）")
    p_lt.add_argument("--filtered", help="filtered xlsx，默认 output/filtered_{today}.xlsx")
    p_lt.add_argument("--dadou", help="dadou_mapping xlsx，默认 output/dadou_mapping_{today}.xlsx")
    p_lt.add_argument("--month", help="标签命名月份 YYYY-MM，默认按当前月")
    p_lt.set_defaults(func=cmd_liuyi_tag)

    p_lg = sub.add_parser("liuyi-group", help="六一工作台 新建用户群 ×2（依赖 liuyi-tag 产出）")
    p_lg.add_argument("--month", help="用户群命名月份 YYYY-MM，默认按当前月")
    p_lg.add_argument("--tag-ids-json", help="tag_ids json，默认 output/tag_ids_{today}.json")
    p_lg.add_argument("--dry-run", action="store_true", help="只构造请求体，不调 add")
    p_lg.set_defaults(func=cmd_liuyi_group)

    p_wt = sub.add_parser("wechat-tag", help="企微：把大账号用户群关联到「【益智】长期标签」标签组")
    p_wt.add_argument("--group-ids-json", help="group_ids json，默认 output/group_ids_{today}.json")
    p_wt.add_argument("--corp-tag-group-id", help="企微标签组 id（默认「【益智】长期标签」）")
    p_wt.add_argument("--dry-run", action="store_true", help="只构造请求体，不调 create")
    p_wt.set_defaults(func=cmd_wechat_tag)

    p_pt = sub.add_parser("polaris-task",
                          help="北极星：克隆「停课120天以内-N月」任务到目标月份")
    p_pt.add_argument("--keyword", help="任务名关键词（默认「停课120天以内」）")
    p_pt.add_argument("--target-month", type=int, help="目标月份 1-12（默认下月）")
    p_pt.add_argument("--source-task-id", type=int, help="基线任务 id（默认按 keyword 搜最新）")
    p_pt.add_argument("--dry-run", action="store_true", help="只构造 payload 不调 add")
    p_pt.set_defaults(func=cmd_polaris_task)

    p_all = sub.add_parser("all", help="端到端跑全流程：export→filter→query-dadou-id→liuyi-tag→liuyi-group→wechat-tag→polaris-task")
    p_all.add_argument("--month", help="目标月份 YYYY-MM，默认 Windows 当月")
    p_all.set_defaults(func=cmd_all)

    args = p.parse_args()
    sys.exit(args.func(args))


if __name__ == "__main__":
    main()
