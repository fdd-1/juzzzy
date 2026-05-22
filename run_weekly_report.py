"""
周报生成调度脚本 - 串联读取、分析、写入全流程
用法: python run_weekly_report.py --source-token <spreadsheet_token> [--folder <folder_token>]
"""
import argparse
import sys
from pathlib import Path
from datetime import date, timedelta

sys.path.insert(0, str(Path(__file__).parent))

from main import load_config, read_sheet, parse_table, get_week_range, fmt_pct
from analyzer import (
    find_total_row, find_group_rows, get_metric_value,
    analyze_anomalies, compare_with_previous, _fmt_pct,
)
from doc_writer import create_doc, write_section


def run(source_token: str, folder_token: str = None, dry_run: bool = False):
    config = load_config()
    week = get_week_range()
    print(f"[INFO] 生成周报: {week['display']}")
    print(f"[INFO] 数据源表格: {source_token}")

    results = {}
    for section in config["sections"]:
        sid = section["id"]
        title = section["title"]
        print(f"\n[SECTION] {title}")

        sheets_cfg = section["sheets"]
        if isinstance(sheets_cfg, dict):
            current_data = {}
            previous_data = {}
            for key, sheet_id in sheets_cfg.items():
                if "previous" in key or "prev" in key:
                    raw = read_sheet(source_token, sheet_id)
                    if raw:
                        previous_data[key] = parse_table(raw)
                else:
                    raw = read_sheet(source_token, sheet_id)
                    if raw:
                        current_data[key] = parse_table(raw)
        else:
            continue

        results[sid] = {
            "title": title,
            "config": section,
            "current": current_data,
            "previous": previous_data,
        }

    if dry_run:
        print("\n[DRY RUN] 数据读取完成，以下是各模块汇总:")
        for sid, r in results.items():
            print(f"  {r['title']}:")
            for k, v in r["current"].items():
                print(f"    {k}: {len(v)} 行数据")
        return results

    # 创建新文档
    doc_title = f"海外教学服务部周报 {week['display']}"
    print(f"\n[INFO] 创建文档: {doc_title}")
    doc_id = create_doc(doc_title, folder_token)
    if not doc_id:
        print("[ERROR] 创建文档失败", file=sys.stderr)
        return None
    print(f"[OK] 文档已创建: {doc_id}")

    # 逐模块写入
    for sid, r in results.items():
        section_cfg = r["config"]
        title = r["title"]
        curr = r["current"]
        prev = r["previous"]

        conclusion_lines = generate_conclusion(sid, curr, prev, section_cfg)
        table_headers, table_rows = extract_display_table(curr, section_cfg)
        prev_headers, prev_rows = extract_display_table(prev, section_cfg)

        success = write_section(
            doc_id, title, conclusion_lines,
            table_headers, table_rows,
            prev_headers, prev_rows,
        )
        status = "OK" if success else "FAIL"
        print(f"  [{status}] {title}")

    print(f"\n[DONE] 周报已生成: https://my.feishu.cn/docx/{doc_id}")
    return doc_id


def generate_conclusion(sid: str, current: dict, previous: dict, config: dict) -> list[str]:
    """根据模块 ID 生成对应的结论文本"""
    key_col = config.get("key_col", "小组")
    lines = []

    curr_records = list(current.values())[0] if current else []
    prev_records = list(previous.values())[0] if previous else []

    if not curr_records:
        return ["数据缺失"]

    total = find_total_row(curr_records, key_col)
    if not total:
        return ["未找到汇总行"]

    metrics = config.get("metrics", [])
    for m in metrics:
        name = m["name"]
        field = m["fields"][0]
        val = get_metric_value(total, field)
        target = m.get("target")

        if val is not None:
            val_str = _fmt_pct(val)
            line = f"{name}：{val_str}"
            if target is not None:
                status = "达标" if val >= target else "未达标"
                line += f"（目标{_fmt_pct(target)}，{status}）"
            comp = compare_with_previous(curr_records, prev_records, key_col, field)
            if comp["change"] is not None:
                direction = "上升" if comp["change"] > 0 else "下降"
                line += f"，环比{direction}{abs(comp['change'])*100:.1f}%"
            lines.append(line)

    # 异常组检测
    groups = find_group_rows(curr_records, key_col)
    for m in metrics:
        threshold = m.get("alert_threshold")
        if threshold is None:
            continue
        anomalies = analyze_anomalies(groups, key_col, m["fields"][0], threshold)
        for a in anomalies[:2]:
            lines.append(f"——{a['group']}{m['name']}仅{_fmt_pct(a['value'])}，需注意")

    return lines if lines else ["各项指标平稳"]


def extract_display_table(data: dict, config: dict) -> tuple[list[str], list[list]]:
    """从数据中提取展示表格的表头和行"""
    if not data:
        return [], []
    records = list(data.values())[0]
    if not records:
        return [], []

    key_col = config.get("key_col", "小组")
    lp_col = config.get("lp_col", "LP")
    metrics = config.get("metrics", [])

    display_fields = [key_col, lp_col]
    for m in metrics:
        display_fields.extend(m["fields"])
    display_fields = list(dict.fromkeys(display_fields))

    available_keys = set()
    for r in records:
        available_keys.update(r.keys())

    headers = []
    field_map = []
    for f in display_fields:
        for ak in available_keys:
            if f in ak:
                headers.append(ak.split("_")[-1] if "_" in ak else ak)
                field_map.append(ak)
                break

    rows = []
    for r in records[:15]:
        row = [r.get(fm) for fm in field_map]
        rows.append(row)

    return headers, rows


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="周报自动化 - 服务模块")
    parser.add_argument("--source-token", required=True, help="数据源飞书电子表格 token")
    parser.add_argument("--folder", default=None, help="目标文件夹 token")
    parser.add_argument("--dry-run", action="store_true", help="仅读取数据不写入")
    args = parser.parse_args()
    run(args.source_token, args.folder, args.dry_run)
