"""分析每个 sheet 第 1 行的合并区间：连续相同值或空值归到前一组。"""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

d = json.loads(Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json").read_text(encoding="utf-8"))

HEADER_ROWS = {
    "4.1 服务指标": 2, "4.1 AI学情": 2, "4.2 组班意向": 2,
    "4.3 群发消息": 2, "4.4 停课监控": 3, "4.5 服务池跟进": 2,
    "4.5 服务池SOP": 2, "4.6 系统外呼监控": 2, "4.6 企微回复比": 2,
}


def find_caliber_start(rows):
    import re
    for i, r in enumerate(rows):
        first = str(r[0]) if r else ""
        if any(kw in first for kw in ["口径说明", "注意：", "说明："]):
            return i
        if re.match(r"^\d+[、）)\.]\s*[^,]+[:：]", first):
            other = [str(c) for c in r[1:] if str(c).strip()]
            if len(other) <= 2:
                return i
    return len(rows)


def get_merge_groups(row):
    """找出第 1 行的合并组：[(start_col_1based, end_col_1based, value), ...]"""
    groups = []
    cur_val = None
    cur_start = None
    for i, cell in enumerate(row):
        v = str(cell).strip()
        if v == "":
            continue  # 空格归到前一组
        if v != cur_val:
            if cur_val is not None:
                groups.append((cur_start, i, cur_val))  # end is exclusive
            cur_val = v
            cur_start = i
    if cur_val is not None:
        groups.append((cur_start, len(row), cur_val))
    # 转成 1-based
    return [(s + 1, e, v) for s, e, v in groups]


def find_pct_cols(header_row):
    """找第 2（或第 3）行中含'占比'或'率'的列索引（1-based）"""
    cols = []
    for i, cell in enumerate(header_row):
        v = str(cell).strip()
        if "占比" in v or "率" in v:
            cols.append(i + 1)
    return cols


for s in d["sheets"]:
    title = s["title"]
    rows = s["rows"]
    cut = find_caliber_start(rows)
    rows = [r for r in rows[:cut] if any(str(c).strip() for c in r)]
    n_cols = max(len(r) for r in rows) if rows else 0
    n_header = HEADER_ROWS[title]
    print(f"{'='*60}")
    print(f"[{title}] {len(rows)} rows × {n_cols} cols, header={n_header}")

    # 第 1 行合并组
    r1 = rows[0] if rows else []
    groups = get_merge_groups(r1)
    multi_col_groups = [(s, e, v) for s, e, v in groups if e - s > 1]
    if multi_col_groups:
        print(f"  合并组（跨 ≥2 列）:")
        for s_col, e_col, val in multi_col_groups:
            print(f"    col {s_col}-{e_col-1}: '{val}'")
    else:
        print(f"  无多列合并（第 1 行每格都不同）")

    # 百分比列
    detail_row = rows[n_header - 1] if len(rows) >= n_header else rows[-1]
    pct_cols = find_pct_cols(detail_row)
    print(f"  占比/率列 (row {n_header}): {pct_cols[:20]}{'...' if len(pct_cols) > 20 else ''}")
    print()
