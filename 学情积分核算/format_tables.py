"""完整格式化：居中 + 合并表头 + 百分比列 + 2位小数 + 颜色 + 数据条"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

SHEET_ID_MAP = {
    "4.1 服务指标": "2RK2qc", "4.1 AI学情": "2S16wg",
    "4.2 组班意向": "2SgAAo", "4.3 群发消息": "2SwpXO",
    "4.4 停课监控": "2SOCo8", "4.5 服务池跟进": "2T3bYQ",
    "4.5 服务池SOP": "2Ti90Y", "4.6 系统外呼监控": "2TwScg",
    "4.6 企微回复比": "2TLJU4",
}
HEADER_ROWS = {
    "4.1 服务指标": 2, "4.1 AI学情": 2, "4.2 组班意向": 2,
    "4.3 群发消息": 2, "4.4 停课监控": 3, "4.5 服务池跟进": 2,
    "4.5 服务池SOP": 2, "4.6 系统外呼监控": 2, "4.6 企微回复比": 2,
}


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def run(cmd, timeout=120):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=(sys.platform == "win32"), timeout=timeout)
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


def find_caliber_start(rows):
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
    """找连续相同值或空值的合并组 → [(start_1based, end_1based_inclusive, value)]"""
    groups = []
    cur_val = None
    cur_start = 0
    for i, cell in enumerate(row):
        v = str(cell).strip()
        if v == "":
            continue
        if v != cur_val:
            if cur_val is not None and i - cur_start > 1:
                groups.append((cur_start + 1, i, cur_val))
            cur_val = v
            cur_start = i
    if cur_val is not None and len(row) - cur_start > 1:
        groups.append((cur_start + 1, len(row), cur_val))
    return groups


def find_pct_cols(header_row):
    cols = []
    for i, cell in enumerate(header_row):
        v = str(cell).strip()
        if "占比" in v or "率" in v:
            cols.append(i + 1)
    return cols


def set_style(sid, rng, style):
    ok, resp = run([LARK_CLI, "sheets", "+set-style",
                    "--spreadsheet-token", SPREADSHEET_TOKEN,
                    "--range", f"{sid}!{rng}", "--style", json.dumps(style, ensure_ascii=False)])
    return ok


def batch_set_style(data):
    dj = json.dumps(data, ensure_ascii=False)
    if len(dj) > 7500:
        half = len(data) // 2
        return batch_set_style(data[:half]) and batch_set_style(data[half:])
    ok, resp = run([LARK_CLI, "sheets", "+batch-set-style",
                    "--spreadsheet-token", SPREADSHEET_TOKEN, "--data", dj])
    if not ok:
        print(f"    batch-set-style FAIL: {str(resp)[:200]}")
    return ok


def merge_cells(sid, rng):
    ok, resp = run([LARK_CLI, "sheets", "+merge-cells",
                    "--spreadsheet-token", SPREADSHEET_TOKEN,
                    "--range", f"{sid}!{rng}", "--merge-type", "MERGE_ALL"])
    return ok


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    excel_data_map = {s["title"]: s["rows"] for s in extracted["sheets"]}

    for title, sid in SHEET_ID_MAP.items():
        rows = excel_data_map[title]
        cut = find_caliber_start(rows)
        rows = [r for r in rows[:cut] if any(str(c).strip() for c in r)]
        if not rows:
            continue
        n_cols = max(len(r) for r in rows)
        n_rows = len(rows)
        n_header = HEADER_ROWS[title]
        end_col = col_letter(n_cols)
        detail_row = rows[n_header - 1]

        print(f"\n{'='*50}")
        print(f"[{title}] sid={sid}, {n_rows}×{n_cols}, header={n_header}")

        # 1) 全表居中
        print("  [1] 全表居中...")
        set_style(sid, f"A1:{end_col}{n_rows}", {"hAlign": 1, "vAlign": 1})
        time.sleep(0.4)

        # 2) 合并表头第 1 行
        print("  [2] 合并表头...")
        r1 = rows[0]
        groups = get_merge_groups(r1)
        multi = [(s, e) for s, e, v in groups if e - s > 1]
        for s_col, e_col in multi:
            rng = f"{col_letter(s_col)}1:{col_letter(e_col)}1"
            merge_cells(sid, rng)
            time.sleep(0.2)
        print(f"    合并了 {len(multi)} 组")

        # 4.4 有 3 行表头，第 2 行也可能需要合并
        if n_header >= 3 and len(rows) >= 2:
            r2 = rows[1]
            groups2 = get_merge_groups(r2)
            multi2 = [(s, e) for s, e, v in groups2 if e - s > 1]
            for s_col, e_col in multi2:
                rng = f"{col_letter(s_col)}2:{col_letter(e_col)}2"
                merge_cells(sid, rng)
                time.sleep(0.2)
            if multi2:
                print(f"    第 2 行合并了 {len(multi2)} 组")

        # 3) 百分比列格式 + 其他数字列 2 位小数
        print("  [3] 数字格式...")
        pct_cols = find_pct_cols(detail_row)
        all_cols = set(range(1, n_cols + 1))
        # 前 2-3 列是文本（层级/团队/LP），跳过
        text_cols = set(range(1, min(4, n_cols + 1)))
        num_cols = list(all_cols - text_cols - set(pct_cols))

        ops = []
        # 百分比列
        if pct_cols:
            pct_ranges = [f"{sid}!{col_letter(c)}{n_header+1}:{col_letter(c)}{n_rows}" for c in pct_cols]
            # 分批（ranges 不能太多）
            for i in range(0, len(pct_ranges), 20):
                ops.append({"ranges": pct_ranges[i:i+20], "style": {"formatter": "0.00%"}})

        # 数字列 2 位小数
        if num_cols:
            num_ranges = [f"{sid}!{col_letter(c)}{n_header+1}:{col_letter(c)}{n_rows}" for c in sorted(num_cols)]
            for i in range(0, len(num_ranges), 20):
                ops.append({"ranges": num_ranges[i:i+20], "style": {"formatter": "0.00"}})

        if ops:
            for op in ops:
                batch_set_style([op])
                time.sleep(0.3)
        print(f"    百分比列: {len(pct_cols)}, 数字列: {len(num_cols)}")

        # 4) 重新上色（合并可能清掉了之前的样式）
        print("  [4] 重新上色...")
        # 表头：粗体 + 浅蓝 + 居中
        ops_color = [
            {"ranges": [f"{sid}!A1:{end_col}{n_header}"],
             "style": {"font": {"bold": True, "font_size": 10},
                       "backColor": "#E1F0FE", "hAlign": 1, "vAlign": 1}},
        ]
        # 总计行：粗体 + 浅黄
        total_idx = []
        for i, r in enumerate(rows):
            if i < n_header:
                continue
            cells4 = [str(c).strip() for c in r[:4]]
            if "总计" in cells4:
                total_idx.append(i + 1)
        if total_idx:
            total_ranges = [f"{sid}!A{r}:{end_col}{r}" for r in total_idx]
            ops_color.append({"ranges": total_ranges, "style": {"font": {"bold": True}, "backColor": "#FFF8DC"}})

        batch_set_style(ops_color)
        time.sleep(0.4)

        # 5) 列宽
        run([LARK_CLI, "sheets", "+update-dimension",
             "--spreadsheet-token", SPREADSHEET_TOKEN,
             "--sheet-id", sid, "--dimension", "COLUMNS",
             "--start-index", "1", "--end-index", "2", "--fixed-size", "130"])
        if n_cols >= 3:
            run([LARK_CLI, "sheets", "+update-dimension",
                 "--spreadsheet-token", SPREADSHEET_TOKEN,
                 "--sheet-id", sid, "--dimension", "COLUMNS",
                 "--start-index", "3", "--end-index", str(n_cols), "--fixed-size", "110"])
        time.sleep(0.3)
        print("  [5] 列宽 OK")

    print("\n\n[DONE] 全部格式化完成")
    print(f"表格：https://hcnig43mb8gp.feishu.cn/sheets/{SPREADSHEET_TOKEN}")


if __name__ == "__main__":
    main()
