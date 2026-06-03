"""把占比/率列转百分比字符串，其他数字保留2位小数，重新写入飞书。"""
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


def safe_text(v):
    if isinstance(v, str):
        return v.replace("<", "＜").replace(">", "＞")
    return v


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


def find_pct_cols(header_row):
    cols = set()
    for i, cell in enumerate(header_row):
        v = str(cell).strip()
        if "占比" in v or "率" in v:
            cols.add(i)
    return cols


def format_cell(val, is_pct):
    """格式化单个单元格值"""
    if isinstance(val, str):
        v = val.strip()
        if v == "" or v == "0":
            return v
        # 尝试转数字
        try:
            num = float(v)
        except ValueError:
            return safe_text(v)
        val = num

    if isinstance(val, (int, float)):
        if val == 0:
            return "0"
        if is_pct:
            # 值本身是小数（如 0.373），转成百分比字符串
            if abs(val) <= 5:  # 合理范围内视为小数比例
                return f"{val * 100:.2f}%"
            else:
                return f"{val:.2f}%"  # 已经是百分比数字
        else:
            if isinstance(val, int) or val == int(val):
                return str(int(val))
            return f"{val:.2f}"
    return safe_text(str(val))


def write_batch(sid, batch, start_row, n_cols):
    end_col = col_letter(n_cols)
    end_row = start_row + len(batch) - 1
    rng = f"{sid}!A{start_row}:{end_col}{end_row}"
    values = json.dumps(batch, ensure_ascii=False)
    if len(values) > 7000:
        if len(batch) == 1:
            # 列分段
            row = batch[0]
            i = 0
            chunk = 15
            while i < len(row):
                sub = row[i:i + chunk]
                sj = json.dumps([sub], ensure_ascii=False)
                if len(sj) > 7000 and chunk > 1:
                    chunk = max(1, chunk // 2)
                    continue
                rg = f"{sid}!{col_letter(i+1)}{start_row}:{col_letter(i+len(sub))}{start_row}"
                run([LARK_CLI, "sheets", "+write",
                     "--spreadsheet-token", SPREADSHEET_TOKEN,
                     "--range", rg, "--values", sj])
                i += chunk
                time.sleep(0.15)
            return True
        half = len(batch) // 2
        return write_batch(sid, batch[:half], start_row, n_cols) and \
               write_batch(sid, batch[half:], start_row + half, n_cols)
    ok, _ = run([LARK_CLI, "sheets", "+write",
                 "--spreadsheet-token", SPREADSHEET_TOKEN,
                 "--range", rng, "--values", values])
    if not ok:
        time.sleep(1.0)
        ok, _ = run([LARK_CLI, "sheets", "+write",
                     "--spreadsheet-token", SPREADSHEET_TOKEN,
                     "--range", rng, "--values", values])
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
        n_header = HEADER_ROWS[title]
        detail_row = rows[n_header - 1]
        pct_cols = find_pct_cols(detail_row)
        # 文本列（前 2-3 列）不做数字格式化
        text_cols = set(range(min(3, n_cols)))

        print(f"[{title}] 格式化数据行 {n_header+1}-{len(rows)}...")

        # 只重写数据行（表头不动）
        data_rows = rows[n_header:]
        formatted = []
        for r in data_rows:
            padded = list(r) + [""] * (n_cols - len(r))
            new_row = []
            for ci, val in enumerate(padded):
                if ci in text_cols:
                    new_row.append(safe_text(str(val)) if val != "" else "")
                else:
                    new_row.append(format_cell(val, ci in pct_cols))
            formatted.append(new_row)

        # 分批写入
        batch_size = 8
        written = 0
        for i in range(0, len(formatted), batch_size):
            batch = formatted[i:i + batch_size]
            if write_batch(sid, batch, n_header + 1 + i, n_cols):
                written += len(batch)
            time.sleep(0.2)
        print(f"  ✓ {written}/{len(formatted)} 行")

    print("\n[DONE]")
    print(f"表格：https://hcnig43mb8gp.feishu.cn/sheets/{SPREADSHEET_TOKEN}")


if __name__ == "__main__":
    main()
