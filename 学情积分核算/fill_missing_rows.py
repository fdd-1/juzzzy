"""逐行补齐失败的行：每行单独 +write，单行如果还超长就只写有内容的列段。"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

# 需要补齐的：sheet_id, 起始行(1-based), 结束行(inclusive)
NEED_FIX = [
    ("4.4 停课监控", "2SOCo8", 97, 115),
    ("4.5 服务池跟进", "2T3bYQ", 121, 128),
    ("4.5 服务池SOP", "2Ti90Y", 97, 112),
    ("4.6 系统外呼监控", "2TwScg", 105, 110),
    ("4.6 企微回复比", "2TLJU4", 289, 297),
]


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


def write_one_row(sid, row, row_num):
    """写单行；若整行 JSON 太长，按列拆"""
    n = len(row)
    full_json = json.dumps([row], ensure_ascii=False)
    if len(full_json) <= 6500:
        end_col = col_letter(n)
        rng = f"{sid}!A{row_num}:{end_col}{row_num}"
        ok, resp = run([LARK_CLI, "sheets", "+write",
                        "--spreadsheet-token", SPREADSHEET_TOKEN,
                        "--range", rng, "--values", full_json])
        if ok:
            return True
        time.sleep(1.5)
        ok, resp = run([LARK_CLI, "sheets", "+write",
                        "--spreadsheet-token", SPREADSHEET_TOKEN,
                        "--range", rng, "--values", full_json])
        if ok:
            return True
        # fall through to chunking
    # 列分段写
    chunk = 20
    i = 0
    while i < n:
        sub = row[i:i + chunk]
        sub_json = json.dumps([sub], ensure_ascii=False)
        if len(sub_json) > 6500 and chunk > 1:
            chunk = max(1, chunk // 2)
            continue
        start_c = col_letter(i + 1)
        end_c = col_letter(i + len(sub))
        rng = f"{sid}!{start_c}{row_num}:{end_c}{row_num}"
        ok, resp = run([LARK_CLI, "sheets", "+write",
                        "--spreadsheet-token", SPREADSHEET_TOKEN,
                        "--range", rng, "--values", sub_json])
        if not ok:
            time.sleep(1.5)
            ok, resp = run([LARK_CLI, "sheets", "+write",
                            "--spreadsheet-token", SPREADSHEET_TOKEN,
                            "--range", rng, "--values", sub_json])
            if not ok:
                # 单单元格降级
                if len(sub) > 1:
                    # 拆半
                    half = len(sub) // 2
                    a_ok = True
                    for c, cell in enumerate(sub):
                        cell_json = json.dumps([[cell]], ensure_ascii=False)
                        rg = f"{sid}!{col_letter(i + 1 + c)}{row_num}:{col_letter(i + 1 + c)}{row_num}"
                        ok2, _ = run([LARK_CLI, "sheets", "+write",
                                      "--spreadsheet-token", SPREADSHEET_TOKEN,
                                      "--range", rg, "--values", cell_json])
                        a_ok = a_ok and ok2
                        time.sleep(0.15)
                    if not a_ok:
                        print(f"      [ERR] row {row_num} cols {i+1}-{i+len(sub)}: 部分单元格失败")
                        return False
                else:
                    print(f"      [ERR] row {row_num} col {i+1}: 单格写失败 {str(resp)[:120]}")
                    return False
        i += chunk
        time.sleep(0.15)
    return True


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    excel_data_map = {s["title"]: s["rows"] for s in extracted["sheets"]}
    for excel_name, sid, start, end in NEED_FIX:
        rows = excel_data_map[excel_name]
        # 补齐列宽到本表统一列数
        max_cols = max(len(r) for r in rows)
        rows = [list(r) + [""] * (max_cols - len(r)) for r in rows]
        print(f"\n=== 补齐 {excel_name} (sid={sid}) 行 {start}-{end} ===")
        ok_count = 0
        for row_num in range(start, end + 1):
            r = rows[row_num - 1]
            r = [v if isinstance(v, (int, float)) else str(v) for v in r]
            if write_one_row(sid, r, row_num):
                ok_count += 1
                print(f"  ✓ row {row_num}")
            else:
                print(f"  ✗ row {row_num}")
        print(f"  total: {ok_count}/{end - start + 1}")


if __name__ == "__main__":
    main()
