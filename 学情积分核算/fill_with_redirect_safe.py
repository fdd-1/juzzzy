"""shell=False 重写最后一批失败行：含 > / < 的字符串在 shell=True 时会被当重定向。"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

NEED_FIX = [
    ("4.4 停课监控", "2SOCo8", [104, 105, 108, 114]),
    ("4.5 服务池SOP", "2Ti90Y", [102]),
    ("4.6 系统外呼监控", "2TwScg", [106, 107, 108]),
]


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def run_no_shell(cmd, timeout=120):
    # 找到 lark-cli 的真实路径，用绝对路径调用，shell=False 避免重定向问题
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=False, timeout=timeout)
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    excel_data_map = {s["title"]: s["rows"] for s in extracted["sheets"]}

    cli_path = LARK_CLI
    print(f"lark-cli path: {cli_path}")

    for sheet, sid, row_nums in NEED_FIX:
        rows = excel_data_map[sheet]
        n_cols = max(len(r) for r in rows)
        rows = [list(r) + [""] * (n_cols - len(r)) for r in rows]
        print(f"\n=== {sheet} (sid={sid}) ===")
        for rn in row_nums:
            r = rows[rn - 1]
            r = [v if isinstance(v, (int, float)) else str(v) for v in r]
            end_col = col_letter(n_cols)
            rng = f"{sid}!A{rn}:{end_col}{rn}"
            values_json = json.dumps([r], ensure_ascii=False)
            cmd = [cli_path, "sheets", "+write",
                   "--spreadsheet-token", SPREADSHEET_TOKEN,
                   "--range", rng, "--values", values_json]
            ok, resp = run_no_shell(cmd)
            print(f"  row {rn}: {'OK' if ok else 'FAIL ' + str(resp)[:200]}")
            time.sleep(0.3)


if __name__ == "__main__":
    main()
