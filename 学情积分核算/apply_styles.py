"""仅做样式：表头加粗+浅蓝；总计行加粗+浅黄。用 batch-set-style 一次写完。"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

SHEET_ID_MAP = {
    "4.1 服务指标": "2RK2qc",
    "4.1 AI学情": "2S16wg",
    "4.2 组班意向": "2SgAAo",
    "4.3 群发消息": "2SwpXO",
    "4.4 停课监控": "2SOCo8",
    "4.5 服务池跟进": "2T3bYQ",
    "4.5 服务池SOP": "2Ti90Y",
    "4.6 系统外呼监控": "2TwScg",
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
    import re as _re
    for i, r in enumerate(rows):
        first = str(r[0]) if r else ""
        if any(kw in first for kw in ["口径说明", "注意：", "说明："]):
            return i
        if _re.match(r"^\d+[、）)\.]\s*[^,]+[:：]", first):
            other = [str(c) for c in r[1:] if str(c).strip()]
            if len(other) <= 2:
                return i
    return len(rows)


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    excel_data_map = {s["title"]: s["rows"] for s in extracted["sheets"]}

    for excel_name, sid in SHEET_ID_MAP.items():
        rows = excel_data_map[excel_name]
        cut = find_caliber_start(rows)
        rows = [r for r in rows[:cut] if any(str(c).strip() for c in r)]
        if not rows:
            continue
        n_cols = max(len(r) for r in rows)
        n_header = HEADER_ROWS[excel_name]
        end_col = col_letter(n_cols)

        # 找总计行
        total_idx = []
        for i, r in enumerate(rows):
            if i < n_header:
                continue
            cells4 = [str(c).strip() for c in r[:4]]
            if "总计" in cells4:
                total_idx.append(i + 1)

        # 构造 batch-set-style 的 data
        ops = []
        # 表头：粗体 + 浅蓝
        ops.append({
            "ranges": [f"{sid}!A1:{end_col}{n_header}"],
            "style": {"font": {"bold": True, "font_size": 10},
                      "backColor": "#E1F0FE",
                      "foreColor": "#1F2D3D"},
        })
        # 总计行：粗体 + 浅黄。把所有 total_idx 合到一个 ranges
        if total_idx:
            ranges = [f"{sid}!A{r}:{end_col}{r}" for r in total_idx]
            ops.append({
                "ranges": ranges,
                "style": {"font": {"bold": True}, "backColor": "#FFF8DC"},
            })

        data_json = json.dumps(ops, ensure_ascii=False)
        ok, resp = run([LARK_CLI, "sheets", "+batch-set-style",
                        "--spreadsheet-token", SPREADSHEET_TOKEN,
                        "--data", data_json])
        if ok:
            print(f"  ✓ {excel_name}: header + {len(total_idx)} total rows")
        else:
            print(f"  ✗ {excel_name}: {str(resp)[:300]}")
        time.sleep(0.4)

        # 列宽
        ok1, _ = run([LARK_CLI, "sheets", "+update-dimension",
                      "--spreadsheet-token", SPREADSHEET_TOKEN,
                      "--sheet-id", sid, "--dimension", "COLUMNS",
                      "--start-index", "1", "--end-index", "2", "--fixed-size", "130"])
        time.sleep(0.2)
        if n_cols >= 3:
            ok2, _ = run([LARK_CLI, "sheets", "+update-dimension",
                          "--spreadsheet-token", SPREADSHEET_TOKEN,
                          "--sheet-id", sid, "--dimension", "COLUMNS",
                          "--start-index", "3", "--end-index", str(n_cols), "--fixed-size", "110"])
        print(f"     列宽设置: cols 1-2=130px, 3-{n_cols}=110px")
        time.sleep(0.3)


if __name__ == "__main__":
    main()
