"""用渐变背景色模拟数据条：每列按值大小从白到蓝渐变上色。"""
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
    cols = []
    for i, cell in enumerate(header_row):
        v = str(cell).strip()
        if "占比" in v or "率" in v or "跟进率" in v:
            cols.append(i)
    return cols


def to_float(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("%", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def lerp_color(ratio):
    """绿(1.0) → 黄(0.5) → 红(0.0)：ratio 越大越绿（好），越小越红（差）"""
    if ratio >= 0.5:
        # 黄→绿：ratio 0.5→1.0
        t = (ratio - 0.5) * 2
        r = int(255 * (1 - t) + 76 * t)   # 255→76
        g = int(215 * (1 - t) + 175 * t)  # 215→175
        b = int(0 * (1 - t) + 80 * t)     # 0→80
    else:
        # 红→黄：ratio 0.0→0.5
        t = ratio * 2
        r = int(244 * (1 - t) + 255 * t)  # 244→255
        g = int(67 * (1 - t) + 215 * t)   # 67→215
        b = int(54 * (1 - t) + 0 * t)     # 54→0
    return f"#{r:02X}{g:02X}{b:02X}"


def run(cmd, timeout=120):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=(sys.platform == "win32"), timeout=timeout)
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


def batch_set_style(data):
    dj = json.dumps(data, ensure_ascii=False)
    if len(dj) > 7500:
        half = len(data) // 2
        return batch_set_style(data[:half]) and batch_set_style(data[half:])
    ok, resp = run([LARK_CLI, "sheets", "+batch-set-style",
                    "--spreadsheet-token", SPREADSHEET_TOKEN, "--data", dj])
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
        n_rows = len(rows)
        n_header = HEADER_ROWS[title]
        detail_row = rows[n_header - 1]
        pct_cols = find_pct_cols(detail_row)
        if not pct_cols:
            print(f"[{title}] 无数据条列，跳过")
            continue

        print(f"[{title}] 处理 {len(pct_cols)} 列色阶...")
        data_rows = rows[n_header:]
        ops = []

        for col_idx in pct_cols:
            values = []
            for row_i, r in enumerate(data_rows):
                padded = list(r) + [""] * (max(len(r2) for r2 in rows) - len(r))
                cell = padded[col_idx] if col_idx < len(padded) else ""
                v = to_float(cell)
                values.append((row_i, v))

            nums = [v for _, v in values if v is not None and v != 0]
            if not nums:
                continue
            min_v = min(nums)
            max_v = max(nums)
            if max_v == min_v:
                continue

            col_l = col_letter(col_idx + 1)
            for row_i, v in values:
                if v is None or v == 0:
                    continue
                ratio = (v - min_v) / (max_v - min_v)
                color = lerp_color(ratio)
                actual_row = n_header + 1 + row_i
                rng = f"{sid}!{col_l}{actual_row}:{col_l}{actual_row}"
                ops.append({"ranges": [rng], "style": {"backColor": color}})

        # 批量提交（每批最多 30 个 range ops）
        total_ops = len(ops)
        sent = 0
        for i in range(0, total_ops, 30):
            batch = ops[i:i + 30]
            if batch_set_style(batch):
                sent += len(batch)
            time.sleep(0.3)
        print(f"  ✓ {sent}/{total_ops} 单元格上色")

    print("\n[DONE] 色阶模拟数据条完成")
    print(f"表格：https://hcnig43mb8gp.feishu.cn/sheets/{SPREADSHEET_TOKEN}")


if __name__ == "__main__":
    main()
