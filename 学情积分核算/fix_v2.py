"""修复版：用 shell=True 读数据，完成色阶+合并"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
DOC_ID = "TBX2dTCIqoc3zJxHVdXcA1bUnoh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def run_shell(cmd, timeout=120):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=True, timeout=timeout)
    try:
        return True, json.loads(r.stdout)
    except:
        return False, r.stdout


def read_range(sid, rng):
    ok, resp = run_shell([LARK_CLI, "sheets", "+read", "--spreadsheet-token",
                          SPREADSHEET_TOKEN, "--range", f"{sid}!{rng}"])
    if ok and isinstance(resp, dict) and resp.get("data"):
        return resp["data"].get("values", [])
    return []


def batch_set_style(data):
    dj = json.dumps(data, ensure_ascii=False)
    if len(dj) > 7500:
        half = len(data) // 2
        return batch_set_style(data[:half]) and batch_set_style(data[half:])
    ok, _ = run_shell([LARK_CLI, "sheets", "+batch-set-style",
                       "--spreadsheet-token", SPREADSHEET_TOKEN, "--data", dj])
    return ok


def merge_cells(sid, rng):
    run_shell([LARK_CLI, "sheets", "+merge-cells", "--spreadsheet-token",
               SPREADSHEET_TOKEN, "--range", f"{sid}!{rng}", "--merge-type", "MERGE_ALL"])
    time.sleep(0.15)


def lerp_color(ratio):
    if ratio >= 0.5:
        t = (ratio - 0.5) * 2
        r = int(255*(1-t)+76*t); g = int(215*(1-t)+175*t); b = int(0*(1-t)+80*t)
    else:
        t = ratio * 2
        r = int(244*(1-t)+255*t); g = int(67*(1-t)+215*t); b = int(54*(1-t)+0*t)
    return f"#{r:02X}{g:02X}{b:02X}"


def apply_color_to_col(sid, col_l, start_row, end_row):
    vals = read_range(sid, f"{col_l}{start_row}:{col_l}{end_row}")
    if not vals:
        return 0
    nums = []
    for row in vals:
        v = row[0] if row else None
        try:
            nums.append(float(str(v).replace("%", "").replace(",", "")))
        except:
            nums.append(None)
    valid = [n for n in nums if n is not None and n != 0]
    if not valid or max(valid) == min(valid):
        return 0
    mn, mx = min(valid), max(valid)
    ops = []
    for i, n in enumerate(nums):
        if n is None or n == 0:
            continue
        ratio = (n - mn) / (mx - mn)
        color = lerp_color(ratio)
        rng = f"{sid}!{col_l}{start_row + i}:{col_l}{start_row + i}"
        ops.append({"ranges": [rng], "style": {"backColor": color}})
    for i in range(0, len(ops), 30):
        batch_set_style(ops[i:i+30])
        time.sleep(0.2)
    return len(ops)


def group_merge(sid, col_l, start_row, end_row):
    vals = read_range(sid, f"{col_l}{start_row}:{col_l}{end_row}")
    if not vals:
        return 0
    cur_val = None; cur_start = None; merges = []
    for i, row in enumerate(vals):
        v = str(row[0]).strip() if row and row[0] else ""
        if v and v != cur_val:
            if cur_val and cur_start is not None and i - cur_start > 1:
                merges.append((cur_start + start_row, i + start_row - 1))
            cur_val = v; cur_start = i
    if cur_val and cur_start is not None and len(vals) - cur_start > 1:
        merges.append((cur_start + start_row, len(vals) + start_row - 1))
    for s, e in merges:
        merge_cells(sid, f"{col_l}{s}:{col_l}{e}")
    return len(merges)


def main():
    # 1) 4.5 色阶
    print("[1] 4.5 色阶（外呼跟进率D + 综合有效跟进率I）...")
    n1 = apply_color_to_col("2T3bYQ", "D", 3, 106)
    n2 = apply_color_to_col("2T3bYQ", "I", 3, 106)
    print(f"  D: {n1}, I: {n2}")

    # 2) 4.4 唤醒率色阶
    print("[2] 4.4 唤醒率色阶...")
    r3 = read_range("2SOCo8", "A3:BV3")
    if r3:
        header = r3[0] if r3 else []
        wake_cols = [i for i, v in enumerate(header) if v and "唤醒率" in str(v)]
        print(f"  唤醒率列: {wake_cols}")
        for ci in wake_cols:
            cl = col_letter(ci + 1)
            n = apply_color_to_col("2SOCo8", cl, 4, 99)
            print(f"    {cl}: {n}")

    # 3) 4.4 小组向下合并
    print("[3] 4.4 小组合并...")
    n = group_merge("2SOCo8", "A", 4, 99)
    print(f"  {n} groups")

    # 4) 4.6 小组向下合并
    print("[4] 4.6 小组合并...")
    n1 = group_merge("2TwScg", "A", 3, 98)
    n2 = group_merge("2TLJU4", "A", 3, 290)
    print(f"  系统外呼: {n1}, 企微回复比: {n2}")

    # 5) 文档 overwrite
    print("[5] 文档 overwrite...")
    from pathlib import Path
    xml_path = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_doc.xml")
    if xml_path.exists():
        xml = xml_path.read_text(encoding="utf-8")
        r = subprocess.run([LARK_CLI, "docs", "+update", "--api-version", "v2",
                            "--doc", DOC_ID, "--command", "overwrite",
                            "--doc-format", "xml", "--content", "-"],
                           capture_output=True, text=True, encoding="utf-8",
                           shell=False, timeout=300, input=xml)
        try:
            resp = json.loads(r.stdout)
            print(f"  result={resp.get('data',{}).get('result')}")
        except:
            print(f"  output: {r.stdout[:200]}")

    print("\n[DONE]")


if __name__ == "__main__":
    main()
