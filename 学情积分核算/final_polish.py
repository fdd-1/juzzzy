"""最终扫尾：4.5 SOP D列统一2位小数 + 4.4 表头纵向合并 + 文档重新嵌入"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
DOC_ID = "TBX2dTCIqoc3zJxHVdXcA1bUnoh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"


def run(cmd, timeout=180):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=True, timeout=timeout)
    try:
        return json.loads(r.stdout)
    except:
        return None


def read_range(sid, rng):
    resp = run([LARK_CLI, "sheets", "+read", "--spreadsheet-token",
                SPREADSHEET_TOKEN, "--range", f"{sid}!{rng}"])
    if resp and resp.get("ok") and resp.get("data"):
        return resp["data"].get("valueRange", {}).get("values", [])
    return []


def write_range(sid, rng, values):
    vj = json.dumps(values, ensure_ascii=False)
    run([LARK_CLI, "sheets", "+write", "--spreadsheet-token",
         SPREADSHEET_TOKEN, "--range", f"{sid}!{rng}", "--values", vj])
    time.sleep(0.1)


def merge_cells(sid, rng):
    run([LARK_CLI, "sheets", "+merge-cells", "--spreadsheet-token",
         SPREADSHEET_TOKEN, "--range", f"{sid}!{rng}", "--merge-type", "MERGE_ALL"])
    time.sleep(0.15)


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


# === Step 1: 4.5 SOP 把 D-I 列全部数字统一改成 2 位小数 ===
print("[1] 4.5 SOP D-I 列统一为 2 位小数...")
sid = "2Ti90Y"
# 读 D3:I100
vals = read_range(sid, "D3:I100")
print(f"  读到 {len(vals)} 行")
new_vals = []
for row in vals:
    new_row = []
    for cell in row:
        if cell is None or cell == "":
            new_row.append("")
            continue
        try:
            num = float(str(cell).replace("%", "").replace(",", ""))
            new_row.append(f"{num:.2f}")
        except:
            new_row.append(str(cell))
    new_vals.append(new_row)

# 批量写回（每批 10 行）
for i in range(0, len(new_vals), 10):
    batch = new_vals[i:i+10]
    write_range(sid, f"D{3+i}:I{3+i+len(batch)-1}", batch)
print(f"  写回 {len(new_vals)} 行")


# === Step 2: 4.4 表头纵向合并（row1 row2 相同值的列）===
print("\n[2] 4.4 表头纵向合并...")
sid = "2SOCo8"
header = read_range(sid, "A1:CV2")
if header and len(header) >= 2:
    r1 = header[0]
    r2 = header[1] if len(header) > 1 else []
    n = max(len(r1), len(r2))
    merge_count = 0
    ci = 0
    while ci < n:
        v1 = str(r1[ci]).strip() if ci < len(r1) and r1[ci] else ""
        v2 = str(r2[ci]).strip() if ci < len(r2) and r2[ci] else ""
        if v1 and v1 == v2:
            cl = col_letter(ci + 1)
            merge_cells(sid, f"{cl}1:{cl}2")
            merge_count += 1
        ci += 1
    print(f"  纵向合并 {merge_count} 列")


# === Step 3: 文档 overwrite ===
print("\n[3] 重新嵌入文档...")
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
        print(f"  output: {r.stdout[:300]}")
else:
    print("  XML 文件不存在")

print("\n[DONE]")
print(f"文档：https://hcnig43mb8gp.feishu.cn/docx/{DOC_ID}")
print(f"表格：https://hcnig43mb8gp.feishu.cn/sheets/{SPREADSHEET_TOKEN}")
