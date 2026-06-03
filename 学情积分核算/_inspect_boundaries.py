"""逐个 sheet 输出前 8 行和后 20 行，方便识别口径区边界。"""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

d = json.loads(Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json").read_text(encoding="utf-8"))

for sheet in d["sheets"]:
    title = sheet["title"]
    rows = sheet["rows"]
    n = len(rows)
    print("=" * 80)
    print(f"[{title}] {n} rows × {len(rows[0]) if rows else 0} cols")
    print("--- 前 8 行 ---")
    for i, r in enumerate(rows[:8], 1):
        cells = [str(c) for c in r if str(c) != ""]
        if not cells:
            preview = "(空行)"
        else:
            preview = " | ".join(cells[:6])
            if len(cells) > 6:
                preview += f" | ...({len(cells)} cols)"
        print(f"  R{i}: {preview[:200]}")
    print("--- 后 20 行 ---")
    for i, r in enumerate(rows[-20:], n - 19):
        cells = [str(c) for c in r if str(c) != ""]
        if not cells:
            preview = "(空行)"
        else:
            preview = " | ".join(cells[:6])
            if len(cells) > 6:
                preview += f" | ...({len(cells)} cols)"
        print(f"  R{i}: {preview[:200]}")
    print()
