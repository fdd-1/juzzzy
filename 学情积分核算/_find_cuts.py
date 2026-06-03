"""精确定位每个 sheet 的口径分界行。"""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

d = json.loads(Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json").read_text(encoding="utf-8"))


def find_caliber_start(rows):
    """找口径说明起始行的 index（0-based）；找不到返回 len(rows)"""
    keywords_strict = ["口径说明", "注意：", "说明："]
    for i, r in enumerate(rows):
        first_cell = str(r[0]) if r else ""
        for kw in keywords_strict:
            if kw in first_cell:
                return i
    # 启发式：行首是「数字、」或「数字)」+冒号 且 第二列起几乎为空 → 视为口径
    for i, r in enumerate(rows):
        first = str(r[0]) if r else ""
        # 如 "1、xxx：" / "1）xxx：" / "2、外呼有效跟进率：..."
        import re as _re
        if _re.match(r"^\d+[、）)\.]\s*[^,]+[:：]", first):
            other = [str(c) for c in r[1:] if str(c).strip()]
            if len(other) <= 2:  # 后续基本都空
                return i
    return len(rows)


for s in d["sheets"]:
    rows = s["rows"]
    cut = find_caliber_start(rows)
    print(f"[{s['title']}] total={len(rows)}, caliber_starts_at_row={cut + 1} (data 1..{cut})")
    if cut < len(rows):
        # 显示口径起始 1-2 行
        print(f"  截断处: R{cut + 1}: {str(rows[cut][0])[:80]}")
