import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
d = json.load(open(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json", encoding="utf-8"))
m = {s["title"]: s["rows"] for s in d["sheets"]}

failed = [
    ("4.4 停课监控", 104), ("4.4 停课监控", 105), ("4.4 停课监控", 108), ("4.4 停课监控", 114),
    ("4.5 服务池SOP", 102),
    ("4.6 系统外呼监控", 106), ("4.6 系统外呼监控", 107), ("4.6 系统外呼监控", 108),
]

for sheet, rn in failed:
    row = m[sheet][rn - 1]
    # 列 1-20
    sub = row[:20]
    print(f"=== {sheet} row {rn} (col 1-20) ===")
    for i, v in enumerate(sub, 1):
        if v == "" or v is None:
            continue
        sv = str(v)
        # 显示有问题的字符
        weird = [(j, c, ord(c)) for j, c in enumerate(sv) if ord(c) < 32 and c not in "\t"]
        marker = f"  weird={weird[:3]}" if weird else ""
        print(f"  col{i}: type={type(v).__name__} len={len(sv)} repr={sv[:80]!r}{marker}")
    print()
