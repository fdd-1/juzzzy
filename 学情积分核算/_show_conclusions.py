import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
d = json.load(open(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json", encoding="utf-8"))
for sid, sec in d["conclusions"].items():
    print("=" * 60)
    print(f"[{sid}] {sec['title']}")
    print(sec["conclusion"])
    print()
