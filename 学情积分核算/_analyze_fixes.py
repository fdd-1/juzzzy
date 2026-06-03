"""分析需要修改的列位置"""
import sys, io, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

d = json.loads(Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json").read_text(encoding="utf-8"))
m = {s["title"]: s["rows"] for s in d["sheets"]}

# 4.1 服务指标：找"执行率加和"列 和 "首课语义点执行" 合并区间
print("=== 4.1 服务指标 ===")
r1 = m["4.1 服务指标"][0]
r2 = m["4.1 服务指标"][1]
print("Row1:", [(i+1, str(c)[:20]) for i, c in enumerate(r1) if str(c).strip()])
print("Row2:", [(i+1, str(c)[:20]) for i, c in enumerate(r2) if str(c).strip()])
# 找"执行率加和"
for i, c in enumerate(r2):
    if "执行率加和" in str(c):
        print(f"  执行率加和 at col {i+1}")
# 找"首课语义点执行"在 row1 的位置
for i, c in enumerate(r1):
    if "首课语义点" in str(c):
        print(f"  首课语义点执行 at row1 col {i+1}")

print("\n=== 4.3 群发消息 ===")
r1 = m["4.3 群发消息"][0]
r2 = m["4.3 群发消息"][1]
print("Row1:", [(i+1, str(c)[:25]) for i, c in enumerate(r1) if str(c).strip()])

print("\n=== 4.4 停课监控 ===")
r1 = m["4.4 停课监控"][0]
r2 = m["4.4 停课监控"][1]
r3 = m["4.4 停课监控"][2]
# 找"微信关键词覆盖率"和"1V1推报率"
for i, c in enumerate(r3):
    v = str(c).strip()
    if "微信关键词覆盖" in v or "1V1推报" in v or "1v1推报" in v:
        print(f"  删除列: col {i+1} = '{v}'")
# 找"唤醒率"
for i, c in enumerate(r3):
    if "唤醒率" in str(c):
        print(f"  唤醒率 at col {i+1}")

print("\n=== 4.5 服务池跟进 ===")
r2 = m["4.5 服务池跟进"][1]
for i, c in enumerate(r2):
    v = str(c).strip()
    if "外呼跟进率" in v or "综合有效跟进率" in v:
        print(f"  保留色阶: col {i+1} = '{v}'")

print("\n=== 4.6 系统外呼监控 ===")
r2 = m["4.6 系统外呼监控"][1]
print("Row2 前5列:", [(i+1, str(c)[:15]) for i, c in enumerate(r2[:5])])

print("\n=== 4.6 企微回复比 ===")
r2 = m["4.6 企微回复比"][1]
print("Row2 前5列:", [(i+1, str(c)[:15]) for i, c in enumerate(r2[:5])])
