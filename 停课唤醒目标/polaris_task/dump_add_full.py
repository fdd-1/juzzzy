"""dump add 请求的完整请求体（不截断）"""
import json, sys, io
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

HAR = Path(__file__).parent / "demo_network.har"
har = json.loads(HAR.read_text(encoding="utf-8"))
entries = har["log"]["entries"]

for e in entries:
    url = e["request"]["url"]
    if "taskTemplate/add" not in url:
        continue
    pd = e["request"].get("postData") or {}
    text = pd.get("text") or ""
    print(f"REQ length: {len(text)}")
    print(f"REQ full body:")
    print(text)
    print()
    body = e["response"].get("content", {}).get("text") or ""
    print(f"RES: {body}")
    break
