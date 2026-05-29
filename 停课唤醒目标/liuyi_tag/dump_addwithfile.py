"""详细 dump userTag/addwithfile 的 multipart 字段"""
import json, sys, io
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

HAR = Path(__file__).parent / "demo_network.har"
har = json.loads(HAR.read_text(encoding="utf-8"))
entries = har["log"]["entries"]

idx = 0
for e in entries:
    url = e["request"]["url"]
    if "userTag/addwithfile" not in url:
        continue
    idx += 1
    print(f"========== 第 {idx} 次 addwithfile ==========")
    print(f"URL: {url}")
    print(f"Method: {e['request']['method']}")
    print(f"Headers:")
    for h in e["request"]["headers"]:
        if h["name"].lower() in ("content-type", "content-length"):
            print(f"  {h['name']}: {h['value']}")
    pd = e["request"].get("postData") or {}
    print(f"PostData mimeType: {pd.get('mimeType')}")
    print(f"PostData params: {pd.get('params')}")
    text = pd.get("text") or ""
    print(f"PostData text length: {len(text)}")
    # 对 multipart 文本直接打印前 3000 字（含 boundary 段）
    print(f"PostData text head:\n{text[:3000]}")
    print(f"\n... 截断 ...\n")
    print(f"PostData text tail:\n{text[-1500:]}")
    print()
