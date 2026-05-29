"""详细 dump 北极星 taskTemplate 相关接口"""
import json, sys, io
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

HAR = Path(__file__).parent / "demo_network.har"
har = json.loads(HAR.read_text(encoding="utf-8"))
entries = har["log"]["entries"]

KEY_PATHS = [
    "taskTemplate/add",
    "taskTemplate/16/list",
    "taskTemplate/list",
    "taskTemplate/detail",
    "taskTemplate/update",
    "taskTemplate/edit",
    "taskTemplate/get",
]

for e in entries:
    url = e["request"]["url"]
    if not any(k in url for k in KEY_PATHS):
        continue
    print("=" * 80)
    print(f"{e['request']['method']} {url}")
    print(f"  status: {e['response']['status']}")
    # headers (重要的)
    important_hdrs = ("authorization", "x-token", "token", "content-type", "x-csrf-token")
    print("  important headers:")
    for h in e["request"]["headers"]:
        if h["name"].lower() in important_hdrs:
            v = h["value"]
            print(f"    {h['name']}: {v[:100]}{'...' if len(v) > 100 else ''}")
    pd = e["request"].get("postData") or {}
    if pd.get("text"):
        print(f"  REQ ({pd.get('mimeType')}):")
        print(f"  {pd['text'][:2000]}")
    body = e["response"].get("content", {}).get("text") or ""
    if body:
        print(f"  RES:")
        print(f"  {body[:2000]}")
    print()
