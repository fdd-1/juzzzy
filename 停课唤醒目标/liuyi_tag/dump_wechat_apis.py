"""dump 企微标签相关接口的请求/响应"""
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
    if "corporate-wechat-backend" not in url:
        continue
    idx += 1
    print(f"========== #{idx} {e['request']['method']} ==========")
    print(f"URL: {url}")
    print(f"Status: {e['response']['status']}")
    pd = e["request"].get("postData") or {}
    if pd.get("text"):
        print(f"REQ ({pd.get('mimeType')}):")
        print(f"  {pd['text'][:1500]}")
    body = e["response"].get("content", {}).get("text") or ""
    if body:
        print(f"RES:")
        print(f"  {body[:1500]}")
    print()
