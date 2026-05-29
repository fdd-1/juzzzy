"""详细 dump 关键接口的请求体和响应体"""
import json, sys, io
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

HAR = Path(__file__).parent / "demo_network.har"
har = json.loads(HAR.read_text(encoding="utf-8"))
entries = har["log"]["entries"]

KEY_PATHS = [
    "userTag/addwithfile",
    "userGroup/add",
    "userGroup/checkData",
    "tagGroup/create",
    "userTag/query",
    "userGroup/query",
    "tagType/getMultiLevel",
    "bizSys/list",
    "bizchannel/list",
]

for e in entries:
    url = e["request"]["url"]
    if not any(k in url for k in KEY_PATHS):
        continue
    print("=" * 80)
    print(f"{e['request']['method']} {url}")
    print(f"  status: {e['response']['status']}")
    pd = e["request"].get("postData") or {}
    if pd.get("text"):
        print(f"  REQ ({pd.get('mimeType')}):")
        print(f"  {pd['text'][:1500]}")
    body = e["response"].get("content", {}).get("text") or ""
    if body:
        print(f"  RES:")
        print(f"  {body[:1500]}")
    print()
