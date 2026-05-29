"""扫描 demo_network.har，列出所有可能是接口的请求 + 关键请求体/响应体片段。"""
import json, sys, io
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

HAR = Path(__file__).parent / "demo_network.har"
har = json.loads(HAR.read_text(encoding="utf-8"))
entries = har["log"]["entries"]

skip_ext = (
    ".js", ".css", ".png", ".jpg", ".jpeg", ".svg", ".woff", ".woff2",
    ".gif", ".ico", ".ttf", ".map", ".html",
)

def is_static(url: str, mime: str) -> bool:
    path = url.split("?")[0].lower()
    if any(path.endswith(e) for e in skip_ext):
        return True
    if any(t in mime for t in ("image/", "font/", "javascript", "text/css", "text/html")):
        return True
    return False

api = []
for e in entries:
    url = e["request"]["url"]
    mime = e["response"].get("content", {}).get("mimeType", "")
    if is_static(url, mime):
        continue
    api.append(e)

print(f"total={len(entries)} api-like={len(api)}")
print()

# 简单分组：按 path
from urllib.parse import urlparse
groups = {}
for e in api:
    pr = urlparse(e["request"]["url"])
    key = f"{e['request']['method']} {pr.netloc}{pr.path}"
    groups.setdefault(key, []).append(e)

for key, items in sorted(groups.items()):
    e = items[0]
    print(f"== {key}  (x{len(items)}, status={e['response']['status']})")
    # 请求体
    pd = e["request"].get("postData")
    if pd and pd.get("text"):
        snippet = pd["text"][:300].replace("\n", " ")
        print(f"   REQ: {snippet}")
    # 返回体
    body = e["response"].get("content", {}).get("text") or ""
    if body:
        snippet = body[:300].replace("\n", " ")
        print(f"   RES: {snippet}")
    print()
