#!/usr/bin/env python3
"""海外LP群发违规播报 - 以组为维度推送钉钉群."""
import argparse
import sys
from datetime import date
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

try:
    import pandas as pd
    import requests
except ImportError:
    print("缺少依赖，请先安装: pip install pandas openpyxl requests")
    sys.exit(1)


def load_webhook_url():
    env_file = Path.home() / ".claude" / "secrets" / "intro_monitor.env"
    if not env_file.exists():
        return None
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line.startswith("DINGTALK_WEBHOOK_URL="):
            return line.split("=", 1)[1].strip()
    return None


def broadcast(stat_excel):
    print(f"[1] 读取违规统计: {stat_excel}")
    df = pd.read_excel(stat_excel, sheet_name="违规统计")
    if df.empty:
        print("    ⚠ 无违规LP，跳过播报")
        return
    print(f"    ✓ 共 {len(df)} 位LP")

    today = date.today().strftime("%Y-%m-%d")
    lines = []
    lines.append(f"## 🚨 海外LP群发违规播报 {today}")
    lines.append("> 规则：续费场景 + 与报备文本相似度<90%\n")

    total = len(df)
    total_count = int(df["续费违规条数"].sum())

    for group, gdf in df.groupby("小组"):
        gdf = gdf.sort_values("续费违规条数", ascending=False)
        worst_idx = gdf["续费违规条数"].idxmax()

        lines.append(f"### {group}")
        lines.append("| LP姓名 | 续费违规条数 |")
        lines.append("|--------|------------|")
        for idx, row in gdf.iterrows():
            lp = row["LP姓名"]
            cnt = int(row["续费违规条数"])
            if idx == worst_idx:
                lines.append(f'| <font color="red">**{lp}**</font> | <font color="red">{cnt}</font> |')
            else:
                lines.append(f"| **{lp}** | {cnt} |")
        lines.append("")

    lines.append("---")
    lines.append(f"**汇总**：{total} 位LP，共 {total_count} 条续费违规（已剔除报备命中 + 转介绍场景）")

    md = "\n".join(lines)
    print("\n" + "=" * 60)
    print(md)
    print("=" * 60 + "\n")

    print("[2] 播报到钉钉群...")
    url = load_webhook_url()
    if not url:
        print("    ⚠ 未找到 DINGTALK_WEBHOOK_URL，跳过播报")
        return

    payload = {"msgtype": "markdown", "markdown": {"title": "海外LP群发违规播报", "text": md}}
    try:
        resp = requests.post(url, json=payload, timeout=10)
        resp.raise_for_status()
        result = resp.json()
        if result.get("errcode") == 0:
            print("    ✓ 播报成功")
        else:
            print(f"    ✗ 播报失败: {result}")
    except Exception as e:
        print(f"    ✗ 播报异常: {e}")


def main():
    p = argparse.ArgumentParser(description="海外LP群发违规播报")
    p.add_argument("stat_excel", help="违规统计 Excel（check.py 的输出）")
    args = p.parse_args()
    if not Path(args.stat_excel).exists():
        print(f"错误: 文件不存在 {args.stat_excel}")
        sys.exit(1)
    broadcast(args.stat_excel)
    print("\n播报完成")


if __name__ == "__main__":
    main()
