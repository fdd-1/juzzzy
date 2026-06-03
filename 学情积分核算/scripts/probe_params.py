#!/usr/bin/env python3
"""探测BI报表的实际参数"""

import asyncio
import sys
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from smartbi_browser_export import export_simple_report

async def probe():
    result = await export_simple_report(
        username="76218",
        password="123456",
        report_id="I2c928087019b236723675f9c019b353f6027505b",
        output_path=Path("/tmp/probe.xlsx"),
        filters=[],  # 不设置任何过滤器，只探测参数
        headless=False,  # 显示浏览器窗口便于调试
    )
    print("\n\n=== 探测结果 ===")
    print(f"参数列表:")
    for p in result.get('params', []):
        print(f"  - alias: {p.get('alias')}, name: {p.get('name')}, id: {p.get('id')}")

if __name__ == "__main__":
    asyncio.run(probe())
