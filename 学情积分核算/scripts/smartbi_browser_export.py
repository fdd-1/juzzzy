#!/usr/bin/env python3
"""Browser-backed SmartBI SIMPLE_REPORT export helper.

Based on smartbi-data-cli-internal-20260526. Uses Playwright to open a
SIMPLE_REPORT, set filters via QueryView, and export to Excel.

Credentials: SMARTBI_USERNAME / SMARTBI_PASSWORD environment variables.
"""

from __future__ import annotations

import asyncio
import json
import os
from pathlib import Path
from typing import Any

BASE_URL = "https://bi.61info.cn/smartbi/vision"


class SmartbiBrowserExportError(RuntimeError):
    pass


def normalize_filters(
    filters: list[tuple[str, str, str]] | list[list[str]] | None,
) -> list[list[str]]:
    normalized: list[list[str]] = []
    for item in filters or []:
        if len(item) != 3:
            raise SmartbiBrowserExportError(f"Filter needs [alias, value, displayValue]: {item}")
        normalized.append([str(item[0]), str(item[1]), str(item[2])])
    return normalized


async def export_simple_report(
    *,
    username: str,
    password: str,
    report_id: str,
    output_path: Path,
    max_rows: int = 50000,
    browser_channel: str = "chrome",
    headless: bool = True,
    filters: list[list[str]] | None = None,
    base_url: str = BASE_URL,
) -> dict[str, Any]:
    """Export a SIMPLE_REPORT to Excel via browser."""
    try:
        from playwright.async_api import async_playwright
    except ImportError as error:
        raise SmartbiBrowserExportError(
            "Playwright is required: pip install playwright && playwright install chrome"
        ) from error

    effective_filters = normalize_filters(filters)

    async with async_playwright() as pw:
        launch_kwargs: dict[str, Any] = {
            "headless": headless,
            # 增大V8堆内存，避免大数据量（如上课明细9w+行）导出时OOM
            "args": [
                "--js-flags=--max-old-space-size=8192",
                "--disable-dev-shm-usage",
            ],
        }
        if browser_channel:
            launch_kwargs["channel"] = browser_channel
        browser = await pw.chromium.launch(**launch_kwargs)
        context = await browser.new_context()
        page = await context.new_page()
        try:
            result = await _do_export(
                page, username, password, report_id,
                output_path, max_rows, effective_filters, base_url,
            )
            return result
        finally:
            await browser.close()


async def _do_export(
    page,
    username: str,
    password: str,
    report_id: str,
    output_path: Path,
    max_rows: int,
    filters: list[list[str]],
    base_url: str,
) -> dict[str, Any]:
    await page.goto(
        f"{base_url}/index.jsp?time=1778858593300",
        wait_until="domcontentloaded",
        timeout=60000,
    )
    login_payload = await page.evaluate(
        """async ({username, password}) => {
          const body = new URLSearchParams({
            className: 'UserService',
            methodName: 'clickLogin',
            params: JSON.stringify([username, password])
          });
          const r = await fetch('RMIServlet', {
            method: 'POST',
            headers: {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'},
            body
          });
          return await r.json();
        }""",
        {"username": username, "password": password},
    )
    if login_payload.get("result") is not True:
        raise SmartbiBrowserExportError("SmartBI login failed")

    report_url = f"{base_url}/openresource.jsp?isBrowse=true&showLeftTree=default&resid={report_id}"
    await page.goto(report_url, wait_until="domcontentloaded", timeout=120000)
    await page.wait_for_timeout(8000)

    result = await page.evaluate(_EXPORT_JS, {
        "filters": filters,
        "maxRows": max_rows,
        "exportFile": True,
    })

    body_values = result.pop("bytes", None)
    if body_values is None:
        if result.get("exportSkipped"):
            raise SmartbiBrowserExportError(
                f"Export skipped: {result.get('skipReason', 'unknown')} "
                f"(rowCount={result.get('rowCount')}, maxRows={result.get('maxRows')})"
            )
        raise SmartbiBrowserExportError(f"Export returned no bytes: {result}")

    body = bytes(body_values)
    content_type = result.get("contentType", "")
    if "excel" not in content_type.lower() and not body.startswith(b"PK\x03\x04"):
        preview = body[:300].decode("utf-8", "replace")
        raise SmartbiBrowserExportError(
            f"Export did not return Excel: {content_type}: {preview}"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(body)
    result["output"] = str(output_path)
    result["bytes"] = len(body)
    return result


_EXPORT_JS = """async ({filters, maxRows, exportFile}) => {
  async function rmi(className, methodName, params) {
    const body = new URLSearchParams({className, methodName, params: JSON.stringify(params)});
    const response = await fetch('RMIServlet', {
      method: 'POST',
      headers: {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'},
      body
    });
    try {
      const payload = await response.json();
      if (payload.retCode !== 0 && payload.retCode !== '0') {
        throw new Error(JSON.stringify(payload).slice(0, 500));
      }
      return payload.result;
    } catch (e) {
      throw new Error('RMI response parse failed: ' + e.message);
    }
  }
  const adapter = window.getReportAdapter && window.getReportAdapter();
  const query =
    adapter && adapter.queryViewCommand && adapter.queryViewCommand.query ||
    adapter && adapter.processor && adapter.processor.queryViewCommand && adapter.processor.queryViewCommand.query ||
    adapter && adapter.processor && adapter.processor.query;
  if (!query || !query.paramPanelObj || !query.params) {
    throw new Error('SmartBI QueryView is not ready');
  }
  const params = query.params.map((p) => ({
    id: p.id, alias: p.alias, name: p.name, value: p.value, displayValue: p.displayValue
  }));
  console.log('Available params:', JSON.stringify(params));
  const paramIdByAlias = (alias) => {
    // 精确匹配
    let p = query.params.find((c) => c.alias === alias || c.name === alias);
    if (p) return p.id;

    // 模糊匹配：包含关键词
    const keywords = alias.split(/[_\\-\\s]+/).filter(k => k.length > 0);
    p = query.params.find((c) => {
      const fullText = (c.alias + ' ' + c.name).toLowerCase();
      return keywords.every(k => fullText.includes(k.toLowerCase()));
    });
    if (p) return p.id;

    throw new Error('SmartBI parameter not found: ' + alias);
  };
  const applied = [];
  for (const [alias, value, displayValue] of filters) {
    const paramId = paramIdByAlias(alias);
    query.paramPanelObj.setParamValue(paramId, value, displayValue, null, null, true);
    query.setParamValue(paramId, value, displayValue, true);
    applied.push({alias, value, displayValue});
  }
  await rmi('CompositeService', 'refreshDataWithDefaultEx', [query.clientId, false, false]);
  const rowCount = await rmi('ClientReportService', 'getTotalRowsCountWithFuture', [query.clientId, 0]);
  const base = {clientId: query.clientId, params, applied, rowCount, maxRows};
  if (!Number.isInteger(rowCount) || rowCount < 0) {
    throw new Error('SmartBI returned invalid row count: ' + rowCount);
  }
  if (rowCount > maxRows) {
    return {...base, exportSkipped: true, skipReason: 'rowCount ' + rowCount + ' exceeds maxRows ' + maxRows};
  }
  if (!exportFile) {
    return {...base, exportSkipped: true, skipReason: 'probe_only'};
  }
  const form = new URLSearchParams({
    type: 'EXCEL2007', clientId: query.clientId, delimiter: '', maxRow: String(maxRows),
    mode: '', valueType: '', headerHtml: '', tailHtml: '', result: '',
    contentType: 'gridOnly', pageId: '', exportHiddenField: '', needHiddenParamIds: '',
    checkedRows: '', chartAlign: '', codeType: 'UTF-8', isBOM: ''
  });
  const response = await fetch('ExportServlet', {
    method: 'POST',
    headers: {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'},
    body: form
  });
  const contentType = response.headers.get('content-type') || '';
  const contentDisposition = response.headers.get('content-disposition') || '';
  const buffer = await response.arrayBuffer();
  return {...base, contentType, contentDisposition, bytes: Array.from(new Uint8Array(buffer))};
}"""
