# SmartBI Data CLI 学习总结

## 已完成的工作

我已经学习了 `smartbi-data-cli-internal-20260526` 的实现方法，并成功集成到你的「学情积分核算」项目中。

## 核心文件

### 1. `scripts/smartbi_browser_export.py`
核心浏览器导出模块，基于 smartbi-data-cli 实现。

**功能：**
- 使用 Playwright 打开 SmartBI SIMPLE_REPORT
- 通过 RMI 接口登录
- 通过 QueryView JavaScript API 设置筛选条件
- 调用 ExportServlet 导出 Excel

**关键代码：**
```python
async def export_simple_report(
    username: str,
    password: str,
    report_id: str,
    output_path: Path,
    max_rows: int = 50000,
    filters: list[list[str]] | None = None,  # [[alias, value, displayValue], ...]
) -> dict[str, Any]:
```

### 2. `configs/smartbi_simple_report_tasks.json`
报表任务配置文件。

**配置结构：**
```json
{
  "tasks": {
    "xufei_guihua_new": {
      "report": {"id": "I2c928087019b236723675f9c019b353f6027505b"},
      "filters": {
        "date_mapping": {
          "当前课包签单年月开始": "start_date",
          "当前课包签单年月结束": "end_date"
        }
      },
      "output": {"filename": "海外思维续费规划表_新版_26年启用.xlsx"},
      "max_rows": 50000
    }
  }
}
```

### 3. `scripts/fetch_reports_smartbi.py`
新的报表下载脚本。

**功能：**
- 读取 JSON 配置
- 根据 date_mapping 自动构建筛选条件
- 并发下载多张报表
- 输出详细日志和结果

### 4. `scripts/fetch_reports.py` (已更新)
原有脚本已更新为调用新方法。

## 使用方法

### 基本用法
```bash
# 下载报表（使用原有脚本，已更新为新方法）
python scripts/fetch_reports.py --start 2026-05-01 --end 2026-05-15

# 或直接使用新脚本
python scripts/fetch_reports_smartbi.py --start 2026-05-01 --end 2026-05-15

# 调试模式（显示浏览器窗口）
python scripts/fetch_reports.py --start 2026-05-01 --end 2026-05-15 --headful
```

### 环境变量
```bash
export SMARTBI_USERNAME="76218"
export SMARTBI_PASSWORD="123456"
```

## 技术原理

### SmartBI 报表类型

1. **SPREADSHEET_REPORT** (电子表格报表)
   - 使用 HTTP/RMI 接口 (`ssreportServlet`)
   - 无需浏览器，速度快

2. **SIMPLE_REPORT** (简单报表) ← **你的项目使用这种**
   - 使用 Playwright 浏览器自动化
   - 通过 QueryView 设置筛选条件
   - 调用 ExportServlet 导出

### 浏览器端实现流程

1. **登录**
   ```javascript
   await fetch('RMIServlet', {
     method: 'POST',
     body: new URLSearchParams({
       className: 'UserService',
       methodName: 'clickLogin',
       params: JSON.stringify([username, password])
     })
   });
   ```

2. **打开报表**
   ```
   /openresource.jsp?isBrowse=true&resid={report_id}
   ```

3. **设置筛选条件**
   ```javascript
   const query = window.getReportAdapter().queryViewCommand.query;
   const paramId = paramIdByAlias(alias);
   query.paramPanelObj.setParamValue(paramId, value, displayValue, null, null, true);
   ```

4. **刷新数据**
   ```javascript
   await rmi('CompositeService', 'refreshDataWithDefaultEx', [query.clientId, false, false]);
   ```

5. **获取行数**
   ```javascript
   const rowCount = await rmi('ClientReportService', 'getTotalRowsCountWithFuture', [query.clientId, 0]);
   ```

6. **导出 Excel**
   ```javascript
   await fetch('ExportServlet', {
     method: 'POST',
     body: new URLSearchParams({
       type: 'EXCEL2007',
       clientId: query.clientId,
       maxRow: String(maxRows),
       contentType: 'gridOnly'
     })
   });
   ```

## 你的两张报表

| 报表名称 | 类型 | Report ID | 筛选字段 |
|---------|------|-----------|---------|
| 海外思维续费规划表_新版_26年启用 | SIMPLE_REPORT | `I2c928087019b236723675f9c019b353f6027505b` | 当前课包签单年月开始/结束<br>当前课包签单时间开始/结束 |
| 海外思维学员上课明细 | SIMPLE_REPORT | `I2c9280870198767976798e4f0198889e7cc27654` | 开始日期<br>结束日期 |

## 与原 bi_skill 的对比

| 特性 | smartbi-data-cli | 原 bi_skill |
|------|------------------|-------------|
| 配置方式 | JSON 配置文件 | 命令行参数 + profile 查找 |
| 报表定位 | 直接使用 report_id | 通过 profile 名称查找 |
| 筛选设置 | date_mapping 自动映射 | --extra-dates 手动指定 |
| 依赖 | Playwright | Playwright + profile.py |
| 可维护性 | 配置与代码分离 | 命令行参数较长 |
| 扩展性 | 易于添加新报表 | 需修改脚本 |

## 优势

1. **配置与代码分离** - 报表配置在 JSON 文件中，易于维护
2. **自动筛选映射** - date_mapping 自动将日期参数映射到多个筛选字段
3. **直接使用 report_id** - 无需依赖 profile 查找
4. **标准化接口** - 基于 smartbi-data-cli 的标准实现
5. **易于扩展** - 添加新报表只需修改 JSON 配置

## 依赖

```bash
pip install playwright
playwright install chrome
```

## 文档

详细文档请查看：
- `docs/smartbi_data_cli_integration.md` - 完整集成文档
- `configs/smartbi_simple_report_tasks.json` - 报表配置示例
- `scripts/smartbi_browser_export.py` - 核心模块源码

## 测试

```bash
# 测试语法
python -c "import ast; ast.parse(open('scripts/smartbi_browser_export.py', encoding='utf-8').read())"
python -c "import ast; ast.parse(open('scripts/fetch_reports_smartbi.py', encoding='utf-8').read())"

# 测试下载（需要 Playwright）
python scripts/fetch_reports.py --start 2026-05-01 --end 2026-05-15 --headful
```

## 下一步

你现在可以：
1. 运行 `python scripts/fetch_reports.py --start 2026-05-01 --end 2026-05-15` 测试新方法
2. 查看 `docs/smartbi_data_cli_integration.md` 了解详细原理
3. 根据需要在 `configs/smartbi_simple_report_tasks.json` 中添加更多报表
4. 原有的 `xueqing_credit_skill.py` 无需修改，因为它调用的 `scripts/fetch_reports.py` 已更新

## 总结

我已经成功将 smartbi-data-cli 的 SIMPLE_REPORT 导出方法集成到你的项目中。新方法使用 Playwright 浏览器自动化，通过 QueryView JavaScript API 设置筛选条件，然后调用 ExportServlet 导出 Excel。配置与代码分离，易于维护和扩展。
