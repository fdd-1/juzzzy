# SmartBI Data CLI 集成说明

本项目已集成 smartbi-data-cli 方法，用于从 BI 系统下载报表。

## 架构说明

### 原理

smartbi-data-cli 提供两种报表导出方式：

1. **SPREADSHEET_REPORT** (电子表格报表)
   - 使用 HTTP/RMI 接口 (`ssreportServlet`)
   - 无需浏览器，速度快
   - 适用于传统的电子表格类报表

2. **SIMPLE_REPORT** (简单报表) ← **本项目使用**
   - 使用 Playwright 浏览器自动化
   - 通过 QueryView 设置筛选条件
   - 调用 `ExportServlet` 导出 Excel
   - 适用于动态查询类报表

### 本项目的两张报表

| 报表名称 | 类型 | Report ID |
|---------|------|-----------|
| 海外思维续费规划表_新版_26年启用 | SIMPLE_REPORT | `I2c928087019b236723675f9c019b353f6027505b` |
| 海外思维学员上课明细 | SIMPLE_REPORT | `I2c9280870198767976798e4f0198889e7cc27654` |

## 文件结构

```
学情积分核算/
├── configs/
│   └── smartbi_simple_report_tasks.json    # 报表任务配置
├── scripts/
│   ├── smartbi_browser_export.py           # 核心浏览器导出模块
│   ├── fetch_reports_smartbi.py            # 新的取数脚本
│   └── fetch_reports.py                    # 已更新为调用新方法
└── 01_bi_exports/                          # 报表输出目录
```

## 使用方法

### 方式1：使用原有脚本（推荐）

```bash
# 原有脚本已更新为使用 smartbi-data-cli 方法
python scripts/fetch_reports.py --start 2026-05-01 --end 2026-05-15

# 调试模式（显示浏览器窗口）
python scripts/fetch_reports.py --start 2026-05-01 --end 2026-05-15 --headful
```

### 方式2：直接使用新脚本

```bash
# 下载报表
python scripts/fetch_reports_smartbi.py --start 2026-05-01 --end 2026-05-15

# 输出 JSON 结果
python scripts/fetch_reports_smartbi.py --start 2026-05-01 --end 2026-05-15 --json

# 调试模式
python scripts/fetch_reports_smartbi.py --start 2026-05-01 --end 2026-05-15 --headful
```

## 配置说明

### 报表任务配置 (`configs/smartbi_simple_report_tasks.json`)

```json
{
  "version": 1,
  "base_url": "https://bi.61info.cn/smartbi/vision",
  "tasks": {
    "xufei_guihua_new": {
      "enabled": true,
      "description": "海外思维续费规划表_新版_26年启用",
      "report": {
        "id": "I2c928087019b236723675f9c019b353f6027505b",
        "type": "SIMPLE_REPORT"
      },
      "filters": {
        "date_mapping": {
          "当前课包签单年月开始": "start_date",
          "当前课包签单年月结束": "end_date",
          "当前课包签单时间开始": "start_date",
          "当前课包签单时间结束": "end_date"
        }
      },
      "output": {
        "filename": "海外思维续费规划表_新版_26年启用.xlsx"
      },
      "max_rows": 50000
    }
  }
}
```

### 筛选条件映射

`date_mapping` 定义了如何将命令行参数映射到报表筛选字段：

- `start_date` → 所有映射为 `start_date` 的字段都会被设置为 `--start` 参数值
- `end_date` → 所有映射为 `end_date` 的字段都会被设置为 `--end` 参数值

## 环境变量

```bash
# BI 账号密码（可选，默认值已内置）
export SMARTBI_USERNAME="76218"
export SMARTBI_PASSWORD="123456"
```

## 核心模块说明

### `smartbi_browser_export.py`

核心浏览器导出模块，基于 smartbi-data-cli-internal-20260526 实现。

**主要功能：**
- 使用 Playwright 打开 SmartBI 报表页面
- 通过 RMI 接口登录
- 通过 QueryView JavaScript API 设置筛选条件
- 调用 ExportServlet 导出 Excel
- 支持行数限制保护（默认 50000 行）

**关键 API：**
```python
async def export_simple_report(
    *,
    username: str,
    password: str,
    report_id: str,
    output_path: Path,
    max_rows: int = 50000,
    browser_channel: str = "chrome",
    headless: bool = True,
    filters: list[list[str]] | None = None,  # [[alias, value, displayValue], ...]
    base_url: str = BASE_URL,
) -> dict[str, Any]:
```

**筛选条件格式：**
```python
filters = [
    ["开始日期", "2026-05-01", "2026-05-01"],
    ["结束日期", "2026-05-15", "2026-05-15"],
]
```

### `fetch_reports_smartbi.py`

报表下载脚本，读取 `configs/smartbi_simple_report_tasks.json` 配置，
根据 `date_mapping` 自动构建筛选条件，调用 `smartbi_browser_export` 下载报表。

**主要函数：**
- `load_config()` - 加载任务配置
- `build_filters_for_task()` - 根据 date_mapping 构建筛选条件
- `fetch_one_report()` - 下载单张报表
- `fetch_all()` - 下载所有启用的报表

## 浏览器端实现原理

导出过程在浏览器端执行以下 JavaScript 代码：

1. **获取 QueryView 对象**
   ```javascript
   const adapter = window.getReportAdapter();
   const query = adapter.queryViewCommand.query;
   ```

2. **设置筛选条件**
   ```javascript
   const paramId = paramIdByAlias(alias);
   query.paramPanelObj.setParamValue(paramId, value, displayValue, null, null, true);
   query.setParamValue(paramId, value, displayValue, true);
   ```

3. **刷新数据**
   ```javascript
   await rmi('CompositeService', 'refreshDataWithDefaultEx', [query.clientId, false, false]);
   ```

4. **获取行数**
   ```javascript
   const rowCount = await rmi('ClientReportService', 'getTotalRowsCountWithFuture', [query.clientId, 0]);
   ```

5. **导出 Excel**
   ```javascript
   const response = await fetch('ExportServlet', {
     method: 'POST',
     body: new URLSearchParams({
       type: 'EXCEL2007',
       clientId: query.clientId,
       maxRow: String(maxRows),
       contentType: 'gridOnly',
       // ...
     })
   });
   ```

## 与原 bi_skill 的对比

| 特性 | smartbi-data-cli | 原 bi_skill |
|------|------------------|-------------|
| 配置方式 | JSON 配置文件 | 命令行参数 + profile 查找 |
| 报表定位 | 直接使用 report_id | 通过 profile 名称查找 |
| 筛选设置 | date_mapping 自动映射 | --extra-dates 手动指定 |
| 依赖 | Playwright | Playwright + profile.py |
| 可维护性 | 配置与代码分离 | 命令行参数较长 |
| 扩展性 | 易于添加新报表 | 需修改脚本 |

## 依赖安装

```bash
# 安装 Playwright
pip install playwright

# 安装 Chrome 浏览器驱动
playwright install chrome
```

## 故障排查

### 1. 登录失败

**现象：** `SmartBI login failed`

**解决：**
- 检查环境变量 `SMARTBI_USERNAME` 和 `SMARTBI_PASSWORD`
- 确认账号密码正确

### 2. 筛选字段未找到

**现象：** `SmartBI parameter not found: xxx`

**解决：**
- 使用 `--headful` 模式查看浏览器中的实际字段名
- 更新 `configs/smartbi_simple_report_tasks.json` 中的 `date_mapping`

### 3. 行数超限

**现象：** `rowCount xxx exceeds maxRows xxx`

**解决：**
- 调整 `max_rows` 配置
- 缩小日期范围

### 4. 导出超时

**现象：** 浏览器等待超时

**解决：**
- 检查网络连接
- 增加 `timeout` 参数（需修改代码）
- 使用 `--headful` 模式观察浏览器行为

## 参考资料

- smartbi-data-cli-internal-20260526 源码
- SmartBI QueryView API 文档
- Playwright 文档: https://playwright.dev/python/

## 更新日志

### 2026-05-28
- ✨ 集成 smartbi-data-cli 方法
- ✨ 创建 `smartbi_browser_export.py` 核心模块
- ✨ 创建 `configs/smartbi_simple_report_tasks.json` 配置
- ✨ 创建 `fetch_reports_smartbi.py` 新取数脚本
- 🔧 更新 `fetch_reports.py` 使用新方法
