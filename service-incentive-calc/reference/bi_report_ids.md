# BI 报表 ID 与下载规则

> 4 张服务绩效报表的 SmartBI report.id 与日期筛选键。`service_incentive_*.json` 必须与本文一致。

## 4 张报表

| 任务 key | 报表名称 | report.id | 日期筛选键 |
|----------|---------|-----------|-----------|
| service_metrics | 海外思维学管服务指标统计表 | `I2c9280870189f6f1f6f10d05018a64543b1960d0` | 开始日期 / 结束时间 / LP做工开始时间 / LP做工结束时间 |
| service_sop | 海外思维服务SOP执行情况 | `I2c928087019364b764b704c8019375e98bea20d9` | 做工开始时间 / 做工结束时间 / 日期 |
| first_call_monitor | **益智海外新生首通监控** | `I2c928087018722bf22bf7d4d0187ff868fab30da` | 首次分配开始时间 / 首次分配结束时间 |
| tingke_monitor | 思维停课学员执行监控 | `I2c928087018b5acd5acd0fdc018c24e77b3c3525` | 开始日期 / 结束日期 |

## ⚠️ 易错报表识别

- **首通报表别下错**：同目录下还有一张「M0管理_益智海外新生首通监控」（ID `I2c92808701896211621123220189ba98c0360e32`），那张表用错会导致首通指标提取失败。**必须用不带 M0 前缀的「益智海外新生首通监控」**。

## 下载工具

- 工具：`smartbi_cli`（HTTP 直连），路径 `C:\Users\fengjianyi\Desktop\周报自动化\smartbi_cli\smartbi_cli.py`
- ❌ 禁止用 `bi_skill` 的 playwright 方式（monitor 类报表会超时、控制台中文易报 GBK 编码错）
- 凭据：环境变量 `SMARTBI_USERNAME` / `SMARTBI_PASSWORD`，从项目根目录 `.env` 加载（参考 `.env.example`）

## 下载示例

```powershell
# 1. 加载环境变量（或手动 set）
# .env 在项目根目录，包含 SMARTBI_USERNAME / SMARTBI_PASSWORD

# 2. 复用模板配置改日期
copy service_incentive_0605.json service_incentive_<日期>.json
# 编辑 overrides 里的日期值

# 3. 下载
python "C:\Users\fengjianyi\Desktop\周报自动化\smartbi_cli\smartbi_cli.py" run `
    --config service_incentive_<日期>.json `
    --task <service_metrics|service_sop|first_call_monitor|tingke_monitor> `
    --overwrite
```

下载默认落到 `周报自动化/data/<月份>/`，需手动 Copy 到 skill 的 `data/<月份>/` 后用 `--skip-download` 计算。
