# 服务周报自动化 - SmartBI CLI 使用说明

## 一、前提条件

### 1.1 SmartBI 凭据配置

在执行前设置环境变量：

```powershell
# PowerShell
$env:SMARTBI_USERNAME = "你的用户名"
$env:SMARTBI_PASSWORD = "你的密码"
```

```bash
# Bash
export SMARTBI_USERNAME='你的用户名'
export SMARTBI_PASSWORD='你的密码'
```

**注意**：凭据仅在当前会话有效，不会持久化到配置文件。

### 1.2 工具路径

SmartBI CLI 位置（按本机情况，通过环境变量解析；**不要**写死绝对路径）：
- 解析顺序：`$env:SMARTBI_CLI_DIR` → 与项目同级目录下的 `smartbi-data-cli-internal-*`
- 主脚本：`<SMARTBI_CLI_DIR>/scripts/smartbi_cli.py`
- 配置目录：`<SMARTBI_CLI_DIR>/configs/`
- 输出目录：`<SMARTBI_CLI_DIR>/outputs/bi_exports/`

---

## 二、获取报表 ID（首次配置）

服务周报配置文件位于（项目相对路径）：
```
service_weekly_report\configs\service_weekly_smartbi_tasks.json
```

**首次使用需要填写各报表的 `id`**。获取方式：

### 2.1 列出目录下的报表

```bash
# 先指向 SmartBI CLI（一次会话设一次即可）
$env:SMARTBI_CLI_DIR = "<smartbi-data-cli-internal-* 绝对路径>"
Push-Location -LiteralPath $env:SMARTBI_CLI_DIR

python scripts/smartbi_cli.py catalog-list \
  --path "分析报表/海外直播业务线/首通精准识别报表"
```

输出示例：
```
- 益智海外新生首通监控 (I2c928xxx...)
- 其他报表...
```

### 2.2 查看报表详细信息

```bash
python scripts/smartbi_cli.py inspect-report \
  --report-id I2c928xxx... \
  --report-path "分析报表/海外直播业务线/首通精准识别报表/益智海外新生首通监控" \
  --json
```

输出会显示：
- 报表ID
- 可用筛选参数（filters）
- 报表类型

### 2.3 填写配置文件

将获取到的 `id` 填入 `service_weekly_smartbi_tasks.json` 中对应任务的 `report.id` 字段。

**需要填写的 12 个报表**：
1. `service_weekly_4_1_shoutong` - 益智海外新生首通监控
2. `service_weekly_4_1_shouke` - 海外思维学管服务指标统计表（首课）
3. `service_weekly_4_1_shouzhuan` - 海外思维学管服务指标统计表（首专）
4. `service_weekly_4_1_sop` - 海外思维服务SOP执行情况
5. `service_weekly_4_1_lp_arch` - 海外思维LP架构表
6. `service_weekly_4_2` - 思维LP组班意向提交播报
7. `service_weekly_4_3` - 思维海外群发消息汇总数据播报
8. `service_weekly_4_4` - 思维停课学员执行监控
9. `service_weekly_4_5_fuwuyue` - 思维转介绍过程跟进报表_末次渠道
10. `service_weekly_4_5_sop` - 海外思维服务SOP执行情况（服务池）
11. `service_weekly_4_6_waihu` - LP系统外呼监控-分池子
12. `service_weekly_4_6_qiwei` - LP企微回复比监控-分池子

---

## 三、下载报表

### 3.1 单个任务下载（测试）

```bash
Push-Location -LiteralPath $env:SMARTBI_CLI_DIR

# 配置路径以项目根为基准（建议先 $env:WEEKLY_REPORT_ROOT = "<service_weekly_report 绝对路径>"）
$cfg = "$env:WEEKLY_REPORT_ROOT\configs\service_weekly_smartbi_tasks.json"

# Dry-run（不实际下载，仅验证配置）
python scripts/smartbi_cli.py run `
  --config $cfg `
  --task service_weekly_4_1_shoutong `
  --dry-run `
  --json

# 实际下载
python scripts/smartbi_cli.py run `
  --config $cfg `
  --task service_weekly_4_1_shoutong `
  --overwrite `
  --json
```

### 3.2 批量下载所有报表

```bash
Push-Location -LiteralPath $env:SMARTBI_CLI_DIR

python scripts/run_smartbi_batch.py `
  --config "$env:WEEKLY_REPORT_ROOT\configs\service_weekly_smartbi_tasks.json" `
  --task service_weekly_4_1_shoutong \
  --task service_weekly_4_1_shouke \
  --task service_weekly_4_1_shouzhuan \
  --task service_weekly_4_1_sop \
  --task service_weekly_4_1_lp_arch \
  --task service_weekly_4_2 \
  --task service_weekly_4_3 \
  --task service_weekly_4_4 \
  --task service_weekly_4_5_fuwuyue \
  --task service_weekly_4_5_sop \
  --task service_weekly_4_6_waihu \
  --task service_weekly_4_6_qiwei \
  --max-workers 3 \
  --execute \
  --json
```

**参数说明**：
- `--max-workers 3`：并发下载 3 个任务（建议值，避免过载）
- `--execute`：实际执行下载（不加则为 dry-run）
- `--overwrite`：覆盖已存在的文件

---

## 四、输出路径

下载的文件会保存到：

```
<SMARTBI_CLI_DIR>\outputs\bi_exports\service_weekly\{run_date}\
├── 4_1_shoutong\
│   └── 益智海外新生首通监控.xlsx
├── 4_1_shouke\
│   └── 海外思维学管服务指标统计表.xlsx
├── 4_1_shouzhuan\
│   └── 海外思维学管服务指标统计表.xlsx
...
```

**`{run_date}`** 会自动替换为执行日期（如 `2026-06-09`）。

---

## 五、自动化集成

### 5.1 在服务周报自动化脚本中调用

创建 Python 包装脚本 `download_reports_smartbi.py`：

```python
import os
import sys
import subprocess
from pathlib import Path
from datetime import datetime

# 项目根 = 本文件所在仓库的 service_weekly_report 目录
PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
from _paths import resolve_smartbi_cli_dir   # 共享解析逻辑

SMARTBI_CLI_DIR = resolve_smartbi_cli_dir()
CONFIG_PATH = PROJECT_ROOT / "configs" / "service_weekly_smartbi_tasks.json"

TASKS = [
    "service_weekly_4_1_shoutong",
    "service_weekly_4_1_shouke",
    "service_weekly_4_1_shouzhuan",
    "service_weekly_4_1_sop",
    "service_weekly_4_1_lp_arch",
    "service_weekly_4_2",
    "service_weekly_4_3",
    "service_weekly_4_4",
    "service_weekly_4_5_fuwuyue",
    "service_weekly_4_5_sop",
    "service_weekly_4_6_waihu",
    "service_weekly_4_6_qiwei",
]

def download_all_reports():
    """批量下载所有服务周报报表"""
    cmd = [
        sys.executable,
        str(SMARTBI_CLI_DIR / "scripts" / "run_smartbi_batch.py"),
        "--config", str(CONFIG_PATH),
        *[f"--task {t}" for t in TASKS],
        "--max-workers", "3",
        "--execute",
        "--json"
    ]
    
    result = subprocess.run(
        cmd,
        cwd=str(SMARTBI_CLI_DIR),
        capture_output=True,
        text=True,
        timeout=1800  # 30分钟超时
    )
    
    if result.returncode != 0:
        print(f"下载失败: {result.stderr}")
        return None
    
    # 返回输出目录
    run_date = datetime.now().strftime("%Y-%m-%d")
    return SMARTBI_CLI_DIR / "outputs" / "bi_exports" / "service_weekly" / run_date

if __name__ == "__main__":
    output_dir = download_all_reports()
    if output_dir:
        print(f"✓ 报表下载完成: {output_dir}")
    else:
        print("✗ 报表下载失败")
```

---

## 六、日期窗口配置

配置文件中的 `date_window` 会自动计算：

| 值 | 含义 | 示例（今天=2026-06-09） |
|---|---|---|
| `previous_week` | 上周一到上周日 | 2026-06-01 到 2026-06-07 |
| `current_week_snapshot` | 本周一到今天 | 2026-06-08 到 2026-06-09 |
| `previous_month` | 上个完整月 | 2026-05-01 到 2026-05-31 |
| `month_to_last_sunday` | 当月1号到上周日 | 2026-06-01 到 2026-06-07 |

**注意**：`month_to_last_sunday` 是自定义窗口，可能需要在 smartbi_cli.py 中扩展支持。

---

## 七、故障排查

### 7.1 认证失败
```
auth_error: SMARTBI_USERNAME/SMARTBI_PASSWORD is required
```
**解决**：设置环境变量 `SMARTBI_USERNAME` 和 `SMARTBI_PASSWORD`。

### 7.2 报表 ID 不存在
```
report_not_found: I2c928...
```
**解决**：用 `catalog-list` 和 `inspect-report` 重新获取正确的 ID。

### 7.3 筛选参数不匹配
```
filter_key_not_found: 首次分配开始时间
```
**解决**：用 `inspect-report --json` 查看报表支持的筛选参数，更新配置文件的 `overrides`。

### 7.4 超时
```
timeout: exceeded 180s
```
**解决**：增大 `--worker-timeout-sec` 参数，或减少 `--max-workers` 并发数。

---

## 八、注意事项

1. **凭据安全**：不要将用户名密码写入配置文件或代码
2. **并发控制**：`--max-workers` 建议不超过 5，避免对 SmartBI 服务器造成压力
3. **输出目录**：定期清理 `outputs/bi_exports/` 避免磁盘占满
4. **网络稳定**：确保能访问 `https://bi.61info.cn`
5. **首次配置**：所有报表 ID 必须填写完整才能批量下载

---

## 九、下一步：集成到服务周报自动化

1. 填写所有报表 ID（见第二节）
2. 测试单个任务下载（见 3.1）
3. 测试批量下载（见 3.2）
4. 集成到自动化脚本（见 5.1）
5. 修改后续数据处理脚本，从 smartbi-data-cli 的输出目录读取文件

**当前输出路径**（需要适配，**不要写死绝对路径**）：
```python
# ❌ 写死绝对路径（个人机器才有效，他人/换机即报路径不存在）
# BASE_DIR = Path("C:/Users/fengjianyi/Desktop/服务周报skill/6.1-.6.7周报数据")
# BASE_DIR = Path("C:/Users/fengjianyi/Desktop/smartbi-data-cli-internal-20260526/...") / run_date

# ✅ 项目内路径用 _paths.PROJECT_ROOT
from _paths import PROJECT_ROOT, resolve_smartbi_cli_dir
LEGACY_SOURCE = PROJECT_ROOT.parent / "6.1-.6.7周报数据"
BASE_DIR = resolve_smartbi_cli_dir() / "outputs" / "bi_exports" / "service_weekly" / run_date
```
