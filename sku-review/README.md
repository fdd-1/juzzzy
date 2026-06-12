# SKU 复盘自动化 Skill

海外SKU复盘（港澳/欧美澳/台湾）：BI下载 → 数据筛选 → 人群匹配 → 指标计算 → 实际 vs 测算对标 → HTML/Excel 报告。

完整规则、命令、边界处理在 [SKILL.md](SKILL.md)。

## 快速上手（首次拿到这份 Skill）

### 1. 环境
- Python 3.10+
- 依赖：`pip install openpyxl`
- 配套 Skill：`bi_skill`（用于自动下载 BI 报表）

### 2. 配置 bi_skill 路径
本 Skill 通过 `config.py:_resolve_bi_skill_path()` 自动定位 `bi_skill.py`，依次查找：
1. 环境变量 `BI_SKILL_PATH`（推荐显式指定，跨电脑最稳）
2. `~/.workbuddy/skills/bi_skill/bi_skill.py`
3. 与本 Skill 同级的 `../bi_skill/bi_skill.py`
4. `../.workbuddy/skills/bi_skill/bi_skill.py`

如以上都不命中，先安装 bi_skill 或显式注入路径：
```powershell
$env:BI_SKILL_PATH = "D:\path\to\bi_skill\bi_skill.py"
```
```bash
export BI_SKILL_PATH="$HOME/path/to/bi_skill/bi_skill.py"
```

验证：
```bash
python -c "from config import BI_SKILL_PATH; print(BI_SKILL_PATH, BI_SKILL_PATH.exists())"
```

### 3. 准备每月数据
按月份新建 `data/{月份}/` 目录（中文如 `data/4月/`），放：
- 正式池：文件名含"正式池"（如 `4月正式池-04.01-v1.xlsx`）
- SKU 测算文件：文件名含 区域名+SKU（如 `港澳4月SKU复盘.xlsx`、`欧美澳4月SKU复盘.xlsx`、`台湾4月SKU复盘.xlsx`）
- BI 主订单宽表：可手动放（含"主订单宽表"字样），也可让脚本自动下载

### 4. 运行
```bash
# 单区域复盘（指定月份）
python run_sku_review.py --region gangao --month 6 --year 2026

# 单区域复盘（指定日期范围）
python run_sku_review.py --region gangao --start 2026-06-01 --end 2026-06-10

# 全区域批跑
python run_sku_review.py --region all --month 6 --year 2026

# 跳过BI下载用已有文件
python run_sku_review.py --region gangao --month 6 --year 2026 --skip-download
```

输出在 `output/{YYYYMMDD}_{月份}复盘/{区域}/`：
- `SKU复盘分析_{区域}_{月}.html` — 主报告
- `SKU精细对标_{区域}_{月}.xlsx` — 实际 vs 测算对标 + 套餐明细 + 计算口径
- `套餐明细_{区域}_{月}.csv` — 套餐粒度明细（用于跨月二次分析）

## 文件清单

| 文件/目录 | 作用 |
|---|---|
| `SKILL.md` | Skill 元数据 + 完整规则、命令、边界处理、反模式、校验清单 |
| `run_sku_review.py` | 主入口（带 `--region` 等参数） |
| `config.py` | 路径解析、区域筛选、套餐分类、人群规则、对标阈值 |
| `extract_data.py` | BI 报表 / 正式池 / 字段提取与人群匹配 |
| `analyze.py` | 节点/套餐聚合、SKU 测算文件预算提取（自适应表头） |
| `generate_report.py` | HTML 主报告 + Excel 对标表 + CSV 明细 |
| `export_audit.py` | 审计版 Excel（带匹配过程，便于人工核对） |
| `references/` | 字段映射、规则口径、阈值（只读参考；规则真实位置在 `config.py`） |
| `data/` | **不交接** — 每月业务数据（正式池、SKU测算、BI报表） |
| `output/` | **不交接** — 历次复盘输出 |

## 交接 / 跨电脑迁移注意

- **`data/` 和 `output/` 不参与交接**（已在 `.gitignore` 中），新电脑只需克隆代码 + `references/` + `SKILL.md`，按月添加自己的数据
- BI 下载依赖 `bi_skill` Skill — 请同步交接或在新电脑配置 `BI_SKILL_PATH`
- `config.py` 中的 `REGION_FILTERS` 是区域筛选条件，业务方更名（如"港澳益智教学服务区"被改）时改这里
- SKU 测算文件结构每月可能变化 — 详见 [`references/sku_budget_layout.md`](references/sku_budget_layout.md) 的修复路径

## 常见问题

| 现象 | 排查 |
|---|---|
| `预算数据: 0 个组合` | SKU 测算文件结构变了，参考 `references/sku_budget_layout.md` 锚点失效修复表 |
| `BI 下载失败` 重试 3 次失败 | 检查 BI 报表名是否更名 / 网络VPN / `BI_SKILL_PATH` 是否生效 |
| 筛选后订单为 0 | 检查 BI 字段名是否变更（如"五级部门"被改） / 正式池月份是否对齐 |
| `PermissionError: Excel 文件被占用` | 关闭已打开的 Excel 后重跑 |

更多反模式 / 校验清单见 [SKILL.md](SKILL.md)。
