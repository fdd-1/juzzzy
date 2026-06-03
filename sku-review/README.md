# sku-review · 港澳 SKU 月度复盘自动化

每月对**港澳益智**主订单做 SKU 维度复盘：BI 主订单宽表 + 正式池 + SKU 测算文档 → 自动匹配人群 / 套餐 → 计算 ASP、单课时价、占比 → 与预算对标 → 输出 HTML 报告 + Excel 对标表 + CSV 明细。

> 这是一个 Claude Code Skill。也可以脱离 Claude 直接用 `python` 跑。

## 目录结构

```
sku-review/
├── SKILL.md                # Claude Skill 元信息（user-invocable: /sku_review）
├── run_sku_review.py       # 主入口
├── config.py               # 筛选条件 / 字段映射 / 套餐分类规则
├── extract_data.py         # 读 BI 宽表 + 正式池，按订单匹配人群
├── analyze.py              # 节点 / 套餐聚合，预算对标
├── export_audit.py         # 审计明细导出
├── generate_report.py      # HTML + Excel + CSV 输出
├── data/                   # 输入数据（按月份子目录，已 gitignore）
│   └── 4月/
│       ├── 海外益智主订单宽表-4月.xlsx   ← BI 自动下载或手动放入
│       ├── 4月正式池-04.01-v1.xlsx       ← 正式池
│       └── 港澳4月SKU复盘-0421.xlsx     ← SKU 测算
└── output/                 # 输出报告（按"YYYYMMDD_X月复盘"命名，已 gitignore）
    └── 20260602_4月复盘/
        ├── SKU复盘分析_4月.html
        ├── SKU精细对标_4月.xlsx
        └── 套餐明细_4月.csv
```

## 前置依赖

| 依赖 | 用途 | 安装 |
|------|------|------|
| Python 3.9+ | 运行环境 | — |
| openpyxl | 读写 Excel | `pip install openpyxl` |
| bi_skill | 自动下载 BI 主订单宽表 | 见下文 |

`bi_skill` 默认路径写在 [config.py](config.py)：

```python
BI_SKILL_PATH = Path(r"C:\Users\fengjianyi\.workbuddy\skills\bi_skill\bi_skill.py")
```

如果你的 `bi_skill` 装在别处，改这一行即可；或者直接用 `--skip-download` 跳过下载，手动把 BI 文件放到 `data/{月份}/`。

## 使用方式

### 一、Claude Code 内（推荐）

将本目录复制为 Skill（或在 Claude Code 中 `/skills add` 指向该目录），然后：

```
/sku_review
```

按提示填月份/年份即可。

### 二、命令行直接跑

```powershell
# 自动下载 BI + 全流程分析（指定月份）
python run_sku_review.py --month 4 --year 2026

# 指定任意日期范围
python run_sku_review.py --start 2026-04-01 --end 2026-04-30

# 跳过下载，只用 data/{月份}/ 里已经放好的文件
python run_sku_review.py --month 4 --year 2026 --skip-download
```

### 三、输入文件命名约定

放进 `data/{月份}/` 时，文件名必须包含以下关键词（脚本按通配符匹配）：

| 文件类型 | 关键词 | 示例 |
|----------|--------|------|
| BI 主订单宽表 | `主订单宽表` | `海外益智主订单宽表-4月.xlsx` |
| 正式池 | `正式池` | `4月正式池-04.01-v1.xlsx` |
| SKU 测算 | `SKU复盘` 或 `SKU` | `港澳4月SKU复盘-0421.xlsx` |

> 同名多版本时取**最新修改**的那个，会自动忽略 Excel 临时锁文件 `~$*.xlsx`。

## 业务规则

### 筛选条件（[config.py](config.py)）

下载后在 Python 中按字段筛选：

- 订单支付时业绩归属人五级部门 = 港澳益智教学服务区
- 区域等级 = 港澳

### 人群分类

读正式池 sheet `池内（剔已续不可续）`，按"当前课包顺序"字段：

| 当前课包顺序 | 人群 |
|--------------|------|
| 1 | 一续 |
| > 1 | 多续 |
| 0 | 池外（排除） |

### 套餐分类（按优先级匹配）

`升舱` → `早鸟` → `全量限定` → `学情限定` → `SVIP` → `全量` → `其余`

### 指标计算

- ASP = 总金额 / 订单数
- 含积分单课时价 = 总金额 / 课时数（含积分）
- 不含积分单课时价 = 总金额 / 课时数（不含积分）
- 占比 = 套餐订单数 / 同人群总订单数
- 对标 = 实际 vs SKU 测算文档中的预算

## 输出说明

每次执行会在 `output/YYYYMMDD_X月复盘/` 下生成 3 份文件：

1. **`SKU复盘分析_{月份}.html`** — 可直接在浏览器打开/分享，含节点表 + 套餐对标 + 偏差高亮
2. **`SKU精细对标_{月份}.xlsx`** — Excel 格式对标表，带预算列和偏差列
3. **`套餐明细_{月份}.csv`** — 套餐级原始聚合数据，方便二次透视

`审计明细_{月份}.xlsx` 由 [export_audit.py](export_audit.py) 单独导出，可手动调用做订单级核查。

## 常见问题

**1. BI 下载超时？**
脚本内置 3 次重试。仍失败可手动从 BI 下载，文件名含「主订单宽表」放入 `data/{月份}/`，再用 `--skip-download` 重跑。

**2. 没有 SKU 测算文件？**
预算对标会自动跳过，只输出实际指标，不影响主流程。

**3. 想改筛选条件 / 套餐规则？**
全部集中在 [config.py](config.py)，改完直接重跑即可。
