# 海外益智服务池学员自动化处理工具

每月一键生成海外益智服务池学员名单 → 上传六一工作台 → 同步豌豆数仓。

## 一句话说明

```bash
cd C:\Users\fengjianyi\Desktop\服务池拆解&上传
python service_pool_automation.py
```

7 个步骤全自动跑完，预计 2-4 分钟（取决于 BI 下载速度）。

## 每月需要改两处配置

`config.json` 里改：

```json
{
  "current_month": "2026-06",
  "liuzhuang_file": "C:\\Users\\fengjianyi\\Desktop\\每月池子\\6月\\学员流转定稿-终版v1.xlsx"
}
```

> 月份变更时跟着改这两个字段，其他不用动。

## 7 步流程

| # | 步骤 | 输出 |
|---|------|------|
| 1 | BI 下载（海外思维续费规划表_新版_26年启用，筛选 开课M=当月1号 / 退费结束=上月末 / 池子节点3=服务月） | `data/downloads/服务池学员_YYYYMM.xlsx` |
| 2 | 二次筛选（是否可续学员=1 + 月初是否续费=空白） | 内存 |
| 3 | 与「学员流转定稿-终版v1」匹配（BI 大账号ID ↔ 流转 学员id），算最终归属LP / 最终归属小组 | 内存 |
| 4 | 保存结果（双 Sheet + 大账号ID 独立文件） | `data/processed/服务池学员_YYYYMM.xlsx` + `dadou_ids_YYYYMM.xlsx` |
| 5 | 六一工作台建标签 + 用户群（标签名：`【海外益智】26年X月服务池学员名单`） | tag_id, group_id |
| 6 | 标签数据同步：新增 → 业务类型=豌豆 / 业务系统=豌豆数仓 / 用户群=本月名 / 频率=每天 / 状态=启用 → 确认 | 同步配置已建 |
| 7 | 完成日志统计 | logs/ 下 |

## 输出文件长啥样

### `服务池学员_YYYYMM.xlsx`（双 Sheet）

- **Sheet1**：完整数据（91列），最右两列是 `最终归属LP` / `最终归属小组`
- **Sheet2** (`六一标签数据`)：仅 `大账号ID` 一列

### `dadou_ids_YYYYMM.xlsx`

单列大账号ID，专门用来喂给 `create.py`（它默认读第一个 Sheet）。

## 命令行参数

```bash
python service_pool_automation.py [选项]
```

| 选项 | 说明 |
|------|------|
| `--month YYYY-MM` | 指定月份（默认从 config.json 读） |
| `--liuzhuang-file PATH` | 临时指定学员流转文件 |
| `--skip-liuyi` | 只跑 1-4 步，跳过六一工作台 |
| `--dry-run` | 全部步骤模拟跑，不真发请求 |
| `--config PATH` | 指定其他配置文件 |

## 分步执行（出错时排查用）

```bash
# 只跑数据处理（1-4 步）
python service_pool_automation.py --skip-liuyi

# 单跑标签上传
python C:\Users\fengjianyi\Desktop\六一标签\create.py \
  --input "C:\Users\fengjianyi\Desktop\服务池拆解&上传\data\processed\dadou_ids_202606.xlsx" \
  --id-type wandou \
  --tag-name "【海外益智】26年6月服务池学员名单" \
  --group-name "【海外益智】26年6月服务池学员名单"

# 单跑标签数据同步
python sync_tag_data.py --user-group-name "【海外益智】26年6月服务池学员名单"
```

## 前置依赖

- **登录态**：`C:\Users\fengjianyi\Desktop\六一标签\auth_state.json`
  - 过期了重跑：`python C:\Users\fengjianyi\Desktop\六一标签\login.py`
- **BI 账号**：内置 76218 / 123456，跑前确认浏览器没占用 BI 登录
- **学员流转文件**：当月的 `学员流转定稿-终版v1.xlsx` 已经放好

## 关键文件

| 文件 | 作用 |
|------|------|
| `service_pool_automation.py` | 主入口，7 步串起来 |
| `sync_tag_data.py` | 标签数据同步（新增 → 填表 → 确认） |
| `config.json` | 月份、路径、筛选条件配置 |
| `SKILL.md` | Skill 调用指南 |

## 实战记录（2026年6月）

- BI 下载：3253 行 → 筛选 2381 人
- 学员流转匹配：99 人匹配到，2282 人沿用原 LP
- 六一标签：`【海外益智】26年6月服务池学员名单`（tag_id=19125）
- 用户群：同名（group_id=17326）
- 数仓同步：已配置（豌豆数仓 / 每天 / 启用）

## 排错

| 现象 | 原因 | 处理 |
|------|------|------|
| BI 报错「账号在其他地方登录」 | 浏览器有人在登 BI | 关闭其他 BI 标签后重跑 |
| 六一找不到「新增」按钮 | auth_state.json 过期 | 跑 login.py 重新扫码 |
| 匹配数为 0 | 用了上月学员流转文件 | 改 config.json 的 `liuzhuang_file` |
| 标签创建报 `code: 101` | 上传文件格式问题 | 检查 `dadou_ids_*.xlsx` 第一列是不是数字 |
