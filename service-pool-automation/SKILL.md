# 海外益智服务池学员名单 Skill

## 触发场景

每月初执行，生成海外益智服务池学员名单并同步到六一工作台。

## 前置条件

- 当月学员流转文件已就绪（`每月池子/X月/学员流转定稿-终版v1.xlsx`）
- 六一工作台已登录（`auth_state.json` 有效）

## 每月需要更新的配置

编辑 `config.json`，修改两处：

```json
{
  "current_month": "2026-07",
  "liuzhuang_file": "C:\\Users\\fengjianyi\\Desktop\\每月池子\\7月\\学员流转定稿-终版v1.xlsx"
}
```

## 完整自动化命令

### 一键全流程（推荐）

```bash
cd C:\Users\fengjianyi\Desktop\服务池拆解&上传
python service_pool_automation.py
```

### 分步执行

```bash
# Step 1-4：只跑数据处理，跳过六一工作台
python service_pool_automation.py --skip-liuyi

# Step 5：上传标签和用户群（需要 auth_state.json 有效）
cd C:\Users\fengjianyi\Desktop\六一标签
python create.py \
  --input "C:\Users\fengjianyi\Desktop\服务池拆解&上传\data\processed\dadou_ids_YYYYMM.xlsx" \
  --id-type wandou \
  --tag-name "【海外益智】26年X月服务池学员名单" \
  --group-name "【海外益智】26年X月服务池学员名单"

# Step 6：标签数据同步到豌豆数仓
cd C:\Users\fengjianyi\Desktop\服务池拆解&上传
python sync_tag_data.py --user-group-name "【海外益智】26年X月服务池学员名单"
```

## 7步流程说明

| 步骤 | 内容 | 脚本 |
|------|------|------|
| 1 | BI 下载：海外思维续费规划表_新版_26年启用（筛选：开课M=当月1号，退费结束=上月末，池子节点3=服务月） | smartbi_browser_export.py |
| 2 | 数据筛选：是否可续学员=1，月初是否续费=空白 | service_pool_automation.py |
| 3 | 匹配：BI 大账号ID ↔ 学员流转 学员id → 最终归属LP、最终归属小组 | service_pool_automation.py |
| 4 | 输出 Excel（Sheet1 完整数据 + Sheet2 大账号ID） | service_pool_automation.py |
| 5 | 六一工作台新建标签 + 用户群（`【海外益智】26年X月服务池学员名单`） | 六一标签/create.py |
| 6 | 标签数据同步：新增 → 填写表单 → 确认 | sync_tag_data.py |
| 7 | 完成日志统计 | — |

## 输出文件

```
data/processed/
├── 服务池学员_YYYYMM.xlsx   # Sheet1: 完整数据含最终归属LP/小组
│                             # Sheet2: 大账号ID（六一标签数据）
└── dadou_ids_YYYYMM.xlsx    # 单独的大账号ID文件（create.py 用）
```

## 标签命名规则

`【海外益智】{年份}年{月份}月服务池学员名单`

例：`【海外益智】26年6月服务池学员名单`

## 登录六一工作台（登录态过期时）

```bash
python C:\Users\fengjianyi\Desktop\六一标签\login.py
```

## 关键文件路径

| 文件 | 路径 |
|------|------|
| 主脚本 | `服务池拆解&上传/service_pool_automation.py` |
| 同步脚本 | `服务池拆解&上传/sync_tag_data.py` |
| 配置文件 | `服务池拆解&上传/config.json` |
| 六一登录态 | `Desktop/六一标签/auth_state.json` |
| BI 导出工具 | `smartbi-data-cli-internal-20260526/.../smartbi_browser_export.py` |
| 六一标签工具 | `Desktop/六一标签/create.py` |

## 常见问题

| 问题 | 原因 | 解决 |
|------|------|------|
| BI 下载失败：账号在其他地方登录 | BI 账号被占用 | 关闭其他 BI 标签页后重新运行 |
| 六一工作台找不到「新增」按钮 | auth_state.json 过期 | 重新运行 `login.py` |
| 匹配数为 0 | 学员流转文件路径或月份不对 | 检查 config.json 的 `liuzhuang_file` |
