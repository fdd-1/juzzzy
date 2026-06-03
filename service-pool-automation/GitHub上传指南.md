# 上传到 GitHub 操作指南

## 已完成的步骤 ✅

1. ✅ 初始化 git 仓库
2. ✅ 创建 `.gitignore`（已排除敏感文件和数据文件）
3. ✅ 创建 `config.json.example`（配置文件示例）
4. ✅ 首次提交已完成

## 下一步：手动创建 GitHub 仓库并推送

### 步骤 1：在 GitHub 网站创建新仓库

1. 打开浏览器，访问 https://github.com/new
2. 填写仓库信息：
   - **Repository name**: `service-pool-automation`（或你喜欢的名字）
   - **Description**: `海外益智服务池学员自动化处理工具 - BI下载 → 数据匹配 → 六一工作台 → 豌豆数仓同步`
   - **Public/Private**: 建议选 **Private**（包含业务流程）
   - ⚠️ **不要勾选** "Add a README file"
   - ⚠️ **不要勾选** "Add .gitignore"
   - ⚠️ **不要勾选** "Choose a license"
3. 点击 **Create repository** 按钮

### 步骤 2：推送本地代码到 GitHub

创建完仓库后，GitHub 会显示推送命令。复制你的仓库 URL，然后在 PowerShell 执行：

```powershell
cd "C:\Users\fengjianyi\Desktop\服务池拆解&上传"

# 添加远程仓库（替换成你的 GitHub 用户名）
git remote add origin https://github.com/你的用户名/service-pool-automation.git

# 推送到 GitHub
git push -u origin master
```

**如果遇到认证问题**，GitHub 会提示你登录。

### 步骤 3：验证上传成功

刷新 GitHub 仓库页面，应该能看到：
- ✅ README.md（完整使用说明）
- ✅ SKILL.md（Skill 调用指南）
- ✅ service_pool_automation.py（主脚本）
- ✅ sync_tag_data.py（同步脚本）
- ✅ config.json.example（配置示例）
- ✅ 目录结构（data/, logs/）

## 已排除的敏感文件

以下文件**不会**被上传到 GitHub（已在 `.gitignore` 中）：

- `config.json`（包含真实路径）
- `data/downloads/*`（BI 原始数据）
- `data/processed/*`（处理结果）
- `logs/*`（日志文件）
- `auth_state.json`（六一工作台登录态）
- `*.xlsx`（Excel 数据文件）
- `学员流转定稿-终版v1.xlsx`（学员数据）

## 以后更新代码

修改代码后：

```powershell
cd "C:\Users\fengjianyi\Desktop\服务池拆解&上传"

# 查看修改了哪些文件
git status

# 添加修改的文件
git add .

# 提交修改
git commit -m "描述你做了什么修改"

# 推送到 GitHub
git push
```

## 仓库结构（GitHub 上的）

```
service-pool-automation/
├── README.md                     # 完整使用说明
├── SKILL.md                      # Skill 调用指南
├── service_pool_automation.py    # 主入口（7步自动化）
├── sync_tag_data.py              # 标签数据同步
├── debug_page.py                 # 调试工具
├── config.json.example           # 配置文件示例
├── .gitignore                    # Git 忽略文件
├── 操作文档.md                   # 操作流程
├── 实施完成总结.md                # 实施总结
├── data/
│   ├── downloads/.gitkeep
│   └── processed/.gitkeep
└── logs/.gitkeep
```

## 常见问题

### Q: 推送时提示 "Authentication failed"

**A**: 使用 Personal Access Token 认证：

1. 访问 https://github.com/settings/tokens
2. 点击 "Generate new token (classic)"
3. 勾选 `repo` 权限
4. 复制生成的 token
5. 推送时，密码处输入这个 token（不是你的 GitHub 密码）

### Q: 不小心提交了敏感文件怎么办？

**A**: 从历史中删除：

```powershell
# 从历史记录中删除文件
git filter-branch --force --index-filter "git rm --cached --ignore-unmatch 敏感文件名" --prune-empty --tag-name-filter cat -- --all

# 强制推送（危险操作，确认后再执行）
git push origin --force --all
```

### Q: 如何下载到其他电脑？

**A**: 克隆仓库：

```powershell
git clone https://github.com/你的用户名/service-pool-automation.git
cd service-pool-automation

# 复制 config.json.example 并修改
copy config.json.example config.json
# 然后编辑 config.json 填入真实路径
```
