1️⃣ crm-lesson-package/ - CRM课时包批量创建
作用：在豌豆思维CRM「财务 → 课时包管理/套餐管理」批量创建课时包和对应商品套餐

技术方案：

基于 Playwright 自动化 Element UI 表单填充
处理级联下拉、动态字段、表单校验等复杂交互
输入输出：

输入：Excel SKU配置表（横向KV格式）
输出：CSV日志 + 失败截图（logs/目录）
操作方式：


cd crm-lesson-package-skill
python crm_batch_create_lesson_packages.py --xlsx "<excel路径>" --skip-existing
运行频次：按需触发（有新SKU配置需求时）

关键文件：

crm_batch_create_lesson_packages.py - 批量执行
auth_state.json - 登录态缓存
config.local.json - 环境配置
2️⃣ lp_moments/ - LP企微朋友圈排期与话术生成
作用：为学管（Learning Planner）生成企微朋友圈发布排期和营销文案

核心能力：

区域差异化：港澳台（升学压力）vs 欧美澳（思维培养）
内容类型：情感共鸣、产品价值、学员成果、互动类
升阶季专项：官方通知、新阶段预告、转班说明、私聊话术
输出格式：

Markdown文件（output/目录）
包含：发布时间、内容类型、文案（≤150字）、配图建议、私聊脚本
运行频次：按营销周期或升阶季触发

目录结构：

content_library/ - 内容素材库
schedules/ - 排期管理
user_profiles/ - 用户画像
scripts/ - 生成脚本
3️⃣ service-incentive-calc/ - 服务激励核算自动化
作用：每月从BI报表提取服务绩效指标，自动计算团队激励金额

数据源（4张BI报表）：

益智海外新生首通监控
海外思维学管服务指标统计表
海外思维服务SOP执行情况
思维停课学员执行监控
核心指标（5个）：

指标	目标值	激励总额
首通及时跟进率	90%	200元
首课及时跟进率	80%	200元
首专及时跟进率	75%	400元
语义点执行率加和	2.0	600元
外呼跟进率	70%	600元
操作方式：


python build_incentive.py --month 5月
运行频次：每月一次（月中查进度 / 月初结算）

预计耗时：5-10分钟

输出路径：桌面/服务激励/{月份}/{月份}服务激励.xlsx

4️⃣ xueqing-review/ - 学情总结抽查自动化
作用：从BI导出教学协作跟进明细，按老师维度抽样评估学情总结质量

评估维度：

模板匹配：检查阶段性关键词（专注力、学习习惯、校内成绩等）
同质化检测：用SequenceMatcher算法计算文本相似度
评级规则：

🟢 优秀：符合模板 + 平均相似度 < 35%
🟡 及格：符合模板 + 平均相似度 ≥ 35%
🔴 不及格：不符合模板或疑似课后反馈
操作方式：


python xueqing_review.py \
  --input "<BI导出Excel>" \
  --output "<输出路径>" \
  --sample 20
运行频次：按需触发（质量检查时）

默认抽样：20个老师（优先选≥2条学情的老师）

输出内容（3个Sheet）：

抽查总览 - 老师信息、符合模板数、同质化评估、综合评级
详细条目 - 每条学情的具体内容和匹配情况
评估规则 - 规则说明文档
5️⃣ 周报自动化/ - 周报数据处理与飞书发布
作用：BI报表 → 数据清洗整合 → 结论生成 → 格式化Excel → 飞书电子表格嵌入文档

当前覆盖范围：4.1 服务指标跟进 & 语义分析（含AI学情助手）

完整流程（6个阶段）：

BI导出：使用 bi_skill 批量导出6份报表
整合宽表：consolidate_4_1_v2.py → _merged_4_1.xlsx（含辅助列）
AI学情单独成表：consolidate_4_1_ai.py → 独立格式化表
结论文本生成：conclusions_4_1.py → 纯文本结论（按缩进层级）
格式化Excel：export_4_1_excel.py → 多级表头 + 汇总行加色
飞书发布：create_feishu_sheets.py → 创建电子表格 + 文档嵌入
操作方式：


cd 周报自动化/weekly_report
python consolidate_4_1_v2.py
python consolidate_4_1_ai.py
python export_4_1_excel.py
python conclusions_4_1.py
python create_feishu_sheets.py
运行频次：人工触发（无自动调度），便于核对

设计原则：不向飞书文档写原生表格（>200行会卡顿），统一使用电子表格+文档嵌入

关键技术点：

团队级汇总：海外团队(合) + 9个组
辅助列：跟进率、接通率、企微绑定率、AI占比
Windows命令行≤8KB限制：sheets +write 分批3-5行
台湾组数据保留但不进结论
🔄 运行频次总结
工具	频次	触发方式
crm-lesson-package	按需	有新SKU配置时
lp_moments	按周期	营销周期/升阶季
service-incentive-calc	每月1次	月中/月末
xueqing-review	按需	质量检查时
周报自动化	按需	人工触发（周报制作时）
📦 技术依赖
Python环境：Python 3.10+

核心依赖：

pandas - 数据处理
openpyxl - Excel读写
playwright - 浏览器自动化
lark-cli - 飞书API操作
外部系统：

BI系统（Smartbi）- 数据源
豌豆思维CRM - 课时包管理
飞书 - 文档/电子表格发布
