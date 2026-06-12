# 常见问题 FAQ

学情积分核算 skill 运行中常见问题与处理方案。

## Q: OA 提交失败，URL 还是 `method=add`？
A: 说明字段验证未通过。先看 `scripts/oa_login/submit_step3_filled.png` 截图判断哪个字段空。多数是 K2 widget 初始化没就绪 —— 重跑一次通常就过了。

## Q: 部门字段填了但被误识别为「首页」之类？
A: 双列布局陷阱 + autocomplete 没刷新。已修复（`fill_inputselect_field`），靠 `INPUTSELECT_SEARCH["积分成本归属部门"] = "海外教学服务运营组"` 关键词触发。

## Q: 定时任务没执行？
A:
1. 检查注册：`schtasks /Query /TN "学情积分核算_月初"`
2. 任务计划程序 → 查看任务历史
3. 手动触发：`schtasks /Run /TN "学情积分核算_月初"`

## Q: BI 报表下载提示 `extendsion_loginlock_forbid` / "账号在其他地方登录"？
A: SmartBI 限制同账号同时登录。两次连续运行（或 headful 调试 + 正式运行）会冲突。**等 5-10 秒**让会话释放后重试即可。

## Q: 报表显示「下载完成 0 bytes」，但程序没报错？
A: 历史 bug，已修复（v3.1）。原因是当 `rowCount > max_rows` 时浏览器跳过导出但程序不报错。现在会明确 raise `SmartbiBrowserExportError: Export skipped`。如果再遇到，先调大 `max_rows`，或对该报表配置 `split_days` 分段下载。

## Q: 上课明细下载报 `JavaScript heap out of memory`？
A: 浏览器 V8 堆 ~2GB 上限不足以装 9w+ 行数据。**Playwright `args: ["--js-flags=--max-old-space-size=8192"]` 不生效**（实测）。解决方案：在 `configs/smartbi_simple_report_tasks.json` 给报表配 `"split_days": 4`，按 4 天分段下载然后合并。已为「上课明细」默认开启。

## Q: 合并后的 xlsx 只有几 KB，数据全没了？
A: openpyxl `read_only=True` 模式读 SmartBI 导出的 xlsx 时 `iter_rows()` 会返回空。`merge_xlsx_files` 已改为非 read_only 模式（v3.1）。同时增加合并后大小校验：合并文件 < 段文件总和 30% 会 raise 并保留段文件，方便人工排查/重合并。

## Q: 处理时报 `PermissionError: ...上课明细_带标注.xlsx`？
A: 输出文件被 Excel 占用。看 `01_bi_exports/` 或 `03_output/<期次>/` 是否有 `~$xxx.xlsx` 临时锁文件 —— 关闭对应 Excel 窗口再跑。`run.py` 的 `find_latest_xlsx` 已自动跳过 `~$` 文件。

## Q: 数据处理结果是 0 行 / 0 积分，但 BI 报表是有数据的？
A: 大概率是 `01_bi_exports/` 里有同名旧文件被错用。**v3.1 起 BI 文件名带时间段后缀**（如 `海外思维学员上课明细_20260516-20260531.xlsx`），`run.py` 也按后缀精确匹配。如果文件名没后缀，删了重跑。

## Q: 提交 OA 时点完提交按钮，脚本报 `TargetClosedError: Target page, context or browser has been closed`？
A: OA 系统在提交成功后会跳转/关闭弹窗，脚本最后等待 5s 验证时浏览器已关闭。**通常意味着提交已成功**，但脚本无法自动确认。处理：
1. 登录 OA →「我的申请」查看是否有刚提交的豌豆币添加申请
2. 没看到再 `python xueqing_credit_skill.py submit-oa --yes` 重提
