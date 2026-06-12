# 避坑指南（执行前必读）

> 本章节记录2026-06-07 P1试跑过程中发现的所有问题和修复方式。**下次执行前请逐项检查，避免重复踩坑**。

## 坑1：模板表头不能改（最容易犯的错误）

**问题**：把导出文件的列名改成"大账号ID"或"学员ID"，导致六一工作台API返回 `code: 102` 类型转换错误，无法识别上传文件中的ID。

**根本原因**：六一工作台标签上传API要求文件列名必须严格匹配模板原始列名"**用户id**"。

**修复方式**：
- `filter.py` / `filter_p1.py` 在拆分输出时必须重命名列：
  ```python
  dadou_df = filtered[["大账号ID"]].drop_duplicates().dropna()
  dadou_df.columns = ["用户id"]  # 必须保持模板原始列名
  user_df = filtered[["学员ID"]].drop_duplicates().dropna()
  user_df.columns = ["用户id"]   # 必须保持模板原始列名
  ```
- **铁律**：下载完模板文件后，处理导入数据外，**表头不能做任何改动**

## 坑2：BI报表读取需要header=13

**问题**：直接 `pd.read_excel()` 读取BI报表，列名都是"Unnamed: X"，因为前13行是表头说明。

**修复方式**：
```python
df = pd.read_excel(input_path, header=13)  # 表头在第14行（索引13）
```

## 坑3：六一工作台API成功判断错误

**问题**：把 `code != 200` 当作失败，但实际六一工作台API返回 `code: 0` 才是成功。

**修复方式**：
```python
body = r.get("body") if isinstance(r.get("body"), dict) else {}
code = body.get("code")

if code == 0:
    # 创建成功
elif code == 101 and "已存在" in body.get("msg", ""):
    # 已存在，复用
else:
    # 真失败
```

## 坑4：创建标签/用户群后没有直接返回ID

**问题**：以为API返回的data字段就是新ID，但实际API只返回 `{"code": 0, "msg": "OK"}`。

**修复方式**：
```python
# 创建后等异步落库
liuyi.wait_for_timeout(2000)
# 通过搜索API获取刚创建的ID
tag_id, info = find_tag_id(liuyi, name)
```

## 坑5：标签/用户群已存在时不应失败退出

**问题**：测试或重跑时，遇到"已存在"错误（code=101），整个流程中断。

**修复方式**：
- 已存在时（code=101 且 msg含"已存在"）：自动通过搜索API获取已有ID，复用即可
- 真失败时（其他code）：才退出

## 坑6：Playwright下载Chromium网络问题

**问题**：`python -m playwright install chromium` 经常因网络问题失败（ECONNRESET）。

**修复方式**：使用系统Chrome，不用Chromium
```python
browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=80)
```

## 坑7：北极星任务命名月份替换错误（P1容易犯）

**问题**：北极星历史任务名格式是"`【海外】26年5月教学协作任务-P1（学情反馈）`"，月份在中间。
- 错误正则 `r"-\d+月\b"` 匹配不到，导致变成 "`...26年5月...-6月`"
- 应该把"5月"改为"6月"

**修复方式**：
```python
def replace_month(name, target_month):
    # 优先匹配"N年N月"格式（北极星历史任务用这种格式）
    new, n = re.subn(r"(\d+年)(\d+)月", lambda m: f"{m.group(1)}{target_month}月", name)
    if n == 0:
        # 兜底：末尾的"-N月"
        new, n = re.subn(r"-\d+月\b", f"-{target_month}月", name)
        if n == 0:
            return f"{name}-{target_month}月"
    return new
```

## 坑8：北极星接口字段名

**问题**：以为教师任务完成时间字段叫 `teacherTaskCompleteTime`，实际叫 `latestServiceTime`。

**修复方式**：
```python
p["latestServiceTime"] = teacher_complete_time  # 不是 teacherTaskCompleteTime
```

## 坑9：北极星搜索关键词

- **P0**：`P0（学位预警）`
- **P1**：`P1（学情反馈）`（**不是"服务池"**）

## 坑10：用户群挂载规则

| 配置场景 | 挂载规则 |
|---------|---------|
| 北极星任务 | **只挂普通群**（不挂益智群） |
| 标签数据同步 | **只同步益智群**（不同步普通群） |
| 企微标签关联 | **只关联普通群** |
