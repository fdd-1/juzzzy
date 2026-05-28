# 学情积分核算 - Windows 定时任务注册脚本
# 每月1号和16号各触发一次
# 1号: 计算上月16号 ~ 上月最后一天
# 16号: 计算本月1号 ~ 本月15号

$TaskName1 = "学情积分核算_月初"
$TaskName16 = "学情积分核算_月中"
$PythonExe = "C:\Users\fengjianyi\AppData\Local\Programs\Python\Python312\python.exe"
$ScriptPath = "C:\Users\fengjianyi\Desktop\学情积分核算\xueqing_credit_skill.py"
$WorkDir = "C:\Users\fengjianyi\Desktop\学情积分核算"

# 删除已有同名任务
schtasks /Delete /TN $TaskName1 /F 2>$null
schtasks /Delete /TN $TaskName16 /F 2>$null

# 每月1号 09:30 触发 (计算上月16号~月底)
schtasks /Create /TN $TaskName1 `
    /TR "`"$PythonExe`" -X utf8 `"$ScriptPath`" run --auto" `
    /SC MONTHLY /D 1 /ST 09:30 `
    /RL HIGHEST /F

# 每月16号 09:30 触发 (计算本月1号~15号)
schtasks /Create /TN $TaskName16 `
    /TR "`"$PythonExe`" -X utf8 `"$ScriptPath`" run --auto" `
    /SC MONTHLY /D 16 /ST 09:30 `
    /RL HIGHEST /F

Write-Host ""
Write-Host "定时任务已注册:"
Write-Host "  $TaskName1  - 每月1号 09:30 (计算上月16号~月底)"
Write-Host "  $TaskName16 - 每月16号 09:30 (计算本月1号~15号)"
Write-Host "  执行命令: python xueqing_credit_skill.py run --auto"
Write-Host ""
Write-Host "查看任务: schtasks /Query /TN '学情积分核算_月初'"
Write-Host "手动触发: schtasks /Run /TN '学情积分核算_月初'"
Write-Host "删除任务: schtasks /Delete /TN '学情积分核算_月初' /F"

