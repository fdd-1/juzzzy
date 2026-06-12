# 服务周报自动化 - 一键执行脚本
# 每周一 7:30 由 Windows 定时任务调用
#
# 凭据来源：当前 PowerShell 会话的 $env:SMARTBI_USERNAME / $env:SMARTBI_PASSWORD
# 计划任务方式：在创建任务时把这两个变量写到任务的 "环境变量" 配置里
# 切勿在本脚本里写明文凭据。

# 1. 校验凭据
if (-not $env:SMARTBI_USERNAME -or -not $env:SMARTBI_PASSWORD) {
    Write-Error "[X] 缺少 SmartBI 凭据。请先设置 `$env:SMARTBI_USERNAME / `$env:SMARTBI_PASSWORD（不要写进脚本）。"
    exit 1
}
$env:PYTHONIOENCODING = "utf-8"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 2. 项目目录（基于本脚本位置推算，跨机器可移植）
$PROJECT_DIR = Split-Path -LiteralPath (Split-Path -LiteralPath $PSCommandPath -Parent) -Parent
Set-Location -LiteralPath $PROJECT_DIR

# 3. 日志目录
$LOG_DIR = "$PROJECT_DIR\logs"
if (-not (Test-Path $LOG_DIR)) {
    New-Item -ItemType Directory -Path $LOG_DIR -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LOG_FILE = "$LOG_DIR\weekly_run_$timestamp.log"

function Log {
    param($msg)
    $line = "[$(Get-Date -Format 'HH:mm:ss')] $msg"
    Write-Host $line
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
}

Log "=================================="
Log "服务周报自动化 - 开始执行"
Log "=================================="

# 计算时间窗口
$today = Get-Date
$daysSinceMonday = ($today.DayOfWeek.value__ + 6) % 7
if ($daysSinceMonday -eq 0) { $daysSinceMonday = 7 }
$lastMonday = $today.AddDays(-$daysSinceMonday - 7)
$lastSunday = $lastMonday.AddDays(6)
$monthStart = Get-Date -Year $today.Year -Month $today.Month -Day 1
$runDate = $today.ToString("yyyy-MM-dd")
$startStr = $lastMonday.ToString("yyyyMMdd")
$endStr = $lastSunday.ToString("yyyyMMdd")
$startDate = $lastMonday.ToString("yyyy-MM-dd")
$endDate = $lastSunday.ToString("yyyy-MM-dd")

Log "时间窗口：$startDate ~ $endDate（4.1）"
Log "时间窗口：$($monthStart.ToString('yyyy-MM-dd')) ~ $endDate（4.2-4.6）"

# Step 1: 下载报表
Log ""
Log "=== Step 1: 下载 SmartBI 报表 ==="
try {
    python scripts\download_smartbi_reports.py --output-dir "downloads\smartbi_reports" 2>&1 | ForEach-Object { Log $_ }
    python scripts\download_4_5_fuwuyue.py 2>&1 | ForEach-Object { Log $_ }
    Log "✓ 报表下载完成"
} catch {
    Log "✗ 报表下载失败: $_"
    exit 1
}

# Step 2: 数据整合
Log ""
Log "=== Step 2: 数据整合与格式化 ==="
$downloadsDir = "downloads\smartbi_reports\$runDate"
$exportDir = "exports\weekly_${startStr}_${endStr}"

try {
    python scripts\process_data.py --downloads-dir $downloadsDir --output-dir $exportDir 2>&1 | ForEach-Object { Log $_ }
    Log "✓ 数据整合完成"
} catch {
    Log "✗ 数据整合失败: $_"
    exit 1
}

# Step 3: 创建飞书表格
Log ""
Log "=== Step 3: 创建飞书电子表格 ==="

# 复制文件到子目录（feishu_builder_4_1 需要）
Copy-Item "$exportDir\_merged_4_1.xlsx" "$exportDir\4_1\_merged_4_1.xlsx" -Force -ErrorAction SilentlyContinue

$tokens = @{}

try {
    # 4.1
    $out = python modules\feishu_builder_4_1.py --merged "$exportDir\_merged_4_1.xlsx" --start-date $startDate --end-date $endDate 2>&1
    $out | ForEach-Object { Log $_ }
    if ($out -match '"token":\s*"([^"]+)"') {
        $tokens["4.1"] = $matches[1]
    }

    # 4.2-4.6
    $tasks = @(
        @{title="4.2 组班意向 $($lastMonday.ToString('MMdd'))-$($lastSunday.ToString('MMdd'))"; file="_merged_4_2.xlsx"; key="4.2"},
        @{title="4.3 群发消息 $($lastMonday.ToString('MMdd'))-$($lastSunday.ToString('MMdd'))"; file="_merged_4_3.xlsx"; key="4.3"},
        @{title="4.4 停课唤醒 $($lastMonday.ToString('MMdd'))-$($lastSunday.ToString('MMdd'))"; file="_merged_4_4.xlsx"; key="4.4"},
        @{title="4.5 服务月跟进 $($lastMonday.ToString('MMdd'))-$($lastSunday.ToString('MMdd'))"; file="_merged_4_5_fuwuyue.xlsx"; key="4.5_fuwuyue"},
        @{title="4.5 服务池SOP $($lastMonday.ToString('MMdd'))-$($lastSunday.ToString('MMdd'))"; file="_merged_4_5_sop.xlsx"; key="4.5_sop"},
        @{title="4.6 外呼监控 $($lastMonday.ToString('MMdd'))-$($lastSunday.ToString('MMdd'))"; file="_merged_4_6_waihu.xlsx"; key="4.6_waihu"},
        @{title="4.6 企微回复 $($lastMonday.ToString('MMdd'))-$($lastSunday.ToString('MMdd'))"; file="_merged_4_6_qiwei.xlsx"; key="4.6_qiwei"}
    )

    foreach ($task in $tasks) {
        $out = python modules\feishu_simple_builder.py --input "$exportDir\$($task.file)" --title $task.title 2>&1
        $out | ForEach-Object { Log $_ }
        if ($out -match '"token":\s*"([^"]+)"') {
            $tokens[$task.key] = $matches[1]
        }
    }

    # 保存 tokens
    $tokens | ConvertTo-Json | Out-File "$exportDir\sheet_tokens.json" -Encoding UTF8
    Log "✓ 8 个飞书表格已创建"

    # 更新 final_doc_builder_v3 的 SHEET_TOKENS（动态写入）
    Log "更新 final_doc_builder_v3 的 token..."
    $builderFile = "modules\final_doc_builder_v3.py"
    $content = Get-Content $builderFile -Encoding UTF8 -Raw

    foreach ($key in $tokens.Keys) {
        $newToken = $tokens[$key]
        # 替换该 key 对应的 token
        $pattern = "(`"$key`":\s*)`"[^`"]+`""
        $replacement = "`"$key`": `"$newToken`""
        $content = $content -replace $pattern, $replacement
    }

    Set-Content -Path $builderFile -Value $content -Encoding UTF8
    Log "✓ token 已更新到 final_doc_builder_v3"
} catch {
    Log "✗ 飞书表格创建失败: $_"
    exit 1
}

# Step 4: 飞书表格优化（删空行 + 上色阶）
Log ""
Log "=== Step 4: 飞书表格优化（删空行 + 色阶）==="
try {
    # 更新 polish_feishu_sheets.py 的 token
    $polishFile = "scripts\polish_feishu_sheets.py"
    $polishContent = Get-Content $polishFile -Encoding UTF8 -Raw
    foreach ($key in $tokens.Keys) {
        $newToken = $tokens[$key]
        $pattern = "(`"$key`":\s*\{[^}]*?`"token`":\s*)`"[^`"]+`""
        $replacement = "`${1}`"$newToken`""
        $polishContent = $polishContent -replace $pattern, $replacement
    }
    Set-Content -Path $polishFile -Value $polishContent -Encoding UTF8

    python scripts\polish_feishu_sheets.py 2>&1 | ForEach-Object { Log $_ }
    Log "✓ 飞书表格优化完成"
} catch {
    Log "✗ 飞书表格优化失败: $_"
    # 不退出，继续生成文档
}

# Step 5: 生成结论 + 创建最终文档
Log ""
Log "=== Step 5: 生成结论 + 创建最终文档 ==="
try {
    # 生成 4.1 callout
    python modules\conclusion_4_1.py --merged "$exportDir\_merged_4_1.xlsx" 2>&1 | Out-File -Encoding utf8 "$exportDir\4_1\_callout_4_1.xml"

    # 复制其他 _merged 文件到对应子目录
    Copy-Item "$exportDir\_merged_4_2.xlsx" "$exportDir\4_2\_merged_4_2_v2.xlsx" -Force -ErrorAction SilentlyContinue
    Copy-Item "$exportDir\_merged_4_3.xlsx" "$exportDir\4_3\_merged_4_3_v2.xlsx" -Force -ErrorAction SilentlyContinue
    Copy-Item "$exportDir\_merged_4_4.xlsx" "$exportDir\4_4\_merged_4_4_v3.xlsx" -Force -ErrorAction SilentlyContinue
    Copy-Item "$exportDir\_merged_4_5_fuwuyue.xlsx" "$exportDir\4_5\_merged_4_5_fuwuyue_v2.xlsx" -Force -ErrorAction SilentlyContinue
    Copy-Item "$exportDir\_merged_4_5_sop.xlsx" "$exportDir\4_5\_merged_4_5_sop_v2.xlsx" -Force -ErrorAction SilentlyContinue
    Copy-Item "$exportDir\_merged_4_6_waihu.xlsx" "$exportDir\4_6\_merged_4_6_waihu_v2.xlsx" -Force -ErrorAction SilentlyContinue
    Copy-Item "$exportDir\_merged_4_6_qiwei.xlsx" "$exportDir\4_6\_merged_4_6_qiwei_v2.xlsx" -Force -ErrorAction SilentlyContinue

    $out = python modules\final_doc_builder_v3.py 2>&1
    $out | ForEach-Object { Log $_ }

    # 提取 doc_id
    $docId = $null
    if ($out -match '"doc_id":\s*"([^"]+)"') {
        $docId = $matches[1]
    }

    if ($docId) {
        Log "✓ 文档已创建: $docId"

        # 移动到目标文件夹
        Log "移动到目标文件夹..."
        lark-cli drive +move --file-token $docId --folder-token "JpSRflVoWlwxZxdBgg7cFbBNnrc" --type docx 2>&1 | ForEach-Object { Log $_ }

        Log ""
        Log "=================================="
        Log "✅ 服务周报自动化完成"
        Log "=================================="
        Log "文档 URL: https://hcnig43mb8gp.feishu.cn/docx/$docId"
    } else {
        Log "✗ 无法提取 doc_id"
        exit 1
    }
} catch {
    Log "✗ 文档生成失败: $_"
    exit 1
}

Log ""
Log "执行结束时间: $(Get-Date)"
