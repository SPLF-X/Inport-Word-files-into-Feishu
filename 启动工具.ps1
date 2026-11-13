$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

# 始终使用当前工具文件夹的 Python 虚拟环境；绝不回退到系统 Python。
$venvPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host '未找到项目独立虚拟环境 .venv。' -ForegroundColor Yellow
    Write-Host '请先双击「双击首次安装环境.bat」，创建 .venv 并安装依赖。'
    Read-Host '按 Enter 结束'
    exit 1
}
$checkScript = Join-Path $PSScriptRoot '检查独立环境.py'
if (-not (Test-Path -LiteralPath $checkScript)) {
    Write-Host '缺少 检查独立环境.py，请重新解压完整工具包。' -ForegroundColor Red
    Read-Host '按 Enter 结束'
    exit 1
}
& $venvPython -B $checkScript
if ($LASTEXITCODE -ne 0) {
    Write-Host '当前 .venv 中的依赖不完整，请再次双击「双击首次安装环境.bat」。' -ForegroundColor Yellow
    Read-Host '按 Enter 结束'
    exit 1
}

if (-not (Test-Path -LiteralPath '.\config.json')) {
    Write-Host '未找到 config.json。请先复制 config.example.json 为 config.json，填写真实 app_token 和两个 table_id。' -ForegroundColor Yellow
    Read-Host '按 Enter 结束'
    exit 1
}


# 先从配置读取子表名称/ID，用编号选择，而不是由用户手动输入 table_id。
# Windows PowerShell 5.1 可使用 ConvertFrom-Json（不要求升级 PowerShell）。
try {
    $configData = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'config.json') -Encoding UTF8 -Raw | ConvertFrom-Json
    $tableEntries = @($configData.tables)
    if ($tableEntries.Count -lt 1) { throw 'config.json 的 tables 为空' }
    $seenIds = @{}
    foreach ($t in $tableEntries) {
        if ([string]::IsNullOrWhiteSpace([string]$t.name) -or [string]::IsNullOrWhiteSpace([string]$t.table_id)) {
            throw 'config.json 存在未填写名称或 table_id 的子表'
        }
        if ($seenIds.ContainsKey([string]$t.table_id)) { throw 'config.json 中存在重复的子表 ID' }
        $seenIds[[string]$t.table_id] = $true
    }
} catch {
    Write-Host ('读取 config.json 的子表清单失败：' + $_.Exception.Message) -ForegroundColor Red
    Read-Host '按 Enter 结束'
    exit 1
}

function Select-TargetTables {
    param([object[]] $Entries)
    while ($true) {
        Write-Host ''
        Write-Host '=============== 第一步：选择本次目标子表 ===============' -ForegroundColor Cyan
        for ($i = 0; $i -lt $Entries.Count; $i++) {
            Write-Host ('[{0}] {1}' -f ($i + 1), $Entries[$i].name)
        }
        Write-Host '[0] 不选择任何子表（不读写，仍可返回菜单重新选择）'
        Write-Host '可单选 1、多选 1,2；输入 0 或直接回车表示不选择。' -ForegroundColor Yellow
        $typed = [string](Read-Host '请输入子表编号')
        $typed = [regex]::Replace($typed, '\s+', '')
        if ($typed -eq '' -or $typed -eq '0') {
            return [pscustomobject]@{ Items = @() }
        }
        if ($typed -notmatch '^[1-9][0-9]*(,[1-9][0-9]*)*$') {
            Write-Host '输入格式无效。请用半角逗号分隔，例如 1,2；不能把 0 和其他编号一起选择。' -ForegroundColor Yellow
            continue
        }
        $indexes = $typed.Split(',')
        $seen = @{}
        $chosen = @()
        $valid = $true
        foreach ($idxText in $indexes) {
            $idx = 0
            if ((-not [int]::TryParse($idxText, [ref]$idx)) -or $idx -lt 1 -or $idx -gt $Entries.Count) {
                $valid = $false
                break
            }
            if ($seen.ContainsKey($idx)) {
                $valid = $false
                break
            }
            $seen[$idx] = $true
            $chosen += $Entries[$idx - 1]
        }
        if (-not $valid) {
            Write-Host '编号超出范围或重复，请重新选择。' -ForegroundColor Yellow
            continue
        }
        return [pscustomobject]@{ Items = @($chosen) }
    }
}

Write-Host '飞书 Word 批量导入工具 V3.1' -ForegroundColor Cyan
Write-Host '当前项目独立环境：' $venvPython
Write-Host '凭证只在当前启动进程使用，不会写入配置文件。'
$env:FEISHU_APP_ID = Read-Host '输入飞书 App ID（开发者后台 → 凭证与基础信息）'
$secureSecret = Read-Host '输入飞书 App Secret（输入时隐藏）' -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
try { $env:FEISHU_APP_SECRET = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }

try {
    # 方案 A：启动时选择一次，菜单 1～4 共用当前选中范围。
    # 菜单 5 可随时重新选择；不选时不运行任何业务脚本。
    $selection = Select-TargetTables -Entries $tableEntries
    $selected = @($selection.Items)
    while ($true) {
        Write-Host ''
        Write-Host '================ 飞书 Word 批量导入工具 V3.1 ================' -ForegroundColor Cyan
        if ($selected.Count -eq 0) {
            Write-Host '当前选中：无（所有读写菜单已禁用）' -ForegroundColor Yellow
        } else {
            Write-Host ('当前选中：' + (($selected | ForEach-Object { $_.name }) -join '、')) -ForegroundColor Green
        }
        Write-Host '1. 生成预览 + 差异审核 Excel（不修改飞书）'
        Write-Host '2. 正式只填空白 + 生成差异审核 Excel【写入飞书】'
        Write-Host '3. 审核覆盖预检（不修改飞书）'
        Write-Host '4. 根据 Excel 人工审批正式覆盖【写入飞书】'
        Write-Host '5. 重新选择目标子表'
        Write-Host '0. 退出'
        Write-Host '重要：菜单 2、4 写入需输入全部大写 YES，小写 yes 不会执行。' -ForegroundColor Yellow
        $choice = Read-Host '请输入操作编号'
        if ($choice -eq '0') { break }
        if ($choice -eq '5') {
            $selection = Select-TargetTables -Entries $tableEntries
            $selected = @($selection.Items)
            continue
        }
        if ($choice -notin @('1', '2', '3', '4')) {
            Write-Host '无效选项，请输入 0、1、2、3、4 或 5。' -ForegroundColor Yellow
            continue
        }
        if ($selected.Count -eq 0) {
            Write-Host '本次未选择子表；不读取 Word、不连接飞书，也不进行写入。请按 5 重新选择，或按 0 退出。' -ForegroundColor Yellow
            continue
        }

        if ($choice -eq '1' -or $choice -eq '2') {
            $argsList = @('01_导入并生成差异报告.py')
            if ($choice -eq '2') {
                $argsList += '--apply-blanks'
            }
        } else {
            $report = ([string](Read-Host '输入审核完成的 .xlsx 文件路径（可从资源管理器复制文件地址）')).Trim('"')
            if ([string]::IsNullOrWhiteSpace($report)) { continue }
            $argsList = @('02_执行审核覆盖.py', '--report', $report)
            if ($choice -eq '4') {
                $argsList += '--execute'
            }
        }

        # 将所选真实 table_id 逐个传入 Python；未选中的子表不会进入脚本一/二。
        foreach ($entry in $selected) {
            $argsList += @('--table-id', [string]$entry.table_id)
        }
        Write-Host ('本次操作限定子表：' + (($selected | ForEach-Object { $_.name }) -join '、')) -ForegroundColor Cyan
        if ($choice -eq '2' -or $choice -eq '4') {
            Write-Host '写入范围仅限上述子表；最终确认仍须输入大写 YES。' -ForegroundColor Yellow
        }
        & $venvPython @argsList
        Write-Host "本次命令退出码：$LASTEXITCODE"
    }
} finally {
    $env:FEISHU_APP_ID = $null
    $env:FEISHU_APP_SECRET = $null
}
