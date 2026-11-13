$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot

# 始终使用当前工具文件夹的 Python 虚拟环境；绝不回退到系统 Python。
$venvPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host '未找到项目独立虚拟环境 .venv。' -ForegroundColor Yellow
    Write-Host '请先双击「双击首次安装环境.bat」，创建 .venv 并安装依赖。'
    Read-Host '【需要按键】请按 Enter 键结束'
    exit 1
}
$checkScript = Join-Path $PSScriptRoot '检查独立环境.py'
if (-not (Test-Path -LiteralPath $checkScript)) {
    Write-Host '缺少 检查独立环境.py，请重新解压完整工具包。' -ForegroundColor Red
    Read-Host '【需要按键】请按 Enter 键结束'
    exit 1
}
& $venvPython -B $checkScript
if ($LASTEXITCODE -ne 0) {
    Write-Host '当前 .venv 中的依赖不完整，请再次双击「双击首次安装环境.bat」。' -ForegroundColor Yellow
    Read-Host '【需要按键】请按 Enter 键结束'
    exit 1
}

$catalogScript = Join-Path $PSScriptRoot 'read_feishu_tables.py'
if (-not (Test-Path -LiteralPath $catalogScript)) {
    Write-Host '缺少 read_feishu_tables.py，请重新解压完整工具包。' -ForegroundColor Red
    exit 1
}

function Read-FeishuCatalog {
    param([string] $Python, [string[]] $Arguments, [string] $OutputPath)
    & $Python @Arguments | Out-Host
    if ($LASTEXITCODE -ne 0) { throw '获取飞书表格名称和子表清单失败；请核对凭据、表格链接及读取权限。' }
    return (Get-Content -LiteralPath $OutputPath -Encoding UTF8 -Raw | ConvertFrom-Json)
}

function Select-TargetTables {
    param([object[]] $Entries, [string] $AppName)
    Write-Host '【链接说明】粘贴链接用于确定整份多维表格，链接中的子表和视图不会自动作为处理范围；请在下方清单中另行选择。' -ForegroundColor Yellow
    Write-Host '【操作说明】请按子表名称核对后输入编号。实际处理范围以主菜单中的「当前选中」为准。'
    while ($true) {
        Write-Host ''
        Write-Host ('当前多维表格：' + $AppName) -ForegroundColor Cyan
        Write-Host '=============== 第一步：选择本次目标子表 ===============' -ForegroundColor Cyan
        for ($i = 0; $i -lt $Entries.Count; $i++) {
            Write-Host ('[{0}] {1}  子表ID：{2}' -f ($i + 1), $Entries[$i].name, $Entries[$i].table_id)
        }
        Write-Host '[0] 不选择任何子表（不读写，仍可返回菜单重新选择）'
        Write-Host '可单选 1、多选 1,2；输入 0 或直接回车表示不选择。' -ForegroundColor Yellow
        $typed = [string](Read-Host '【需要选择】请按上方名称输入子表编号；输入 0 或直接按 Enter 可暂不选择')
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
Write-Host '当前运行环境：' $venvPython
Write-Host 'App Secret 仅用于本次使用，不会保存到配置文件。'
$env:FEISHU_APP_ID = Read-Host '【需要输入】请输入飞书 App ID（飞书开放平台 → 凭证与基础信息）'
$secureSecret = Read-Host '【需要输入】请输入飞书 App Secret（输入内容不显示）' -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
try { $env:FEISHU_APP_SECRET = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
$sessionConfig = $null
try {
    $sessionConfig = Join-Path ([System.IO.Path]::GetTempPath()) ('feishu_word_' + [guid]::NewGuid().ToString('N') + '.json')
    $catalogArgs = @('-B', $catalogScript, '--output', $sessionConfig)
    $configPath = Join-Path $PSScriptRoot 'config.json'
    if (Test-Path -LiteralPath $configPath) {
        Write-Host '【操作说明】正在使用 config.json 中保存的目标多维表格；请在接下来的清单中选择所需子表。'
        $catalogArgs += @('--config', $configPath)
    } else {
        Write-Host '【操作说明】请打开目标多维表格，复制浏览器地址栏中包含 /base/ 的完整链接。示例：https://example.feishu.cn/base/abc123?table=tbl001&view=vew001'
        $appTokenInput = ([string](Read-Host '【需要输入】请粘贴多维表格完整链接，或填写 app_token')).Trim()
        if ([string]::IsNullOrWhiteSpace($appTokenInput)) { throw '未填写多维表格链接或 app_token。' }
        $catalogArgs += @('--app-token', $appTokenInput)
    }
    Write-Host '【操作说明】正在读取飞书中的多维表格名称和子表清单，请稍候；此步骤不会修改表格内容。'
    $configData = Read-FeishuCatalog -Python $venvPython -Arguments $catalogArgs -OutputPath $sessionConfig
    $tableEntries = @($configData.tables)
    # 启动时选择一次，菜单 1～4 共用当前选中范围。
    # 菜单 5 可随时重新选择；不选时不运行任何业务脚本。
    $selection = Select-TargetTables -Entries $tableEntries -AppName $configData.app_name
    $selected = @($selection.Items)
    while ($true) {
        Write-Host ''
        Write-Host '================ 飞书 Word 批量导入工具 V3.1 ================' -ForegroundColor Cyan
        Write-Host ('当前多维表格：' + $configData.app_name)
        if ($selected.Count -eq 0) {
            Write-Host '当前选中：无（所有读写菜单已禁用）' -ForegroundColor Yellow
        } else {
            Write-Host ('当前选中：' + (($selected | ForEach-Object { $_.name }) -join '、')) -ForegroundColor Green
        }
        Write-Host '【操作说明】更换子表请选择菜单 5；仅在飞书网页中切换子表或视图，不会改变本工具的选择。'
        Write-Host '1. 生成预览 + 差异审核 Excel（不修改飞书）'
        Write-Host '2. 正式只填空白 + 生成差异审核 Excel【写入飞书】'
        Write-Host '3. 审核覆盖预检（不修改飞书）'
        Write-Host '4. 根据 Excel 人工审批正式覆盖【写入飞书】'
        Write-Host '5. 重新选择目标子表'
        Write-Host '0. 退出'
        Write-Host '重要：菜单 2、4 写入需输入全部大写 YES，小写 yes 不会执行。' -ForegroundColor Yellow
        $choice = Read-Host '【需要选择】请输入操作编号；输入 0 可退出'
        if ($choice -eq '0') { break }
        if ($choice -eq '5') {
            Write-Host '【操作说明】正在刷新子表清单；请按名称重新选择，编号以接下来显示的清单为准。'
            $configData = Read-FeishuCatalog -Python $venvPython -Arguments $catalogArgs -OutputPath $sessionConfig
            $tableEntries = @($configData.tables)
            $selection = Select-TargetTables -Entries $tableEntries -AppName $configData.app_name
            $selected = @($selection.Items)
            continue
        }
        if ($choice -notin @('1', '2', '3', '4')) {
            Write-Host '无效选项，请输入 0、1、2、3、4 或 5。' -ForegroundColor Yellow
            continue
        }
        if ($selected.Count -eq 0) {
            Write-Host '【操作提示】尚未选择子表，未开始处理。请先选择菜单 5，再按名称选择所需子表；输入 0 可退出。' -ForegroundColor Yellow
            continue
        }

        if ($choice -eq '1' -or $choice -eq '2') {
            $argsList = @('01_导入并生成差异报告.py')
            if ($choice -eq '2') {
                $argsList += '--apply-blanks'
            }
        } else {
            $report = ([string](Read-Host '【需要输入】请粘贴审核完成的 Excel 文件完整路径（.xlsx；可从资源管理器复制文件地址）')).Trim('"')
            if ([string]::IsNullOrWhiteSpace($report)) { continue }
            $argsList = @('02_执行审核覆盖.py', '--report', $report)
            if ($choice -eq '4') {
                $argsList += '--execute'
            }
        }

        # 使用本次 API 获取的名称与原有文件夹映射，不改写用户的 config.json。
        $argsList += @('--config', $sessionConfig)
        # 将所选真实 table_id 逐个传入 Python；未选中的子表不会进入脚本一/二。
        foreach ($entry in $selected) {
            $argsList += @('--table-id', [string]$entry.table_id)
        }
        Write-Host ('【操作说明】本次仅处理以下选中子表：' + (($selected | ForEach-Object { $_.name }) -join '、')) -ForegroundColor Cyan
        if ($choice -eq '2' -or $choice -eq '4') {
            Write-Host '写入范围仅限上述子表；最终确认仍须输入大写 YES。' -ForegroundColor Yellow
        }
        & $venvPython @argsList
        if ($LASTEXITCODE -eq 0) {
            Write-Host '【操作结束】请查看上方处理结果，再选择下一项操作。'
        } else {
            Write-Host '【操作提示】处理未正常完成，请根据上方提示检查后再继续。' -ForegroundColor Yellow
        }
    }
} catch {
    Write-Host ('运行已停止：' + $_.Exception.Message) -ForegroundColor Red
    exit 1
} finally {
    if ($sessionConfig -and (Test-Path -LiteralPath $sessionConfig)) {
        Remove-Item -LiteralPath $sessionConfig -ErrorAction SilentlyContinue
    }
    $env:FEISHU_APP_ID = $null
    $env:FEISHU_APP_SECRET = $null
}
