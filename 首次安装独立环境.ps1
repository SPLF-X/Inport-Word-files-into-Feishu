# 只将依赖安装进当前项目的 .venv，不影响电脑上的其他 Python 项目。
Write-Host '飞书 Word 批量导入工具 V3.1 - 独立虚拟环境安装' -ForegroundColor Cyan
$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$venvFolder = Join-Path $PSScriptRoot '.venv'
$venvPython = Join-Path $venvFolder 'Scripts\python.exe'
$requirements = Join-Path $PSScriptRoot 'requirements.txt'

if (-not (Test-Path -LiteralPath $requirements)) {
    Write-Host '错误：当前文件夹找不到 requirements.txt。请解压完整的工具包。' -ForegroundColor Red
    exit 1
}

function Test-PythonVersion {
    param([string] $CommandName, [string[]] $Prefix)
    if (-not (Get-Command $CommandName -ErrorAction SilentlyContinue)) { return $false }
    & $CommandName @Prefix -c 'import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)' 2>$null
    return ($LASTEXITCODE -eq 0)
}

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host '正在当前工具文件夹内创建独立 Python 环境 .venv ...' -ForegroundColor Cyan
    if (Test-PythonVersion -CommandName 'py' -Prefix @('-3')) {
        & py -3 -m venv $venvFolder
    }
    elseif (Test-PythonVersion -CommandName 'python' -Prefix @()) {
        & python -m venv $venvFolder
    }
    else {
        Write-Host '未检测到 Python 3.10 或更高版本。请先安装 Python 并重新运行此脚本。' -ForegroundColor Red
        exit 1
    }
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $venvPython)) {
        Write-Host '虚拟环境创建失败，请检查 Python 安装以及文件夹写入权限。' -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host '检测到项目已有 .venv：将更新该环境中的依赖。' -ForegroundColor Cyan
}

Write-Host '将依赖安装到：' $venvFolder -ForegroundColor Cyan
& $venvPython -m pip --version
if ($LASTEXITCODE -ne 0) {
    Write-Host '虚拟环境中未发现 pip，正在恢复...' -ForegroundColor Yellow
    & $venvPython -m ensurepip --upgrade
    if ($LASTEXITCODE -ne 0) { exit 1 }
}
# 安装时直接选择镜像源，无须调整文件或系统 pip 配置。
# 默认阿里云；亦可选择其他镜像，或手动输入 HTTPS PyPI 索引网址。
# --isolated 忽略电脑已存在的用户级 pip 源配置，不修改系统和全局 pip 配置。
$sourceOptions = @{
    '1' = @{ Name = '阿里云（默认）'; Url = 'https://mirrors.aliyun.com/pypi/simple/' }
    '2' = @{ Name = 'PyPI 官方'; Url = 'https://pypi.org/simple/' }
    '3' = @{ Name = '中科大 USTC'; Url = 'https://mirrors.ustc.edu.cn/pypi/simple/' }
    '4' = @{ Name = '华为云'; Url = 'https://repo.huaweicloud.com/repository/pypi/simple/' }
    '5' = @{ Name = '清华 TUNA'; Url = 'https://pypi.tuna.tsinghua.edu.cn/simple/' }
}
Write-Host ''
Write-Host '请选择此次依赖安装使用的下载源（不改变系统 pip 设置）：' -ForegroundColor Cyan
foreach ($key in @('1','2','3','4','5')) {
    Write-Host ('  ' + $key + '. ' + $sourceOptions[$key].Name + '  ' + $sourceOptions[$key].Url)
}
Write-Host '  6. 自定义 HTTPS 索引地址'
$sourceChoice = (Read-Host '请输入编号 1-6，直接回车使用 1（阿里云）').Trim()
if ([string]::IsNullOrWhiteSpace($sourceChoice)) { $sourceChoice = '1' }
if ($sourceOptions.ContainsKey($sourceChoice)) {
    $mirrorIndex = $sourceOptions[$sourceChoice].Url
    $sourceName = $sourceOptions[$sourceChoice].Name
}
elseif ($sourceChoice -eq '6') {
    $mirrorIndex = (Read-Host '请输入完整 HTTPS PyPI 索引地址（通常以 /simple/ 结尾）').Trim()
    $validUri = $null
    $isValidUrl = [System.Uri]::TryCreate($mirrorIndex, [System.UriKind]::Absolute, [ref]$validUri)
    if ($isValidUrl) {
        $isValidUrl = ($validUri.Scheme -eq 'https' -and -not [string]::IsNullOrWhiteSpace($validUri.Host))
    }
    if (-not $isValidUrl) {
        Write-Host '自定义源必须为有效的 https://... 索引地址。未执行安装。' -ForegroundColor Red
        exit 1
    }
    $sourceName = '自定义源'
}
else {
    Write-Host '无效编号。请重新启动安装工具并选择 1-6。' -ForegroundColor Red
    exit 1
}
Write-Host ('当前安装源：' + $sourceName + '  ' + $mirrorIndex) -ForegroundColor Green
& $venvPython -m pip --isolated install --index-url $mirrorIndex --timeout 10 --retries 1 -r $requirements
if ($LASTEXITCODE -ne 0) {
    Write-Host '依赖安装失败，请检查网络或将错误信息截图。' -ForegroundColor Red
    exit 1
}

# 通过独立 Python 文件进行依赖检查，避免命令行引号兼容问题。
$checkScript = Join-Path $PSScriptRoot '检查独立环境.py'
if (-not (Test-Path -LiteralPath $checkScript)) {
    Write-Host '缺少 检查独立环境.py，请解压完整工具包。' -ForegroundColor Red
    exit 1
}
Write-Host '安装已结束，正在检查虚拟环境中的第三方依赖...' -ForegroundColor Cyan
& $venvPython -B $checkScript
if ($LASTEXITCODE -ne 0) {
    Write-Host '依赖检查失败，请查看上方错误提示。' -ForegroundColor Red
    exit 1
}
Write-Host ''
Write-Host '独立虚拟环境安装成功！以后请双击「双击启动工具.bat」运行。' -ForegroundColor Green
Write-Host '无需激活环境，也无需再次使用全局 pip 安装依赖。'
