$ErrorActionPreference = 'Stop'
$python = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
$checker = Join-Path $PSScriptRoot '检查配置和文件名.py'
if (-not (Test-Path -LiteralPath $python)) {
    Write-Host '未找到 .venv 中的 Python。请先安装独立环境。' -ForegroundColor Red
    exit 1
}
if (-not (Test-Path -LiteralPath $checker)) {
    Write-Host '找不到配置检查脚本。请确认已完整解压工具包。' -ForegroundColor Red
    exit 1
}
& $python -B $checker
if ($LASTEXITCODE -ne 0) { exit 1 }
exit 0
