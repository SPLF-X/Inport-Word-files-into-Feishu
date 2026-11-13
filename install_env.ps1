$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot '首次安装独立环境.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Host '找不到首次安装独立环境.ps1。' -ForegroundColor Red
    exit 1
}
& $target
