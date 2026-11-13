$ErrorActionPreference = 'Stop'
$target = Join-Path $PSScriptRoot '启动工具.ps1'
if (-not (Test-Path -LiteralPath $target)) {
    Write-Host '找不到启动工具.ps1。' -ForegroundColor Red
    exit 1
}
& $target
