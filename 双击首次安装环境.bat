@echo off
setlocal
chcp 65001 >nul
pushd "%~dp0"
if errorlevel 1 goto bad_directory
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_env.ps1"
set "result=%ERRORLEVEL%"
popd
if not "%result%"=="0" echo 【操作提示】运行环境准备未完成，请查看上方提示。
echo.
echo 【需要按键】请按任意键关闭窗口。
pause >nul
exit /b %result%
:bad_directory
echo 【操作提示】无法打开工具所在文件夹，请检查文件夹是否可访问。
echo 【需要按键】请按任意键关闭窗口。
pause >nul
exit /b 1
