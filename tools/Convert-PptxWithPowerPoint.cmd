@echo off
setlocal
powershell.exe -NoLogo -NoProfile -STA -File "%~dp0Convert-PptxWithPowerPoint.ps1" %*
set "pptx_html_exit=%ERRORLEVEL%"
if "%~1"=="" pause
exit /b %pptx_html_exit%
