@echo off
setlocal DisableDelayedExpansion
set "CLI=%~dp0convert-pptx.cjs"
set "PORTABLE_NODE=%~dp0..\node.exe"
if exist "%PORTABLE_NODE%" goto portable
where node >nul 2>nul
if errorlevel 1 goto missing
node "%CLI%" %*
exit /b %ERRORLEVEL%
:portable
"%PORTABLE_NODE%" "%CLI%" %*
exit /b %ERRORLEVEL%
:missing
echo Node.js 22.22.2+, 24.15.0+, or 26+ is required. 1>&2
exit /b 1
