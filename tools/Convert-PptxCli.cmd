@echo off
setlocal
set "CLI=%~dp0convert-pptx.py"
set "VENV_PY=%~dp0..\.venv\Scripts\python.exe"
if exist "%VENV_PY%" goto venv
where py >nul 2>nul
if not errorlevel 1 goto py
where python >nul 2>nul
if not errorlevel 1 goto python
echo Python 3.10 or newer is required. 1>&2
exit /b 1
:venv
"%VENV_PY%" -X utf8 "%CLI%" %*
exit /b %ERRORLEVEL%
:py
py -3 -X utf8 "%CLI%" %*
exit /b %ERRORLEVEL%
:python
python -X utf8 "%CLI%" %*
exit /b %ERRORLEVEL%
