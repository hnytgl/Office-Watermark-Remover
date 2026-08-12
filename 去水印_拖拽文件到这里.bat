@echo off
chcp 65001 >nul
setlocal
set "SCRIPT=%~dp0office_watermark_remover.py"
if "%~1"=="" (
  python "%SCRIPT%"
) else (
  python "%SCRIPT%" %*
)
echo.
pause
