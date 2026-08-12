@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo [1/3] 安装运行依赖...
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [2/3] 安装 PyInstaller...
python -m pip install pyinstaller
if errorlevel 1 goto :error

echo [3/3] 生成 Windows EXE...
python -m PyInstaller --noconfirm --clean --onefile --windowed --name OfficeWatermarkRemover --collect-all tkinterdnd2 office_watermark_gui.py
if errorlevel 1 goto :error

echo.
echo 构建完成：dist\OfficeWatermarkRemover.exe
pause
exit /b 0

:error
echo.
echo 构建失败，请检查上方错误信息。
pause
exit /b 1
