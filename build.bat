@echo off
chcp 65001 >nul 2>&1
title Build Tool
echo ============================================
echo   Build EXE
echo ============================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Please install Python 3.8+
    echo https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
python -m pip install openpyxl pyinstaller -q
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo [2/3] Building EXE...
python -m PyInstaller --onefile --noconsole --name OrderMatchTool --clean app.py
if %errorlevel% neq 0 (
    echo [ERROR] Build failed.
    pause
    exit /b 1
)

echo [3/3] Done!
echo.
echo EXE file: dist\OrderMatchTool.exe
echo.
echo Copy OrderMatchTool.exe to your shared folder.
echo It will create data.db automatically on first run.
echo.
pause
