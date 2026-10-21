@echo off
chcp 65001 >nul
title Award Studio
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo [错误] 没有找到 Python。
    echo 请先安装 Python 3.10 或更高版本：https://www.python.org/downloads/
    echo 安装时记得勾选 "Add Python to PATH"。
    echo.
    pause
    exit /b 1
)

python -c "import customtkinter, docx, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo 首次运行，正在安装依赖，请稍候...
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo 依赖安装失败，请检查网络后重试。
        pause
        exit /b 1
    )
)

where pythonw >nul 2>nul
if errorlevel 1 (
    python app.py --demo
) else (
    start "" pythonw app.py --demo
)
exit /b
