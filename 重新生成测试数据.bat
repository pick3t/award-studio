@echo off
chcp 65001 >nul
title 重新生成测试数据
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 没有找到 Python，请先安装 Python 3.10+。
    pause
    exit /b 1
)

echo 正在生成随机成绩表和奖状模板...
python tools\make_samples.py
if errorlevel 1 (
    echo 生成失败，请检查依赖是否装好。
    pause
    exit /b 1
)

echo.
echo 正在验证整条流程...
python tools\smoke_test.py
echo.
pause
