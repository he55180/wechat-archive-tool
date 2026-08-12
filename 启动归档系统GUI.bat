@echo off
chcp 936 > nul
cd /d "%~dp0"
title 微信公众号一键归档系统
echo [INFO] 正在启动 GUI 界面，请稍候...

REM 优先使用系统 Python 启动（系统环境内置了完整的 docx/requests 库）
python app.py
if %errorlevel% neq 0 (
    REM 回退尝试使用虚拟环境 Python
    venv\Scripts\python.exe app.py
)

pause
