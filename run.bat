@echo off
chcp 65001 > nul
title Word 文档属性高级编辑器
echo 正在启动 Word 文档属性高级编辑器...
cd /d "%~dp0"
python main.py
if %errorlevel% neq 0 (
    echo.
    echo 程序的运行遇到错误，退出代码: %errorlevel%
    pause
)
