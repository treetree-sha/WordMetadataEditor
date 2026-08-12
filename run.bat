@echo off
chcp 65001 > nul
title Word 文档属性高级编辑器
echo 正在启动 Word 文档属性高级编辑器...
cd /d "%~dp0"

if exist "D:\Programs\Anaconda\python.exe" (
    "D:\Programs\Anaconda\python.exe" main.py
) else (
    python main.py
)
