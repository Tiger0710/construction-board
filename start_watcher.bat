@echo off
chcp 65001 >nul
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
title 工事予定表 自動更新モニター
:loop
python watcher.py
echo.
echo [%date% %time%] watcher終了 - 10秒後に再起動...
timeout /t 10 /nobreak >nul
goto loop
