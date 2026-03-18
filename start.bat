@echo off
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
echo Starting Construction Board on http://localhost:5555
python app.py
pause
