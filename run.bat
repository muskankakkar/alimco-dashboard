@echo off
chcp 65001 > nul
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1
echo Starting ALIMCO Dashboard...
python -u app.py > run_log.txt 2>&1
echo.
echo App stopped. Check run_log.txt for details.
type run_log.txt
pause
