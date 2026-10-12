@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m playwright install chromium
python namuna_excel.py
echo.
echo Ornatish tugadi.
pause
