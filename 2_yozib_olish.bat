@echo off
chcp 65001 >nul
cd /d "%~dp0"
set /p Q="Excel nechanchi qatoridagi fuqaroni qolda kiritasiz? (masalan 2): "
python mahalla_bot.py --yozib_olish %Q%
pause
