@echo off
chcp 65001 >nul
cd /d "%~dp0"
python mahalla_bot.py --tanlash
pause
