@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Dasturning yangi versiyasi yuklanmoqda...
curl.exe -sSL -o mahalla_bot_yangi.py "https://raw.githubusercontent.com/xayrullo0522/-/claude/excel-to-website-auto-upload-wz4sgi/mahalla_bot.py"
if errorlevel 1 goto xato
findstr /c:"def asosiy" mahalla_bot_yangi.py >nul
if errorlevel 1 goto xato
if exist mahalla_bot.py copy /y mahalla_bot.py mahalla_bot_eski.py >nul
move /y mahalla_bot_yangi.py mahalla_bot.py >nul
echo.
echo Tayyor! mahalla_bot.py yangilandi (eskisi: mahalla_bot_eski.py).
echo config.yaml va Excel fayllarga tegilmadi.
goto oxiri
:xato
if exist mahalla_bot_yangi.py del mahalla_bot_yangi.py
echo.
echo Yuklab bo'lmadi. Internetni tekshiring va qaytadan urinib ko'ring.
:oxiri
pause
