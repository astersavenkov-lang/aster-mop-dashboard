@echo off
title ADM-ASTER Report Bot
cd /d "%~dp0"
echo ═══════════════════════════════════
echo   ADM-ASTER Report Bot
echo   %date% %time%
echo ═══════════════════════════════════
echo.
python aster_bot.py
echo.
echo Бот остановлен. Нажмите любую клавишу...
pause > nul
