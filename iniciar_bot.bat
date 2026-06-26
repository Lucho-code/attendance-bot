@echo off
cd /d "C:\Users\Lucio\attendance-bot"
call venv\Scripts\activate

set FFMPEG_DIR=C:\Users\Lucio\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-8.1.1-full_build\bin
set PATH=%FFMPEG_DIR%;%PATH%

:loop
echo [%date% %time%] Iniciando bot...
py bot.py
echo [%date% %time%] Bot detenido. Reiniciando en 10 segundos...
timeout /t 10 /nobreak >nul
goto loop
