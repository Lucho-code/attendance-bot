@echo off
cd /d "C:\Users\Lucio\attendance-bot"
call venv\Scripts\activate

:loop
echo [%date% %time%] Iniciando bot...
py bot.py
echo [%date% %time%] Bot detenido. Reiniciando en 10 segundos...
timeout /t 10 /nobreak >nul
goto loop
