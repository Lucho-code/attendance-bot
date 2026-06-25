@echo off
cd /d "C:\Users\Lucio\attendance-bot"
call venv\Scripts\activate

:loop
echo [%date% %time%] Iniciando panel movil...
venv\Scripts\python.exe mobile_panel.py
echo [%date% %time%] Panel movil detenido. Reiniciando en 5 segundos...
timeout /t 5 /nobreak >nul
goto loop
