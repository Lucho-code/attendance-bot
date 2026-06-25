@echo off
cd /d "C:\Users\Lucio\attendance-bot"
call venv\Scripts\activate
echo Panel web iniciado en http://localhost:8501
venv\Scripts\streamlit.exe run admin_panel.py --server.headless true --server.port 8501
