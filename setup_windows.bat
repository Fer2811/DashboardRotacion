@echo off
cd /d %~dp0
py -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if not exist .env copy .env.example .env >nul
echo Listo. Completa .env y ejecuta python run_dashboard.py
pause
