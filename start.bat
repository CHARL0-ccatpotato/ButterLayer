@echo off
chcp 65001 >nul
cd /d "%~dp0"
for /f "tokens=5" %%P in ('netstat -ano ^| findstr "127.0.0.1:5000" ^| findstr "LISTENING"') do (
    taskkill /PID %%P /F >nul 2>nul
)
call venv\Scripts\activate
start "" http://127.0.0.1:5000
python app.py
pause
