@echo off
cd /d "%~dp0"

echo Avvio backend FastAPI...
start "Local Planner V2 - Backend" cmd /k "cd /d backend && .venv\Scripts\activate && uvicorn app.main:app --reload"

echo Avvio frontend...
start "Local Planner V2 - Frontend" cmd /k "cd /d frontend && python -m http.server 5500"

echo Attendo qualche secondo...
timeout /t 3 /nobreak > nul

echo Apro il planner...
start http://127.0.0.1:5500/index.html