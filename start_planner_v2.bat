@echo off
setlocal
cd /d "%~dp0"

if not exist backend\requirements.txt (
  echo ERRORE: file backend\requirements.txt non trovato.
  exit /b 1
)

where py >nul 2>nul
if %errorlevel%==0 (
  set "PYTHON_CMD=py -3"
) else (
  where python >nul 2>nul
  if errorlevel 1 (
    echo ERRORE: Python non trovato nel PATH.
    exit /b 1
  )
  set "PYTHON_CMD=python"
)

if not exist backend\.venv\Scripts\python.exe (
  echo Creo l'ambiente virtuale backend...
  call %PYTHON_CMD% -m venv backend\.venv
  if errorlevel 1 (
    echo ERRORE: creazione ambiente virtuale fallita.
    exit /b 1
  )
)

echo Aggiorno pip e dipendenze backend...
call backend\.venv\Scripts\python.exe -m pip install --upgrade pip
if errorlevel 1 (
  echo ERRORE: aggiornamento pip fallito.
  exit /b 1
)

call backend\.venv\Scripts\python.exe -m pip install -r backend\requirements.txt
if errorlevel 1 (
  echo ERRORE: installazione dipendenze backend fallita.
  exit /b 1
)

echo Avvio backend FastAPI...
start "Local Planner V2 - Backend" cmd /k "cd /d backend && ..\.venv\Scripts\python.exe -m uvicorn app.main:app --reload"
if errorlevel 1 (
  echo ERRORE: avvio backend fallito.
  exit /b 1
)

echo Avvio frontend...
start "Local Planner V2 - Frontend" cmd /k "%PYTHON_CMD% -m http.server 5500 --directory frontend"
if errorlevel 1 (
  echo ERRORE: avvio frontend fallito.
  exit /b 1
)

echo Attendo qualche secondo...
timeout /t 3 /nobreak > nul

echo Apro il planner...
start http://127.0.0.1:5500/index.html
endlocal