@echo off
setlocal

cd /d "%~dp0"

set "PYTHON_BIN=py"
where %PYTHON_BIN% >nul 2>nul
if errorlevel 1 (
  set "PYTHON_BIN=python"
)

where %PYTHON_BIN% >nul 2>nul
if errorlevel 1 (
  echo Error: Python launcher not found. Install Python 3.10+ first.
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  %PYTHON_BIN% -m venv .venv
  if errorlevel 1 exit /b 1
)

.venv\Scripts\python.exe -m pip install --upgrade pip >nul
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

.venv\Scripts\python.exe app.py
endlocal
