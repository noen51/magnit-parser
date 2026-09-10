@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo ERROR: Run INSTALL.cmd first.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" app.py --test
set "RC=%ERRORLEVEL%"
pause
exit /b %RC%
