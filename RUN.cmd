@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo ERROR: Installation was not completed.
  echo Run INSTALL.cmd first.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" app.py
set "RC=%ERRORLEVEL%"

echo.
if not "%RC%"=="0" (
  echo APP ERROR. Send app_error.txt if it exists.
)
pause
exit /b %RC%
