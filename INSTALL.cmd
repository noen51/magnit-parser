@echo off
setlocal
cd /d "%~dp0"

echo ========================================
echo MAGNIT RESET INSTALLER
echo ========================================
echo.

set "PY_CMD="
where py.exe >nul 2>nul
if not errorlevel 1 set "PY_CMD=py -3"

if not defined PY_CMD (
  where python.exe >nul 2>nul
  if not errorlevel 1 set "PY_CMD=python"
)

if not defined PY_CMD (
  echo ERROR: Python was not found.
  echo Install Python 3.11, 3.12 or 3.13.
  pause
  exit /b 1
)

%PY_CMD% install.py
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
  echo INSTALL OK.
  echo Now run RUN.cmd
) else (
  echo INSTALL FAILED.
  echo Send install_log.txt.
)

pause
exit /b %RC%
