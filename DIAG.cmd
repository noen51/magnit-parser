@echo off
setlocal
cd /d "%~dp0"

>diagnostics.txt echo MAGNIT RESET DIAGNOSTICS
>>diagnostics.txt echo Folder: %CD%
>>diagnostics.txt echo Date: %DATE% %TIME%
>>diagnostics.txt echo.

where py.exe >>diagnostics.txt 2>&1
py -3 --version >>diagnostics.txt 2>&1
where python.exe >>diagnostics.txt 2>&1
python --version >>diagnostics.txt 2>&1
>>diagnostics.txt echo.
dir /b >>diagnostics.txt 2>&1

echo diagnostics.txt created.
pause
