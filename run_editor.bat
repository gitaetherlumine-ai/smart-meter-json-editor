@echo off
title Smart Meter JSON Editor
cd /d "%~dp0"
echo ========================================================
echo Starting Smart Meter JSON Editor...
echo ========================================================

if exist "runtime\Scripts\python.exe" (
    "runtime\Scripts\python.exe" web_editor.py
    goto end
)

where python >nul 2>nul
if %ERRORLEVEL% equ 0 (
    python web_editor.py
    goto end
)

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" web_editor.py
    goto end
)

echo [ERROR] Python runtime not found.
pause

:end
