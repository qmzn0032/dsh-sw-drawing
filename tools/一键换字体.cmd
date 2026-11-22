@echo off
rem ===========================================================
rem   One-click font swap for SOLIDWORKS drawings
rem   Drag .SLDDRW files onto this file, or run it with no
rem   argument to use the drawing that is active in SW.
rem ===========================================================
chcp 65001 >nul
setlocal
set "HERE=%~dp0"
set "PY="

if defined SW_BRIDGE_PYTHON set "PY=%SW_BRIDGE_PYTHON%"
if not defined PY if exist "%LOCALAPPDATA%\hermes\hermes-agent\venv\Scripts\python.exe" set "PY=%LOCALAPPDATA%\hermes\hermes-agent\venv\Scripts\python.exe"
if not defined PY (
    where py >nul 2>nul && set "PY=py -3"
)
if not defined PY (
    where python >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo [x] Python not found. Install Python 3 or set SW_BRIDGE_PYTHON.
    pause
    exit /b 2
)

%PY% "%HERE%font_switch.py" %*
set "RC=%ERRORLEVEL%"
echo.
pause
exit /b %RC%
