@echo off
rem ===========================================================
rem  sw_bridge MCP server launcher (stdio).
rem  Finds a Python 3 with pywin32, then runs mcp_sw_server.py.
rem  Order: SW_BRIDGE_PYTHON -> hermes venv (this machine) -> py -3 -> python.
rem  Keep stdout clean: anything printed here would corrupt JSON-RPC.
rem ===========================================================
setlocal
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
    echo [sw-bridge] Python not found. Set SW_BRIDGE_PYTHON to a python.exe with pywin32. 1>&2
    exit /b 2
)

%PY% "%~dp0mcp_sw_server.py"
exit /b %ERRORLEVEL%
