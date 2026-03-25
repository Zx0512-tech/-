@echo off
setlocal
cd /d %~dp0

set SCRIPT=D:\pyansys\Codex_pyansys\plot_time_history.py

where pythonw >nul 2>nul
if %errorlevel%==0 (
    start "" pythonw "%SCRIPT%"
) else (
    python "%SCRIPT%"
)

endlocal
