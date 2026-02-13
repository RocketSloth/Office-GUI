@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %ERRORLEVEL%==0 (
    py TJTaskCenter.py
) else (
    python TJTaskCenter.py
)

endlocal
