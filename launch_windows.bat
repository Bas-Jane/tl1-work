@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if %errorlevel%==0 (
    py -3 launch.py %*
) else (
    python launch.py %*
)

if errorlevel 1 (
    echo.
    echo Tools TL1 Generator did not start successfully.
    echo Review the message above, then press any key to close.
    pause >nul
)
endlocal
