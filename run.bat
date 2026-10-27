@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal enabledelayedexpansion

set "PY="
for %%D in (Python313 Python312 Python311 Python310) do (
    if not defined PY (
        if exist "%LOCALAPPDATA%\Programs\Python\%%D\python.exe" (
            set "PY=%LOCALAPPDATA%\Programs\Python\%%D\python.exe"
        )
    )
)

if not defined PY (
    for /f "delims=" %%P in ('where python 2^>nul') do (
        if not defined PY set "PY=%%P"
    )
)

if not defined PY (
    echo.
    echo   [ERROR] Python not found on this computer.
    echo   Please install Python 3.10+ and run this file again.
    echo.
    pause
    exit /b 1
)

echo   Using Python: "!PY!"
echo.
"!PY!" "%~dp0ebay_tool.py"

if errorlevel 1 (
    echo.
    echo   [ERROR] The program exited with an error. See messages above.
    pause
)
