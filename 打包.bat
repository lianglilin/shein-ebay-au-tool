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
    echo   [ERROR] 没找到 Python，先装 Python 3.10+ 再打包。
    pause
    exit /b 1
)

echo   用这个 Python 打包: "!PY!"
"!PY!" -m pip install --upgrade pyinstaller websocket-client Pillow
"!PY!" -m PyInstaller --noconfirm --onefile --windowed ^
    --name "SHEIN选品工具" ^
    --hidden-import websocket ^
    --hidden-import PIL --hidden-import PIL.Image --hidden-import PIL.ImageOps ^
    --collect-submodules PIL ^
    --exclude-module openpyxl --exclude-module numpy ^
    --exclude-module pandas --exclude-module matplotlib ^
    ebay_tool.py

echo.
if exist "dist\SHEIN选品工具.exe" (
    echo   打包完成: dist\SHEIN选品工具.exe
) else (
    echo   [ERROR] 打包失败，看上面的输出。
)
pause
