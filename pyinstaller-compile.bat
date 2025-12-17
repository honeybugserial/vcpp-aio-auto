@echo off
setlocal

rem === CONFIG ===
set PY=python
set OUTSCRIPT=vcpp-redist-downloader.py
set ICON=vcpp-aio-auto.ico

rem === Echo Config Vars ===
echo Python:     %PY%
echo INSCRIPT:  %INSCRIPT%
echo OUTSCRIPT: %OUTSCRIPT%
echo OBFUS:     %OBFUS%
echo ICON:      %ICON%
pause
echo.



rem === CLEAN PREVIOUS PYINSTALLER OUTPUT ===
echo Cleaning previous build...
if exist build rd /s /q build
if exist dist  rd /s /q dist
if exist __pycache__ rd /s /q __pycache__
del /f /q "*.spec" 2>nul

rem === PYINSTALLER BUILD ===
echo Building executable...
%PY% -m PyInstaller ^
    --onefile ^
    --uac-admin ^
    --icon "%ICON%" ^
    --hidden-import rich.console ^
    --hidden-import rich.panel ^
    --hidden-import rich.prompt ^
    --hidden-import rich.rule ^
    "%OUTSCRIPT%"

if errorlevel 1 (
    echo PyInstaller build failed.
    exit /b 1
)

echo.
echo Build complete.
echo Output EXE: dist\%~nOUTSCRIPT%.exe
pause
