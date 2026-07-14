@echo off
setlocal
title Run8 DLC Manager - build single EXE

rem Works from either the repository root (source beside this file) or the
rem local development root (source in .\github).
set "APP=%~dp0"
if not exist "%APP%Run8DLCManager.pyw" set "APP=%~dp0github\"
if not exist "%APP%Run8DLCManager.pyw" (
    echo Run8DLCManager.pyw was not found beside this script or in .\github.
    pause
    exit /b 1
)
cd /d "%APP%"

set "PY=py -3"
%PY% --version >nul 2>&1 || set "PY=python"
%PY% --version >nul 2>&1 || (
    echo Python 3 not found - install it from python.org first.
    pause
    exit /b 1
)

echo Checking the source...
%PY% -W error::SyntaxWarning -m py_compile Run8DLCManager.pyw
if errorlevel 1 (
    echo Source check failed - EXE was not built.
    pause
    exit /b 1
)

echo Installing PyInstaller if needed...
%PY% -m pip show pyinstaller >nul 2>&1 || ^
    %PY% -m pip install --user pyinstaller
if errorlevel 1 (
    echo PyInstaller could not be installed.
    pause
    exit /b 1
)

echo Extracting the app icon...
%PY% Run8DLCManager.pyw emit-assets >nul
if errorlevel 1 (
    echo Asset extraction failed.
    pause
    exit /b 1
)

echo Building...
%PY% -m PyInstaller --onefile --windowed --clean --name Run8DLCManager ^
    --icon data\run8dlc.ico Run8DLCManager.pyw
if errorlevel 1 (
    echo Build failed - see messages above.
    pause
    exit /b 1
)

echo.
echo Done: %APP%dist\Run8DLCManager.exe
echo Attach that EXE and Run8DLCManager.pyw to the matching GitHub Release.
echo Note: some antivirus software is suspicious of freshly built
echo PyInstaller EXEs; offering the .pyw alongside it is deliberate.
echo.
pause
