@echo off
REM =========================================================
REM Force Automation - Windows build script
REM Run this on a Windows machine that has Python 3.10+ installed.
REM It creates a self-contained application in dist\Force_Automation\
REM =========================================================

echo.
echo === Step 1: Create/activate virtual environment ===
if not exist venv (
    python -m venv venv
)
call venv\Scripts\activate.bat

echo.
echo === Step 2: Install dependencies ===
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo === Step 3: Clean previous build ===
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo.
echo === Step 4: Build production executable (windowed, no console) ===
pyinstaller Force_Automation.spec

echo.
echo === Build complete ===
echo Application folder: dist\Force_Automation\
echo Run it with:         dist\Force_Automation\Force_Automation.exe
echo.
echo If it does not start correctly, build the debug version instead:
echo     pyinstaller Force_Automation_debug.spec
echo and run dist\Force_Automation_Debug\Force_Automation_Debug.exe
echo to see console output and errors.
echo.
pause
