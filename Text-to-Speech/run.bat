@echo off
echo Starting Text-to-Speech application...

REM Get the directory where this script is located
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Check if Python script exists
if not exist "Text-to-Speech.py" (
    echo Error: Text-to-Speech.py not found in current directory
    echo Current directory: %SCRIPT_DIR%
    pause
    exit /b 1
)

REM Set FFmpeg path
set "FFMPEG_PATH=%SCRIPT_DIR%\ffmpeg"
set "PATH=%PATH%;%FFMPEG_PATH%"

REM Check if virtual environment exists
if not exist ".venv\Scripts\activate.bat" (
    echo Error: Virtual environment not found
    echo Please run setup.bat first
    pause
    exit /b 1
)

REM Activate virtual environment and run the script
echo Activating virtual environment...
call ".venv\Scripts\activate.bat"

REM Run the Python script with full path and capture output
echo Running application...
python "%SCRIPT_DIR%Text-to-Speech.py" 2>&1
if errorlevel 1 (
    echo.
    echo Python script failed with error code %errorlevel%
    echo Please check if all dependencies are installed
    echo Try running setup.bat again
    echo.
    echo If the error persists, please check:
    echo 1. All required packages are installed
    echo 2. Tesseract-OCR is properly set up
    echo 3. No other Python processes are running
    pause
    exit /b 1
)

pause 