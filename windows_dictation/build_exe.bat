@echo off
echo === Building Dictation Tool ===
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.10+ from python.org
    pause
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt
echo.

REM Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Build
echo Building executable...
pyinstaller --onefile --windowed --name DictationTool --icon=NONE dictation_app.py

if exist "dist\DictationTool.exe" (
    echo.
    echo === Build successful! ===
    echo.
    echo Executable: dist\DictationTool.exe
    echo.
    echo IMPORTANT: Create a .env file next to the .exe with:
    echo   OPENAI_API_KEY=sk-your-key-here
    echo.
    echo Then double-click DictationTool.exe to run.
) else (
    echo.
    echo ERROR: Build failed. Check errors above.
)

pause
