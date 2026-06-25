@echo off
echo ============================================
echo   Excel Sheet Splitter - Build EXE Script
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3 first.
    echo Download: https://www.python.org/downloads/
    echo IMPORTANT: check "Add python.exe to PATH" during install.
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo [ERROR] Dependency install failed. See messages above.
    pause
    exit /b 1
)

echo.
echo [2/3] Building exe (first build may take 1-2 minutes)...
python -m PyInstaller --onefile --windowed --collect-all tkinterdnd2 --name ExcelSheetSplitter excel_sheet_splitter.py
if errorlevel 1 (
    echo [ERROR] Build failed. See messages above.
    pause
    exit /b 1
)

echo.
echo [3/3] Done!
echo EXE file location: %cd%\dist\ExcelSheetSplitter.exe
echo.
pause
