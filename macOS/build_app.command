#!/bin/bash
cd "$(dirname "$0")"

echo "============================================"
echo "  Excel Sheet Splitter - Build macOS App"
echo "============================================"
echo

if ! command -v python3 &> /dev/null; then
    echo "[ERROR] python3 not found. Please install Python 3 first."
    echo "Download: https://www.python.org/downloads/macos/"
    read -n 1 -s -r -p "Press any key to exit..."
    exit 1
fi

echo "[1/3] Installing dependencies..."
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt pyinstaller
if [ $? -ne 0 ]; then
    echo "[ERROR] Dependency install failed. See messages above."
    echo "If you see a Tk-related error, try: brew install python-tk"
    read -n 1 -s -r -p "Press any key to exit..."
    exit 1
fi

echo
echo "[2/3] Building app (first build may take 1-2 minutes)..."
python3 -m PyInstaller --windowed --collect-all tkinterdnd2 --name "ExcelSheetSplitter" excel_sheet_splitter.py
if [ $? -ne 0 ]; then
    echo "[ERROR] Build failed. See messages above."
    read -n 1 -s -r -p "Press any key to exit..."
    exit 1
fi

echo
echo "[3/3] Done!"
echo "App location: $(pwd)/dist/ExcelSheetSplitter.app"
echo
echo "NOTE: the first time you open this app, macOS Gatekeeper will likely"
echo "block it ('cannot be opened because it is from an unidentified developer')."
echo "To allow it: right-click (or Control-click) the app in Finder -> Open ->"
echo "click 'Open' again in the dialog that appears. You only need to do this once."
echo
read -n 1 -s -r -p "Press any key to exit..."
