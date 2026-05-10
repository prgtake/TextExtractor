@echo off
echo ===================================================
echo   TextExtractor (by.gemini) EXE Build Script
echo ===================================================
echo.

echo [1/3] Installing/Updating required libraries...
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas python-docx python-pptx extract-msg olefile google-genai openpyxl

echo.
echo [2/3] Cleaning up old build files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
timeout /t 2 > nul

echo.
echo [3/3] Starting PyInstaller...
:: TextExtractor を 1つのEXEファイル（コンソールなし）としてビルドします
:: Python 3.13 でエラーの原因となる pytest / py を除外します
python -m PyInstaller --onefile --noconsole --clean --name "TextExtractor" ^
 --exclude-module pytest ^
 --exclude-module py ^
 --hidden-import="pandas" ^
 --hidden-import="docx" ^
 --hidden-import="pptx" ^
 --hidden-import="extract_msg" ^
 --hidden-import="olefile" ^
 --hidden-import="PIL._tkinter_finder" ^
 textextractor.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] PyInstaller failed with error code %ERRORLEVEL%
) else (
    echo.
    echo ===================================================
    echo   Build Process Finished!
    echo   Check the 'dist' folder for TextExtractor.exe
    echo ===================================================
)
pause
