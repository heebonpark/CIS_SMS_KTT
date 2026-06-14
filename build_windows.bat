@echo off
title CIS SMS Macro Builder (Windows)
echo =====================================================
echo   CIS 문자보내기 Windows App Packaging Script
echo =====================================================
echo.
echo 1. Installing required Python dependencies...
python -m pip install --upgrade pip
pip install pyinstaller pandas pyautogui pyperclip openpyxl xlwings
echo.
echo 2. Packaging app.py with PyInstaller...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

pyinstaller --clean --onefile --noconsole --name "CIS_SMS_Macro" app.py
echo.
echo =====================================================
echo   3. Build Completed Successfully!
echo   Package executable file can be found at: dist\CIS_SMS_Macro.exe
echo =====================================================
pause
