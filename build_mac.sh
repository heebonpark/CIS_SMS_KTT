#!/bin/bash
# =====================================================================
#  CIS 문자보내기 macOS App Packaging Script
# =====================================================================
#
# Usage:
#   chmod +x build_mac.sh
#   ./build_mac.sh
#

echo "====================================================="
echo "  1. Python Dependency Check & Install"
echo "====================================================="
pip3 install --upgrade pip
pip3 install pyinstaller pandas pyautogui pyperclip openpyxl xlwings

echo ""
echo "====================================================="
echo "  2. Packaging app.py with PyInstaller"
echo "====================================================="
# Clean old build files
rm -rf build dist

pyinstaller --clean \
            --onefile \
            --windowed \
            --name "CIS_SMS_Macro" \
            app.py

echo ""
echo "====================================================="
echo "  3. Packaging Complete!"
echo "====================================================="
echo "  Stand-alone app generated at: dist/CIS_SMS_Macro.app"
echo "  Executable file generated at: dist/CIS_SMS_Macro"
echo ""
echo "  * Note: When running the app on macOS, you must grant"
echo "    Accessibility & Input Monitoring permissions in"
echo "    System Settings > Privacy & Security."
echo "====================================================="
