import sys
import os
import platform
import socket

def run_diagnostics():
    print("=====================================================")
    print("  CIS SMS Macro Environment Diagnostic Utility")
    print("=====================================================")
    print(f"Python Version: {sys.version}")
    print(f"Platform: {platform.system()} {platform.release()}")
    print(f"Hostname: {socket.gethostname()}")
    print("=====================================================")
    
    modules = ["pandas", "pyautogui", "pyperclip", "openpyxl", "xlwings"]
    missing = []
    
    print("\n[1] Checking Python Modules...")
    for mod in modules:
        try:
            __import__(mod)
            print(f"  - {mod}: OK")
        except ImportError:
            print(f"  - {mod}: MISSING ❌")
            missing.append(mod)
            
    if missing:
        print(f"\n⚠️  Missing libraries detected. Please install them using:")
        print(f"   pip install {' '.join(missing)}")
    else:
        print("  - All Python modules are successfully installed! ✅")
        
    print("\n[2] Checking Network Connection (for Discord Webhook)...")
    try:
        import urllib.request
        with urllib.request.urlopen("https://discord.com", timeout=3) as resp:
            print("  - Connection to Discord: OK ✅")
    except Exception as e:
        print(f"  - Connection to Discord failed: {e} ❌")
        
    print("\n[3] Checking OS Permissions...")
    if platform.system() == "Darwin":
        print("  - Operating System: macOS")
        print("  - IMPORTANT: Make sure to grant 'Accessibility' and 'Input Monitoring'")
        print("    permissions under System Settings > Privacy & Security.")
        print("    If PyAutoGUI fails to move/click, it is always a permission issue.")
    elif platform.system() == "Windows":
        print("  - Operating System: Windows")
        print("  - Note: Run Excel as Administrator if Excel write-back fails.")
        
    print("\n=====================================================")
    print("  Diagnostic Check Finished.")
    print("=====================================================")

if __name__ == "__main__":
    run_diagnostics()
