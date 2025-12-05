"""
Build standalone Windows executable.
Run: python build_exe.py
"""
import subprocess
import sys
import os
import shutil
import time

VERSION = "1.0.0"

def kill_running_exe():
    """Try to kill any running instances of the exe."""
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "LabelPrinter.exe"],
            capture_output=True,
            check=False
        )
        time.sleep(1)  # Give it a moment
    except Exception:
        pass

def clean_build_artifacts():
    """Remove old build artifacts."""
    exe_path = os.path.join("dist", "LabelPrinter.exe")
    
    # Try to remove existing exe
    if os.path.exists(exe_path):
        try:
            os.remove(exe_path)
            print("✓ Removed old executable")
        except PermissionError:
            print("⚠ Executable is in use, attempting to close it...")
            kill_running_exe()
            try:
                os.remove(exe_path)
                print("✓ Removed old executable")
            except PermissionError:
                print("❌ ERROR: Cannot remove LabelPrinter.exe")
                print("   Please close the application and try again.")
                return False
    
    # Clean build folder
    if os.path.exists("build"):
        try:
            shutil.rmtree("build")
            print("✓ Cleaned build folder")
        except Exception as e:
            print(f"⚠ Could not clean build folder: {e}")
    
    # Clean spec file
    if os.path.exists("LabelPrinter.spec"):
        try:
            os.remove("LabelPrinter.spec")
        except Exception:
            pass
    
    return True

def build():
    """Build the standalone executable."""
    
    print(f"🏗️  Building Label Printer v{VERSION}")
    print("=" * 50)
    
    # Clean old artifacts first
    if not clean_build_artifacts():
        return 1
    
    # Check for icon file
    icon_arg = "--icon=NONE"
    if os.path.exists("icon.ico"):
        icon_arg = "--icon=icon.ico"
        print("✓ Using custom icon: icon.ico")
    else:
        print("ℹ No icon.ico found, using default")
    
    # PyInstaller command
    cmd = [
        sys.executable, "-m", "PyInstaller",
        f"--name=LabelPrinter",
        "--onefile",
        "--windowed",
        icon_arg,
        "--add-data=bt_printer.py;.",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=PIL.ImageTk",
        "--hidden-import=PIL.ImageOps",
        "--hidden-import=fitz",
        "--hidden-import=customtkinter",
        "--hidden-import=serial",
        "--hidden-import=serial.tools.list_ports",
        "--collect-data=customtkinter",
        "--clean",
        "--noconfirm",
        "label_printer_app.py"
    ]
    
    print("\n📦 Running PyInstaller...")
    print("-" * 50)
    
    result = subprocess.run(cmd)
    
    print("-" * 50)
    
    if result.returncode == 0:
        print("\n✅ Build successful!")
        print(f"   📁 Output: dist/LabelPrinter.exe")
        
        # Get file size
        exe_path = os.path.join("dist", "LabelPrinter.exe")
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print(f"   📊 Size: {size_mb:.1f} MB")
        
        print("\n🎉 Ready to distribute!")
    else:
        print("\n❌ Build failed")
        return 1
    
    return 0


if __name__ == "__main__":
    # Install PyInstaller if needed
    try:
        import PyInstaller
    except ImportError:
        print("📥 Installing PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
    
    sys.exit(build())
