import os
import sys
import subprocess
import shutil

def build():
    """Build the tray app into a standalone executable using PyInstaller."""
    print("=" * 50)
    print("  Trayprint — Build Script")
    print("=" * 50)

    # Check PyInstaller
    try:
        import PyInstaller
        print(f"✓ PyInstaller {PyInstaller.__version__} found")
    except ImportError:
        print("✗ PyInstaller not found. Installing...")
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'pyinstaller'], check=True)

    app_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(app_dir, 'dist')

    # Determine platform-specific options
    icon_opt = []
    hidden_imports = [
        '--hidden-import=PySide6.QtCore',
        '--hidden-import=PySide6.QtGui',
        '--hidden-import=PySide6.QtWidgets',
        '--hidden-import=flask',
        '--hidden-import=requests',
        '--hidden-import=PIL',
        '--hidden-import=win32print',
        '--hidden-import=win32con',
        '--hidden-import=win32api',
        '--hidden-import=win32ui',
        '--hidden-import=fitz',
    ]

    # Data files to include inside the bundle (read-only templates/icons)
    datas = [
        f'--add-data=config.json{os.pathsep}.',
    ]
    
    binaries = []

    # On Windows, include SumatraPDF if present
    sumatra_path = os.path.join(app_dir, 'SumatraPDF.exe')
    if os.path.exists(sumatra_path):
        datas.append(f'--add-data=SumatraPDF.exe{os.pathsep}.')
        print("✓ SumatraPDF.exe will be bundled")

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',
        '--windowed',
        '--name', 'trayprint',
        '--clean',
    ] + icon_opt + hidden_imports + datas + binaries + [
        'app.py'
    ]

    print(f"\nRunning: {' '.join(cmd)}\n")
    subprocess.run(cmd, cwd=app_dir, check=True)

    # Output location
    if sys.platform == 'win32':
        exe_name = 'trayprint.exe'
    else:
        exe_name = 'trayprint'

    exe_path = os.path.join(dist_dir, exe_name)
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / (1024 * 1024)
        print(f"\n{'=' * 50}")
        print(f"  BUILD SUCCESS!")
        print(f"  Output: {exe_path}")
        print(f"  Size:   {size_mb:.1f} MB")
        print(f"{'=' * 50}")
    else:
        print("\n✗ Build failed — executable not found.")


if __name__ == '__main__':
    build()
