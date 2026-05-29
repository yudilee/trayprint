import os
import sys
import subprocess
import shutil

def build():
    """Build the tray app into a standalone executable using PyInstaller."""
    print("=" * 50)
    print("  Trayprint - Build Script")
    print("=" * 50)

    # Check PyInstaller
    try:
        import PyInstaller
        print("[OK] PyInstaller %s found" % PyInstaller.__version__)
    except ImportError:
        print("[FAIL] PyInstaller not found. Installing...")
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'pyinstaller'], check=True)

    # Check pywin32 (critical for Windows)
    if sys.platform == 'win32':
        try:
            import win32print
            print("[OK] pywin32 (win32print) found")
        except ImportError:
            print("[FAIL] pywin32 not found. Installing...")
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'pywin32'], check=True)
            print("! NOTE: If build still fails with 'win32print not found', run this manually on Windows:")
            print("  python Scripts/pywin32_postinstall.py -install")

    app_dir = os.path.dirname(os.path.abspath(__file__))
    dist_dir = os.path.join(app_dir, 'dist')

    # Determine platform-specific options
    icon_opt = []
    hidden_imports = [
        # PySide6
        '--hidden-import=PySide6.QtCore',
        '--hidden-import=PySide6.QtGui',
        '--hidden-import=PySide6.QtWidgets',
        # Flask & web
        '--hidden-import=flask',
        '--hidden-import=requests',
        '--hidden-import=websockets',
        '--hidden-import=websockets.asyncio.client',
        # Image / PDF
        '--hidden-import=PIL',
        '--hidden-import=PIL.Image',
        '--hidden-import=fitz',
        # Application modules
        '--hidden-import=server',
        '--hidden-import=ui_settings',
        '--hidden-import=autostart',
        '--hidden-import=printer',
        '--hidden-import=path_utils',
        '--hidden-import=logger',
        '--hidden-import=log_utils',
        '--hidden-import=websocket_client',
        '--hidden-import=notification_history',
        '--hidden-import=queue_dialog',
        '--hidden-import=diagnostics_dialog',
        '--hidden-import=updater',
        '--hidden-import=capabilities',
        '--hidden-import=theme',
        # Standard library modules used dynamically
        '--hidden-import=difflib',
        '--hidden-import=html',
        '--hidden-import=http',
        '--hidden-import=email',
        '--hidden-import=encodings',
        '--hidden-import=ctypes',
        '--hidden-import=sqlite3',
        '--hidden-import=queue',
        '--hidden-import=asyncio',
        '--hidden-import=winreg',
        '--hidden-import=pathlib',
        '--hidden-import=hashlib',
        '--hidden-import=platform',
        '--hidden-import=socket',
        '--hidden-import=re',
        '--hidden-import=base64',
        '--hidden-import=uuid',
        '--hidden-import=argparse',
        '--hidden-import=webbrowser',
        '--hidden-import=dataclasses',
        '--hidden-import=typing',
        '--hidden-import=contextlib',
        '--hidden-import=collections',
    ]
    
    if sys.platform == 'win32':
        icon_path = os.path.join(app_dir, 'trayprint.ico')
        if os.path.exists(icon_path):
            icon_opt = ['--icon', icon_path]
        else:
            print("! Windows .ico icon not found at %s — will use default" % icon_path)
        hidden_imports += [
            '--hidden-import=win32print',
            '--hidden-import=win32con',
            '--hidden-import=win32api',
            '--hidden-import=win32ui',
            '--hidden-import=win32timezone',
            '--hidden-import=pywintypes',
            '--hidden-import=pythoncom',
        ]
    elif sys.platform == 'darwin':
        # macOS: .icns icon and macOS-specific hidden imports
        icns_path = os.path.join(app_dir, 'trayprint.icns')
        if os.path.exists(icns_path):
            icon_opt = ['--icon', icns_path]
        else:
            print("! macOS .icns icon not found at %s — will use default", icns_path)
        hidden_imports += [
            '--hidden-import=platform_darwin',
            '--hidden-import=objc',
            '--hidden-import=CoreFoundation',
            '--hidden-import=CoreGraphics',
        ]

    # Data files to include inside the bundle (read-only templates/icons)
    datas = [
        f'--add-data=config.json{os.pathsep}.',
        f'--add-data=templates{os.pathsep}templates',
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
        '--collect-all=charset_normalizer',
        '--collect-all=chardet',
    ] + icon_opt + hidden_imports + datas + binaries + [
        'app.py'
    ]

    print(f"\nRunning: {' '.join(cmd)}\n")
    subprocess.run(cmd, cwd=app_dir, check=True)

    # Output location
    if sys.platform == 'win32':
        exe_name = 'trayprint.exe'
    elif sys.platform == 'darwin':
        exe_name = 'trayprint.app'  # macOS creates a .app bundle
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
