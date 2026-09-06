import os
import sys
import subprocess
import shutil

# ── Helper: generate a PNG icon from trayprint.ico (Linux) ──

def _generate_png_icon(png_path, size=256):
    """
    Generate a PNG icon from the existing trayprint.ico using Pillow.
    Falls back to creating a simple programmatic icon if Pillow is unavailable
    or the .ico file doesn't exist.
    """
    from PIL import Image, ImageDraw

    ico_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'trayprint.ico')

    # Try to load the .ico and convert to PNG
    if os.path.exists(ico_path):
        try:
            img = Image.open(ico_path)
            # Pick the largest available size
            if hasattr(img, 'n_frames') and img.n_frames > 1:
                best = max(range(img.n_frames),
                           key=lambda i: img.thumbnail((size, size), Image.LANCZOS) or
                           img if False else 0)
                img.seek(best)
            img = img.convert('RGBA')
            img.thumbnail((size, size), Image.LANCZOS)
            img.save(png_path, 'PNG')
            print(f"✓ Generated PNG icon from trayprint.ico → {png_path}")
            return True
        except Exception as e:
            print(f"! Could not convert .ico to PNG: {e}")
            # Fall through to programmatic icon

    # Programmatic fallback: draw a simple printer icon
    print("! Creating programmatic printer icon (no trayprint.ico)")
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Printer body
    margin = size // 8
    body_left = margin
    body_top = size // 4
    body_right = size - margin
    body_bottom = size - margin
    body_height = body_bottom - body_top
    body_width = body_right - body_left

    # Draw printer body (rounded rect approximation)
    draw.rectangle([body_left, body_top + body_height // 3, body_right, body_bottom],
                   fill=(50, 150, 250), outline=(30, 100, 200))

    # Draw paper slot (top)
    slot_top = body_top + body_height // 6
    slot_bottom = body_top + body_height // 3
    draw.rectangle([body_left + body_width // 4, slot_top,
                    body_right - body_width // 4, slot_bottom],
                   fill=(255, 255, 255), outline=(200, 200, 200))

    # Draw paper coming out
    paper_top = body_bottom - body_height // 4
    draw.rectangle([body_left + body_width // 4, paper_top,
                    body_right - body_width // 4, body_bottom],
                   fill=(255, 255, 255), outline=(200, 200, 200))

    # Lines on paper
    line_y = paper_top + (body_bottom - paper_top) // 3
    for i in range(3):
        draw.line([body_left + body_width // 3, line_y + i * 8,
                   body_right - body_width // 3, line_y + i * 8],
                  fill=(100, 100, 100), width=2)

    img.save(png_path, 'PNG')
    print(f"✓ Created programmatic PNG icon → {png_path}")
    return True


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

    # On Linux, ensure Pillow is available for icon generation
    if sys.platform.startswith('linux'):
        try:
            import PIL
        except ImportError:
            print("[INFO] Pillow not found — installing for icon generation...")
            subprocess.run([sys.executable, '-m', 'pip', 'install', 'Pillow'], check=True)

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
    elif sys.platform.startswith('linux'):
        # Linux: PNG icon — generate from .ico or create programmatic icon
        png_path = os.path.join(app_dir, 'installer', 'trayprint.png')
        if not os.path.exists(png_path):
            os.makedirs(os.path.dirname(png_path), exist_ok=True)
            _generate_png_icon(png_path, size=256)
        if os.path.exists(png_path):
            icon_opt = ['--icon', png_path]
            print(f"✓ Using PNG icon: {png_path}")
        else:
            print("! No PNG icon available — will use default")

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

    # Base PyInstaller flags (common to all platforms)
    pyinstaller_flags = [
        '--onefile',
        '--name', 'trayprint',
        '--clean',
        '--collect-all=charset_normalizer',
        '--collect-all=chardet',
    ]

    # --windowed is Windows/macOS only; on Linux it suppresses the console
    # which is not the intended behavior (we want stderr visible for diagnostics)
    if sys.platform in ('win32', 'darwin'):
        pyinstaller_flags.append('--windowed')

    cmd = [
        sys.executable, '-m', 'PyInstaller',
    ] + pyinstaller_flags + icon_opt + hidden_imports + datas + binaries + [
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
        # On Linux with --onefile, PyInstaller may place it as 'trayprint' in dist/
        alt_path = os.path.join(dist_dir, 'trayprint')
        if os.path.exists(alt_path):
            size_mb = os.path.getsize(alt_path) / (1024 * 1024)
            print(f"\n{'=' * 50}")
            print(f"  BUILD SUCCESS!")
            print(f"  Output: {alt_path}")
            print(f"  Size:   {size_mb:.1f} MB")
            print(f"{'=' * 50}")
        else:
            print("\n✗ Build failed — executable not found.")


if __name__ == '__main__':
    build()
