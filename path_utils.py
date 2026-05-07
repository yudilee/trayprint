import os
import sys
import platform


def get_root_dir():
    """Returns the root directory of the application.
    If the application is 'frozen' (standalone executable), returns the
    directory where the executable is located.
    Otherwise, returns the directory of the script.
    """
    if getattr(sys, 'frozen', False):
        # Bundled by PyInstaller
        return os.path.dirname(sys.executable)

    # Running as a normal script
    return os.path.dirname(os.path.abspath(__file__))


def get_data_dir():
    """Return a user-writable directory for data files (logs, DBs, config).

    Platform-specific paths (XDG/Apple/Windows conventions):
      - Windows: %LOCALAPPDATA%\\TrayPrint
      - macOS:   ~/Library/Application Support/TrayPrint
      - Linux:   ~/.local/share/TrayPrint  (XDG_DATA_HOME)
    """
    if getattr(sys, 'frozen', False):
        system = platform.system()
        if system == 'Windows':
            base = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
            data_dir = os.path.join(base, 'TrayPrint')
        elif system == 'Darwin':
            data_dir = os.path.join(
                os.path.expanduser('~'),
                'Library', 'Application Support', 'TrayPrint'
            )
        else:
            # Linux / BSD / other Unix — follow XDG Base Directory spec
            xdg_data = os.environ.get(
                'XDG_DATA_HOME',
                os.path.join(os.path.expanduser('~'), '.local', 'share')
            )
            data_dir = os.path.join(xdg_data, 'TrayPrint')
    else:
        # In development, use the project root
        data_dir = os.path.dirname(os.path.abspath(__file__))

    os.makedirs(data_dir, exist_ok=True)
    return data_dir
