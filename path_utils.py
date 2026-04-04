import os
import sys

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
