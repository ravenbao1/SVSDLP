import os
import sys

# Add the local PyInstaller directory to the Python path
local_pyinstaller_path = os.path.join(os.path.dirname(__file__), 'local_pyinstaller')
sys.path.insert(0, local_pyinstaller_path)

# Import and run PyInstaller
import PyInstaller.__main__

PyInstaller.__main__.run([
    '--onefile',
    '--noconsole',
    'AppControl-RE.py'
])
