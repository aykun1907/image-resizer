# build_script.py
import PyInstaller.__main__
import os

# Get the directory containing this script
script_dir = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    'image_resizer.py',  # your main script
    '--onefile',  # create a single executable
    '--windowed',  # don't show console window
    '--name=ImageResizer',  # name of your executable
    '--icon=app_icon.ico',  # path to your icon file
    '--add-data=app_icon.ico;.',  # include the icon file in the exe
    '--uac-admin',  # request admin rights when needed
    '--clean',  # clean cache before building
])