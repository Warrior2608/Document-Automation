@echo off
REM Build the Document Automation Studio Windows executable using PyInstaller.
python -m pip install -e .[dev]
pyinstaller --noconfirm --onefile --windowed --name DocumentAutomationStudio document_automation_studio/app.py
echo Build complete. The executable is in the dist folder.
