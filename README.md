# Document Automation Studio

Document Automation Studio is a Windows desktop application for automating Word and Excel document processing.

## Phase 1 completed

This phase establishes the architectural foundation:

- Clean package layout
- Configuration management with JSON persistence
- Structured logging
- Initial unit tests

## Phase 2 completed

This phase adds the first GUI shell for the desktop workflow:

- Main application window
- Folder selection controls
- Rule and logo selection inputs
- Progress and log panels
- Theme toggle

## Project structure

- document_automation_studio/config - configuration management
- document_automation_studio/models - data models
- document_automation_studio/utils - shared utilities
- tests - unit tests

## Dependencies

Install the following before expanding the application:

- PySide6
- python-docx
- openpyxl
- Pillow
- pandas
- lxml
- pytest
- pyinstaller

## Running tests

Once Python is available in the environment, run:

```bash
pytest -q
```

## Packaging for Windows

To create a standalone Windows executable, install the development dependencies and run the build script:

```bash
python -m pip install -e .[dev]
build_windows.bat
```

This uses PyInstaller to produce a single-file app named `DocumentAutomationStudio.exe` in the `dist` folder.
