from __future__ import annotations

import sys

from PySide6.QtWidgets import QApplication

from document_automation_studio.gui.main_window import MainWindow


def launch_app() -> int:
    """Create and launch the main application window."""
    app = QApplication.instance() or QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(launch_app())
