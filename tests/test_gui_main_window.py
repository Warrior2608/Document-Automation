from __future__ import annotations

import os
import pytest

# Skip GUI tests if PySide6 is not installed in the environment.
pytest.importorskip("PySide6")

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtWidgets import QApplication

from document_automation_studio.gui.main_window import MainWindow
from document_automation_studio.processors.metadata_processor import MetadataValues


def test_main_window_initializes() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()

    assert window.windowTitle() == "Document Automation Studio"
    assert window.input_folder_edit is not None
    assert window.progress_bar.value() == 0
    assert window.profile_combo.count() >= 4
    window.close()
    app.quit()


def test_metadata_values_are_collected_and_forwarded() -> None:
    app = QApplication.instance() or QApplication([])
    window = MainWindow()

    window.prepared_by_name_edit.setText("Jane Doe")
    window.prepared_by_designation_edit.setText("Quality Manager")
    window.issue_spin.setValue(2)
    window.version_spin.setValue(3)

    values = window._collect_metadata_values()

    assert values.prepared_by_name == "Jane Doe"
    assert values.prepared_by_designation == "Quality Manager"
    assert values.issue == "2"
    assert values.version == "3"

    window.controller.set_metadata_values(values)
    assert window.controller.metadata_values.prepared_by_name == "Jane Doe"

    window.close()
    app.quit()
