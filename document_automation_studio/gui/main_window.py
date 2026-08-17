from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDoubleSpinBox,
    QFormLayout,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from document_automation_studio.core.controller import ApplicationController
from document_automation_studio.processors.branding_processor import BrandingSettings


class MainWindow(QMainWindow):
    """Main application window for Document Automation Studio."""

    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.controller = ApplicationController()
        self.logger = self.controller.logger
        self._theme = "dark"
        self._build_ui()
        self._apply_theme(self._theme)

    def _build_ui(self) -> None:
        self.setWindowTitle("Document Automation Studio")
        self.resize(1280, 920)

        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QScrollArea.Shape.NoFrame)

        scroll_content = QWidget(self)
        main_layout = QVBoxLayout(scroll_content)
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(18)

        scroll_area.setWidget(scroll_content)
        outer_layout = QVBoxLayout(central_widget)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.addWidget(scroll_area)

        header_label = QLabel("Document Automation Studio")
        header_label.setStyleSheet("font-size: 24px; font-weight: 600;")
        main_layout.addWidget(header_label)

        subtitle_label = QLabel("Batch document automation for Word and Excel workflows")
        subtitle_label.setStyleSheet("color: #8aa0b8;")
        main_layout.addWidget(subtitle_label)

        form_group = QGroupBox("Processing Setup")
        form_layout = QFormLayout(form_group)
        form_layout.setLabelAlignment(Qt.AlignmentFlag.AlignRight)

        self.input_folder_edit = QLineEdit()
        self.input_folder_edit.setPlaceholderText("Select input folder")
        self.input_folder_edit.setText(self.controller.config.processing.input_folder)
        self.input_folder_button = QPushButton("Browse")
        self.input_folder_button.clicked.connect(self._browse_input_folder)
        input_row = self._create_row(self.input_folder_edit, self.input_folder_button)
        form_layout.addRow("Input folder", input_row)

        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setPlaceholderText("Select output folder")
        self.output_folder_edit.setText(self.controller.config.processing.output_folder)
        self.output_folder_button = QPushButton("Browse")
        self.output_folder_button.clicked.connect(self._browse_output_folder)
        output_row = self._create_row(self.output_folder_edit, self.output_folder_button)
        form_layout.addRow("Output folder", output_row)

        self.rule_file_edit = QLineEdit()
        self.rule_file_edit.setPlaceholderText("Optional rule file")
        self.rule_file_edit.setText(self.controller.config.rule_files[0] if self.controller.config.rule_files else "")

        self.logo_path_edit = QLineEdit()
        self.logo_path_edit.setPlaceholderText("Optional replacement logo")
        self.logo_path_edit.setText(self.controller.config.processing.logo_path)
        self.logo_button = QPushButton("Browse")
        self.logo_button.clicked.connect(self._browse_logo_file)

        self.logo_width_spin = QDoubleSpinBox()
        self.logo_width_spin.setRange(0.1, 12.0)
        self.logo_width_spin.setSingleStep(0.1)
        self.logo_width_spin.setDecimals(2)
        self.logo_width_spin.setValue(float(self.controller.config.processing.logo_width_inches or 1.8))

        self.logo_height_spin = QDoubleSpinBox()
        self.logo_height_spin.setRange(0.1, 12.0)
        self.logo_height_spin.setSingleStep(0.1)
        self.logo_height_spin.setDecimals(2)
        self.logo_height_spin.setValue(float(self.controller.config.processing.logo_height_inches or 0.55))

        logo_controls = QWidget(self)
        logo_controls_layout = QHBoxLayout(logo_controls)
        logo_controls_layout.setContentsMargins(0, 0, 0, 0)
        logo_controls_layout.addWidget(self.logo_path_edit)
        logo_controls_layout.addWidget(self.logo_button)
        logo_controls_layout.addWidget(QLabel("W:"))
        logo_controls_layout.addWidget(self.logo_width_spin)
        logo_controls_layout.addWidget(QLabel("in"))
        logo_controls_layout.addWidget(QLabel("H:"))
        logo_controls_layout.addWidget(self.logo_height_spin)
        logo_controls_layout.addWidget(QLabel("in"))
        form_layout.addRow("Logo", logo_controls)
        self.logo_hint = QLabel("Use the logo width/height fields to match your header box size. Recommended starts around 1.8 in wide and 0.55 in high.")
        self.logo_hint.setWordWrap(True)
        self.logo_hint.setStyleSheet("color: #8aa0b8; font-size: 11px;")
        form_layout.addRow("", self.logo_hint)
        self.rule_file_button = QPushButton("Browse")
        self.rule_file_button.clicked.connect(self._browse_rule_file)
        rule_row = self._create_row(self.rule_file_edit, self.rule_file_button)
        form_layout.addRow("Rule file", rule_row)

        self.backup_checkbox = QCheckBox("Create backups")
        self.backup_checkbox.setChecked(self.controller.config.processing.create_backups)
        form_layout.addRow("Backup", self.backup_checkbox)

        self.backup_directory_edit = QLineEdit()
        self.backup_directory_edit.setPlaceholderText("Backup folder")
        self.backup_directory_edit.setText(self.controller.config.processing.backup_directory)
        self.backup_directory_button = QPushButton("Browse")
        self.backup_directory_button.clicked.connect(self._browse_backup_directory)
        backup_row = self._create_row(self.backup_directory_edit, self.backup_directory_button)
        form_layout.addRow("Backup folder", backup_row)

        self.profile_combo = QComboBox()
        default_profiles = ["Default", "Client A", "ISO 9001", "ISO 27001"]
        profile_items = []
        for profile in default_profiles + self.controller.config.profiles:
            if profile not in profile_items:
                profile_items.append(profile)
        self.profile_combo.addItems(profile_items)
        self.profile_combo.setCurrentText(self.controller.config.active_profile)
        self.profile_combo.currentTextChanged.connect(self._on_profile_changed)
        form_layout.addRow("Profile", self.profile_combo)

        self.include_subfolders_checkbox = QCheckBox("Include subfolders")
        self.include_subfolders_checkbox.setChecked(self.controller.config.processing.include_subfolders)
        self.preserve_structure_checkbox = QCheckBox("Preserve folder structure")
        self.preserve_structure_checkbox.setChecked(self.controller.config.processing.preserve_folder_structure)
        self.process_word_checkbox = QCheckBox("Process Word files")
        self.process_word_checkbox.setChecked(self.controller.config.processing.process_word_files)
        self.process_excel_checkbox = QCheckBox("Process Excel files")
        self.process_excel_checkbox.setChecked(self.controller.config.processing.process_excel_files)
        options_layout = QHBoxLayout()
        options_layout.addWidget(self.process_word_checkbox)
        options_layout.addWidget(self.process_excel_checkbox)
        options_layout.addWidget(self.include_subfolders_checkbox)
        options_layout.addWidget(self.preserve_structure_checkbox)
        form_layout.addRow("Options", options_layout)

        main_layout.addWidget(form_group)

        self.tabs = QTabWidget(self)
        self.tabs.addTab(self._build_metadata_tab(), "Document Metadata")
        main_layout.addWidget(self.tabs)

        actions_row = QHBoxLayout()
        self.preview_button = QPushButton("Preview")
        self.preview_button.clicked.connect(self._preview_changes)
        self.run_button = QPushButton("Run Automation")
        self.run_button.clicked.connect(self._run_automation)
        self.theme_button = QPushButton("Toggle Theme")
        self.theme_button.clicked.connect(self._toggle_theme)

        actions_row.addWidget(self.preview_button)
        actions_row.addWidget(self.run_button)
        actions_row.addStretch()
        actions_row.addWidget(self.theme_button)
        main_layout.addLayout(actions_row)

        progress_group = QGroupBox("Progress")
        progress_layout = QVBoxLayout(progress_group)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_label = QLabel("Ready to process files")
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.progress_label)
        main_layout.addWidget(progress_group)

        backup_group = QGroupBox("Backup Actions")
        backup_layout = QHBoxLayout(backup_group)
        self.ensure_backup_button = QPushButton("Ensure Backup Folder")
        self.ensure_backup_button.clicked.connect(self._ensure_backup_folder)
        self.restore_backup_button = QPushButton("Restore Backups")
        self.restore_backup_button.clicked.connect(self._restore_backups)
        self.delete_backup_button = QPushButton("Delete Backups")
        self.delete_backup_button.clicked.connect(self._delete_backups)
        backup_layout.addWidget(self.ensure_backup_button)
        backup_layout.addWidget(self.restore_backup_button)
        backup_layout.addWidget(self.delete_backup_button)
        main_layout.addWidget(backup_group)

        preview_group = QGroupBox("Preview")
        preview_layout = QVBoxLayout(preview_group)
        self.preview_table = QTableWidget(0, 5)
        self.preview_table.setHorizontalHeaderLabels(["File", "Change", "Old", "New", "Details"])
        self.preview_count_label = QLabel("0 preview items")
        preview_layout.addWidget(self.preview_table)
        preview_layout.addWidget(self.preview_count_label)
        main_layout.addWidget(preview_group)

        log_group = QGroupBox("Live Log")
        log_layout = QVBoxLayout(log_group)
        self.log_view = QPlainTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setPlainText("Application started.\nWaiting for configuration.")
        log_layout.addWidget(self.log_view)
        main_layout.addWidget(log_group)

        self.statusBar().showMessage("Ready")

    def _build_metadata_tab(self) -> QWidget:
        tab_widget = QWidget(self)
        layout = QVBoxLayout(tab_widget)
        form_group = QGroupBox("Metadata Values")
        form_layout = QFormLayout(form_group)

        self.effective_date_edit = QDateEdit()
        self.effective_date_edit.setCalendarPopup(True)
        self.effective_date_edit.setDisplayFormat("dd-MM-yyyy")
        form_layout.addRow("Effective Date", self.effective_date_edit)

        self.latest_revision_date_edit = QDateEdit()
        self.latest_revision_date_edit.setCalendarPopup(True)
        self.latest_revision_date_edit.setDisplayFormat("dd-MM-yyyy")
        form_layout.addRow("Latest Revision Date", self.latest_revision_date_edit)

        self.next_review_date_edit = QDateEdit()
        self.next_review_date_edit.setCalendarPopup(True)
        self.next_review_date_edit.setDisplayFormat("dd-MM-yyyy")
        form_layout.addRow("Next Review Date", self.next_review_date_edit)

        self.issue_spin = QSpinBox()
        self.issue_spin.setRange(0, 9999)
        form_layout.addRow("Issue", self.issue_spin)

        self.version_spin = QSpinBox()
        self.version_spin.setRange(0, 9999)
        form_layout.addRow("Version", self.version_spin)

        self.prepared_by_name_edit = QLineEdit()
        form_layout.addRow("Prepared By Name", self.prepared_by_name_edit)
        self.prepared_by_designation_edit = QLineEdit()
        form_layout.addRow("Prepared By Designation", self.prepared_by_designation_edit)

        self.checked_by_name_edit = QLineEdit()
        form_layout.addRow("Checked By Name", self.checked_by_name_edit)
        self.checked_by_designation_edit = QLineEdit()
        form_layout.addRow("Checked By Designation", self.checked_by_designation_edit)

        self.approved_by_name_edit = QLineEdit()
        form_layout.addRow("Approved By Name", self.approved_by_name_edit)
        self.approved_by_designation_edit = QLineEdit()
        form_layout.addRow("Approved By Designation", self.approved_by_designation_edit)

        layout.addWidget(form_group)

        actions_row = QHBoxLayout()
        self.load_defaults_button = QPushButton("Load Previous Values")
        self.load_defaults_button.clicked.connect(self._load_metadata_defaults)
        self.save_defaults_button = QPushButton("Save as Default")
        self.save_defaults_button.clicked.connect(self._save_metadata_defaults)
        self.clear_metadata_button = QPushButton("Clear")
        self.clear_metadata_button.clicked.connect(self._clear_metadata_form)
        self.preview_metadata_button = QPushButton("Preview")
        self.preview_metadata_button.clicked.connect(self._preview_metadata)
        self.apply_metadata_button = QPushButton("Apply to Documents")
        self.apply_metadata_button.clicked.connect(self._apply_metadata_to_documents)
        actions_row.addWidget(self.load_defaults_button)
        actions_row.addWidget(self.save_defaults_button)
        actions_row.addWidget(self.clear_metadata_button)
        actions_row.addWidget(self.preview_metadata_button)
        actions_row.addWidget(self.apply_metadata_button)
        layout.addLayout(actions_row)

        branding_group = QGroupBox("Company Branding")
        branding_layout = QFormLayout(branding_group)
        self.current_company_name_edit = QLineEdit()
        self.current_company_name_edit.setPlaceholderText("Leave blank to auto-detect")
        branding_layout.addRow("Current Company Name (Optional)", self.current_company_name_edit)
        self.new_company_name_edit = QLineEdit()
        self.new_company_name_edit.setPlaceholderText("Required")
        branding_layout.addRow("New Company Name", self.new_company_name_edit)
        self.replace_header_checkbox = QCheckBox("Replace Company Name in Header")
        self.replace_header_checkbox.setChecked(True)
        branding_layout.addRow("", self.replace_header_checkbox)
        self.replace_footer_checkbox = QCheckBox("Replace Company Name in Footer (Optional)")
        branding_layout.addRow("", self.replace_footer_checkbox)
        self.replace_body_checkbox = QCheckBox("Replace Company Name in Document Body")
        branding_layout.addRow("", self.replace_body_checkbox)
        self.replace_tables_checkbox = QCheckBox("Replace Company Name inside Tables")
        branding_layout.addRow("", self.replace_tables_checkbox)
        self.replace_text_boxes_checkbox = QCheckBox("Replace Company Name inside Text Boxes")
        branding_layout.addRow("", self.replace_text_boxes_checkbox)
        layout.addWidget(branding_group)

        self.metadata_preview = QPlainTextEdit()
        self.metadata_preview.setReadOnly(True)
        self.metadata_preview.setPlainText("Metadata preview will appear here.")
        layout.addWidget(self.metadata_preview)

        self._populate_metadata_form_from_controller()
        return tab_widget

    def _create_row(self, field: QLineEdit, button: QPushButton) -> QWidget:
        container = QWidget(self)
        layout = QHBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(field)
        layout.addWidget(button)
        return container

    def _populate_metadata_form_from_controller(self) -> None:
        values = self.controller.metadata_values
        self.effective_date_edit.setDate(self._parse_date(values.effective_date))
        self.latest_revision_date_edit.setDate(self._parse_date(values.latest_revision_date))
        self.next_review_date_edit.setDate(self._parse_date(values.next_review_date))
        self.issue_spin.setValue(int(values.issue) if values.issue.isdigit() else 0)
        self.version_spin.setValue(int(values.version) if values.version.isdigit() else 0)
        self.prepared_by_name_edit.setText(values.prepared_by_name)
        self.prepared_by_designation_edit.setText(values.prepared_by_designation)
        self.checked_by_name_edit.setText(values.checked_by_name)
        self.checked_by_designation_edit.setText(values.checked_by_designation)
        self.approved_by_name_edit.setText(values.approved_by_name)
        self.approved_by_designation_edit.setText(values.approved_by_designation)

    def _load_metadata_defaults(self) -> None:
        values = self.controller.load_metadata_defaults()
        self.controller.set_metadata_values(values)
        self._populate_metadata_form_from_controller()
        self.log_view.appendPlainText("Loaded metadata defaults from local configuration.")

    def _save_metadata_defaults(self) -> None:
        values = self._collect_metadata_values()
        self.controller.set_metadata_values(values)
        self.controller.save_metadata_defaults(values)
        self.log_view.appendPlainText("Saved metadata defaults to local configuration.")

    def _clear_metadata_form(self) -> None:
        self.effective_date_edit.setDate(self._parse_date(""))
        self.latest_revision_date_edit.setDate(self._parse_date(""))
        self.next_review_date_edit.setDate(self._parse_date(""))
        self.issue_spin.setValue(0)
        self.version_spin.setValue(0)
        for field in [
            self.prepared_by_name_edit,
            self.prepared_by_designation_edit,
            self.checked_by_name_edit,
            self.checked_by_designation_edit,
            self.approved_by_name_edit,
            self.approved_by_designation_edit,
        ]:
            field.clear()
        self.metadata_preview.setPlainText("Metadata form cleared.")

    def _preview_metadata(self) -> None:
        values = self._collect_metadata_values()
        branding_settings = self._collect_branding_settings()
        preview_lines = [
            f"Effective Date\n{values.effective_date or '-'}\n↓\n{values.effective_date or '-'}",
            f"Prepared By\n{self._combine_name_designation(values.prepared_by_name, values.prepared_by_designation) or '-'}",
            f"Issue\n{values.issue or '-'}\n↓\n{values.issue or '-'}",
            f"Version\n{values.version or '-'}\n↓\n{values.version or '-'}",
            f"Header Company Name\n{branding_settings.current_company_name or '-'}\n↓\n{branding_settings.new_company_name or '-'}",
        ]
        self.metadata_preview.setPlainText("\n\n".join(preview_lines))

    def _apply_metadata_to_documents(self) -> None:
        values = self._collect_metadata_values()
        branding_settings = self._collect_branding_settings()
        self.controller.set_metadata_values(values)
        self.controller.set_branding_settings(branding_settings)
        results = self.controller.apply_metadata_to_documents(values)
        self.log_view.appendPlainText(f"Applied metadata to {len(results)} document(s).")
        for result in results:
            status = "OK" if result.success else "FAILED"
            self.log_view.appendPlainText(
                f"{result.document_name}: table={result.metadata_table_found}; updated={','.join(result.updated_fields) or '-'}; missing={','.join(result.missing_fields) or '-'}; status={status}"
            )

    def _collect_branding_settings(self) -> BrandingSettings:
        return BrandingSettings(
            current_company_name=self.current_company_name_edit.text().strip(),
            new_company_name=self.new_company_name_edit.text().strip(),
            replace_in_header=self.replace_header_checkbox.isChecked(),
            replace_in_footer=self.replace_footer_checkbox.isChecked(),
            replace_in_body=self.replace_body_checkbox.isChecked(),
            replace_in_tables=self.replace_tables_checkbox.isChecked(),
            replace_in_text_boxes=self.replace_text_boxes_checkbox.isChecked(),
        )

    def _collect_metadata_values(self) -> object:
        from document_automation_studio.processors.metadata_processor import MetadataValues

        return MetadataValues(
            effective_date=self._format_date(self.effective_date_edit.date()),
            latest_revision_date=self._format_date(self.latest_revision_date_edit.date()),
            next_review_date=self._format_date(self.next_review_date_edit.date()),
            issue=str(self.issue_spin.value()),
            version=str(self.version_spin.value()),
            prepared_by_name=self.prepared_by_name_edit.text().strip(),
            prepared_by_designation=self.prepared_by_designation_edit.text().strip(),
            checked_by_name=self.checked_by_name_edit.text().strip(),
            checked_by_designation=self.checked_by_designation_edit.text().strip(),
            approved_by_name=self.approved_by_name_edit.text().strip(),
            approved_by_designation=self.approved_by_designation_edit.text().strip(),
        )

    def _format_date(self, date) -> str:
        return date.toString("dd-MM-yyyy") if date is not None else ""

    def _parse_date(self, value: str) -> object:
        from PySide6.QtCore import QDate

        if not value:
            return QDate.currentDate()
        try:
            day, month, year = map(int, value.split("-"))
            return QDate(year, month, day)
        except ValueError:
            return QDate.currentDate()

    def _combine_name_designation(self, name: str, designation: str) -> str:
        if not name and not designation:
            return ""
        if name and designation:
            return f"{name} / {designation}"
        return name or designation

    def _browse_input_folder(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Select input folder")
        if directory:
            self.input_folder_edit.setText(directory)

    def _browse_output_folder(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Select output folder")
        if directory:
            self.output_folder_edit.setText(directory)

    def _browse_rule_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "Select rule file", filter="JSON Files (*.json);;CSV Files (*.csv)")
        if file_path:
            self.rule_file_edit.setText(file_path)
            self._load_rule_file(Path(file_path))

    def _load_rule_file(self, rule_file_path: Path) -> bool:
        try:
            self.controller.load_rule_file(rule_file_path)
            self.statusBar().showMessage(f"Loaded rule file: {rule_file_path}")
            self.log_view.appendPlainText(f"Loaded rule file: {rule_file_path}")
            return True
        except Exception as error:
            QMessageBox.warning(self, "Rule Load Error", f"Could not load rule file: {error}")
            self.statusBar().showMessage("Invalid rule file selected")
            self.log_view.appendPlainText(f"Failed to load rule file: {error}")
            return False

    def _browse_logo_file(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(self, "Select logo file", filter="Image Files (*.png *.jpg *.jpeg *.bmp *.gif)")
        if file_path:
            self.logo_path_edit.setText(file_path)
            self.controller.config.processing.logo_path = file_path
            self.controller.config.processing.logo_width_inches = float(self.logo_width_spin.value())
            self.controller.config.processing.logo_height_inches = float(self.logo_height_spin.value())
            self.controller.save_config()

    def _browse_backup_directory(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "Select backup folder")
        if directory:
            self.backup_directory_edit.setText(directory)

    def _ensure_backup_folder(self) -> None:
        input_folder = self.input_folder_edit.text()
        if not input_folder:
            QMessageBox.warning(self, "Backup Error", "Please select an input folder before creating backups.")
            return
        self.controller.config.processing.backup_directory = self.backup_directory_edit.text() or self.controller.config.processing.backup_directory
        self.controller.config.processing.logo_path = self.logo_path_edit.text().strip()
        self.controller.config.processing.logo_width_inches = float(self.logo_width_spin.value())
        self.controller.config.processing.logo_height_inches = float(self.logo_height_spin.value())
        self.controller.update_processing_settings(self.controller.config.processing)
        self.controller.create_backup(Path(input_folder))
        self.log_view.appendPlainText("Backup folder ensured.")

    def _restore_backups(self) -> None:
        output_folder = self.output_folder_edit.text()
        if not output_folder:
            QMessageBox.warning(self, "Restore Error", "Please select an output folder before restoring backups.")
            return
        self.controller.restore_backup(Path(output_folder))
        self.log_view.appendPlainText("Backups restored to output folder.")

    def _delete_backups(self) -> None:
        self.controller.delete_backup()
        self.log_view.appendPlainText("All backups deleted.")

    def _preview_settings(self) -> None:
        message = (
            f"Input: {self.input_folder_edit.text()}\n"
            f"Output: {self.output_folder_edit.text()}\n"
            f"Profile: {self.profile_combo.currentText()}\n"
            f"Rule file: {self.rule_file_edit.text()}\n"
            f"Logo size: {self.logo_width_spin.value():.2f} x {self.logo_height_spin.value():.2f} in"
        )
        QMessageBox.information(self, "Preview", message)
        self.logger.info("Preview requested")

    def _preview_changes(self) -> None:
        self.progress_label.setText("Generating preview")
        self.preview_table.setRowCount(0)
        try:
            if self.rule_file_edit.text():
                self.controller.load_rule_file(Path(self.rule_file_edit.text()))
        except Exception as error:
            QMessageBox.warning(self, "Rule Load Error", f"Could not load rule file: {error}")
            return

        preview_items = self.controller.preview_changes(max_items=50)
        for item in preview_items:
            row = self.preview_table.rowCount()
            self.preview_table.insertRow(row)
            self.preview_table.setItem(row, 0, QTableWidgetItem(str(item.file_path)))
            self.preview_table.setItem(row, 1, QTableWidgetItem(item.change_type))
            self.preview_table.setItem(row, 2, QTableWidgetItem(item.old_value))
            self.preview_table.setItem(row, 3, QTableWidgetItem(item.new_value))
            self.preview_table.setItem(row, 4, QTableWidgetItem(item.details))

        self.preview_count_label.setText(f"{len(preview_items)} preview items")
        self.progress_label.setText(f"Preview items: {len(preview_items)}")
        self.logger.info("Preview generated with %d items", len(preview_items))

    def _run_automation(self) -> None:
        self.progress_bar.setValue(0)
        self.progress_label.setText("Starting batch operation")
        self.log_view.appendPlainText("Batch automation started.")
        self.logger.info("Automation run requested")

        progress_callback = self._on_batch_progress
        self.controller.set_metadata_values(self._collect_metadata_values())
        self.controller.config.processing.input_folder = self.input_folder_edit.text().strip()
        self.controller.config.processing.output_folder = self.output_folder_edit.text().strip()
        self.controller.config.processing.process_word_files = self.process_word_checkbox.isChecked()
        self.controller.config.processing.process_excel_files = self.process_excel_checkbox.isChecked()
        self.controller.config.processing.include_subfolders = self.include_subfolders_checkbox.isChecked()
        self.controller.config.processing.preserve_folder_structure = self.preserve_structure_checkbox.isChecked()
        self.controller.config.processing.create_backups = self.backup_checkbox.isChecked()
        self.controller.config.processing.backup_directory = self.backup_directory_edit.text() or self.controller.config.processing.backup_directory
        self.controller.config.processing.logo_path = self.logo_path_edit.text() or self.controller.config.processing.logo_path
        if self.rule_file_edit.text():
            try:
                self.controller.load_rule_file(Path(self.rule_file_edit.text()))
            except Exception as error:
                QMessageBox.warning(self, "Rule Load Error", f"Could not load rule file: {error}")
        self.controller.update_processing_settings(self.controller.config.processing)

        self.controller.run_batch(progress_callback=progress_callback)

        self.progress_bar.setValue(100)
        self.progress_label.setText("Completed")
        self.log_view.appendPlainText("Batch automation completed.")

    def _on_batch_progress(self, current: int, total: int) -> None:
        progress = int((current / total) * 100)
        self.progress_bar.setValue(progress)
        self.progress_label.setText(f"Processing {current}/{total}")

    def _toggle_theme(self) -> None:
        self._theme = "light" if self._theme == "dark" else "dark"
        self._apply_theme(self._theme)

    def _on_profile_changed(self, profile_name: str) -> None:
        self.controller.update_active_profile(profile_name)

    def _apply_theme(self, theme: str) -> None:
        if theme == "dark":
            self.setStyleSheet(
                "QMainWindow { background-color: #121b26; color: #f5f7fb; }"
                "QGroupBox { border: 1px solid #2a3b4f; border-radius: 8px; padding-top: 10px; }"
                "QLineEdit, QComboBox, QPlainTextEdit, QPushButton { border-radius: 6px; padding: 6px; }"
                "QPushButton { background-color: #1f6feb; color: white; }"
            )
        else:
            self.setStyleSheet(
                "QMainWindow { background-color: #f6f8fb; color: #18212f; }"
                "QGroupBox { border: 1px solid #d9e2ef; border-radius: 8px; padding-top: 10px; }"
                "QLineEdit, QComboBox, QPlainTextEdit, QPushButton { border-radius: 6px; padding: 6px; }"
                "QPushButton { background-color: #1f6feb; color: white; }"
            )
