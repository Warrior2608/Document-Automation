import os
os.environ['QT_QPA_PLATFORM'] = 'offscreen'
from PySide6.QtWidgets import QApplication
from document_automation_studio.gui.main_window import MainWindow

app = QApplication.instance() or QApplication([])
window = MainWindow()
print('window-created')
