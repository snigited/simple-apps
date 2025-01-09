import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QMessageBox,
    QVBoxLayout, QWidget, QDialog, QCheckBox, QPushButton, QScrollArea
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont


class SheetSelectionDialog(QDialog):
    def __init__(self, sheets, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Sheets to Export")
        self.setModal(True)
        self.selected_sheets = []
        self.init_ui(sheets)

    def init_ui(self, sheets):
        layout = QVBoxLayout()

        # Instruction label
        instruction = QLabel("Select the sheets you want to export:")
        instruction.setFont(QFont("Arial", 12))
        layout.addWidget(instruction)

        # Scroll Area in case of many sheets
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        scroll_layout = QVBoxLayout(content)

        self.checkboxes = []
        for sheet in sheets:
            cb = QCheckBox(sheet)
            cb.setChecked(True)
            self.checkboxes.append(cb)
            scroll_layout.addWidget(cb)

        scroll.setWidget(content)
        layout.addWidget(scroll)

        # Confirm button
        btn_confirm = QPushButton("Export Selected Sheets")
        btn_confirm.setFixedHeight(40)
        btn_confirm.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        btn_confirm.clicked.connect(self.confirm)
        layout.addWidget(btn_confirm)

        self.setLayout(layout)
        self.resize(300, 400)

    def confirm(self):
        self.selected_sheets = [cb.text() for cb in self.checkboxes if cb.isChecked()]
        if not self.selected_sheets:
            QMessageBox.warning(self, "No Selection", "Please select at least one sheet to export.")
            return
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Sheet Exporter")
        self.setFixedSize(500, 400)
        self.setAcceptDrops(True)
        self.init_ui()

    def init_ui(self):
        # Central Widget
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        # Layout
        layout = QVBoxLayout()
        self.central_widget.setLayout(layout)

        # Instruction Label
        label = QLabel("Drag and drop an Excel file here")
        label.setAlignment(Qt.AlignCenter)
        label.setFont(QFont("Arial", 16))
        label.setStyleSheet("""
            QLabel {
                color: #333;
            }
        """)
        layout.addStretch()
        layout.addWidget(label)
        layout.addStretch()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            # Check if the dragged file is an Excel file
            urls = event.mimeData().urls()
            if len(urls) == 1 and urls[0].toLocalFile().lower().endswith(('.xlsx', '.xlsm', '.xls')):
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if len(urls) != 1:
            QMessageBox.warning(self, "Multiple Files", "Please drop only one Excel file.")
            return
        file_path = urls[0].toLocalFile()
        try:
            excel_file = pd.ExcelFile(file_path)
            sheets = excel_file.sheet_names
            if not sheets:
                QMessageBox.warning(self, "No Sheets", "The Excel file contains no sheets.")
                return
            dialog = SheetSelectionDialog(sheets, self)
            if dialog.exec_() == QDialog.Accepted:
                selected_sheets = dialog.selected_sheets
                self.export_sheets(file_path, selected_sheets)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read Excel file.\nError: {str(e)}")

    def export_sheets(self, original_path, sheets):
        try:
            # Determine the exported filename
            directory, filename = os.path.split(original_path)
            name, ext = os.path.splitext(filename)
            exported_name = f"{name}-exported{ext}"
            exported_path = os.path.join(directory, exported_name)

            # Export selected sheets
            with pd.ExcelWriter(exported_path, engine='openpyxl') as writer:
                for sheet in sheets:
                    df = pd.read_excel(original_path, sheet_name=sheet)
                    df.to_excel(writer, sheet_name=sheet, index=False)

            QMessageBox.information(self, "Success", f"Exported sheets to:\n{exported_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to export sheets.\nError: {str(e)}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
