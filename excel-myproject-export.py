import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QDialog, QCheckBox, QHBoxLayout, QLabel, QScrollArea
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QDragEnterEvent, QDropEvent
import os
from datetime import datetime

class ExcelProcessorApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MyProject Excel Column Processor")
        self.setGeometry(100, 100, 600, 400)

        # Enable drag and drop
        self.setAcceptDrops(True)

        self.error_log_file = "error_log.txt"
        self.file_path = None
        self.excel_data = None
        self.columns = []
        self.dialog = None

        # Layout setup
        self.layout = QVBoxLayout(self)

        # Notice Label
        self.notice_label = QLabel("Drag and drop an MyProject Excel file here")
        self.notice_label.setAlignment(Qt.AlignCenter)
        self.notice_label.setFont(QFont("Arial", 14))
        self.layout.addWidget(self.notice_label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        file_path = event.mimeData().urls()[0].toLocalFile()
        if file_path.lower().endswith(('.xls', '.xlsx')):
            self.file_path = file_path
            try:
                self.excel_data = pd.read_excel(self.file_path, header=None)
                self.columns = list(self.excel_data.iloc[4])  # Row 5 is index 4
                self.show_column_selection_dialog()
            except Exception as e:
                self.log_error(f"Error reading the file: {str(e)}")

    def show_column_selection_dialog(self):
        if self.dialog:
            self.dialog.close()

        self.dialog = ColumnSelectionDialog(self.columns, self)
        self.dialog.exec_()

    def log_error(self, message):
        with open(self.error_log_file, 'a') as f:
            f.write(f"{datetime.now()}: {message}\n")


class ColumnSelectionDialog(QDialog):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Columns to Retain")
        self.setGeometry(150, 150, 600, 400)

        self.columns = columns
        self.selected_columns = {}

        self.layout = QVBoxLayout(self)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.scroll_content = QWidget()
        self.scroll_content_layout = QVBoxLayout(self.scroll_content)

        for idx, column in enumerate(self.columns):
            row_layout = QHBoxLayout()

            checkbox = QCheckBox(f"Column {idx + 1}: {column}")
            self.selected_columns[idx] = checkbox
            row_layout.addWidget(checkbox)

            self.scroll_content_layout.addLayout(row_layout)

        self.scroll_area.setWidget(self.scroll_content)
        self.layout.addWidget(self.scroll_area)

        button_layout = QHBoxLayout()

        self.select_all_button = QPushButton("Select All")
        self.select_all_button.clicked.connect(self.select_all_columns)
        button_layout.addWidget(self.select_all_button)

        self.unselect_all_button = QPushButton("Unselect All")
        self.unselect_all_button.clicked.connect(self.unselect_all_columns)
        button_layout.addWidget(self.unselect_all_button)

        self.layout.addLayout(button_layout)

        self.proceed_button = QPushButton("Proceed")
        self.proceed_button.clicked.connect(self.proceed)
        self.layout.addWidget(self.proceed_button)

    def select_all_columns(self):
        for idx, checkbox in self.selected_columns.items():
            checkbox.setChecked(True)

    def unselect_all_columns(self):
        for idx, checkbox in self.selected_columns.items():
            checkbox.setChecked(False)

    def proceed(self):
        try:
            retained_columns = [col_idx for col_idx, checkbox in self.selected_columns.items() if checkbox.isChecked()]
            updated_data = self.process_columns(retained_columns)

            # Generate new filename with datestamp prefix (to prevent overwriting the original)
            filename = os.path.basename(self.parent().file_path)
            new_filename = f"{datetime.now().strftime('%Y%m%d')}_{filename}"
            new_filepath = os.path.join(os.path.dirname(self.parent().file_path), new_filename)

            # Save to new Excel file without overwriting the original
            updated_data.to_excel(new_filepath, index=False, header=False)
            self.accept()  # Close dialog only after file is processed
            print(f"New file saved as: {new_filepath}")  # You can remove this later
        except Exception as e:
            self.parent().log_error(f"Error during processing: {str(e)}")

    def process_columns(self, retained_columns):
        data = self.parent().excel_data.copy()

        # Retain selected columns
        data = data.iloc[:, retained_columns]

        # Remove the first 4 rows (rows before column names)
        data = data.iloc[4:].reset_index(drop=True)  # Remove the first 4 rows and reset index

        # Generate new filename with datestamp prefix (to prevent overwriting the original)
        filename = os.path.basename(self.parent().file_path)
        new_filename = f"{datetime.now().strftime('%d.%m.%Y')}_{filename}"
        new_filepath = os.path.join(os.path.dirname(self.parent().file_path), new_filename)

        # Save the processed data to a new Excel file
        data.to_excel(new_filepath, index=False, header=False)
        
        return data


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelProcessorApp()
    window.show()
    sys.exit(app.exec_())
