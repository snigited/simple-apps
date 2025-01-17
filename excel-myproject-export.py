import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QPushButton, QFileDialog,
    QDialog, QScrollArea, QCheckBox, QSpinBox, QHBoxLayout, QGridLayout, QMessageBox
)
from PyQt5.QtCore import Qt, QDateTime


def process_excel(file_path, selected_columns, column_widths, wrap_text_columns, output_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, header=None)

        # Filter columns
        df_filtered = df.iloc[:, selected_columns]

        # Create Excel writer
        writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
        df_filtered.to_excel(writer, index=False, header=False)

        # Format Excel file
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        for idx, col_idx in enumerate(selected_columns):
            col_width = column_widths.get(col_idx, 10)
            wrap = wrap_text_columns.get(col_idx, False)
            col_letter = chr(65 + idx)  # Map index to Excel column letter

            # Set column width
            worksheet.set_column(f'{col_letter}:{col_letter}', col_width)

            # Set text wrapping
            if wrap:
                format_wrap = workbook.add_format({'text_wrap': True})
                worksheet.set_column(f'{col_letter}:{col_letter}', None, format_wrap)

        writer.save()

    except Exception as e:
        with open('error_log.txt', 'a') as error_log:
            error_log.write(f"Error processing file {file_path}: {str(e)}\n")


class ColumnSelectionDialog(QDialog):
    def __init__(self, columns, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Columns and Options")
        self.selected_columns = []
        self.column_widths = {}
        self.wrap_text_columns = {}

        layout = QVBoxLayout()
        self.checkboxes = []
        self.width_spinners = []
        self.wrap_checkboxes = []

        # Scrollable area for checkboxes
        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        grid_layout = QGridLayout()

        # Add columns as checkboxes with options
        for idx, col in enumerate(columns):
            col_checkbox = QCheckBox(f"Column {col}")
            col_checkbox.setChecked(True)
            self.checkboxes.append(col_checkbox)

            width_spinner = QSpinBox()
            width_spinner.setMinimum(5)
            width_spinner.setMaximum(100)
            width_spinner.setValue(10)
            self.width_spinners.append(width_spinner)

            wrap_checkbox = QCheckBox("Wrap Text")
            self.wrap_checkboxes.append(wrap_checkbox)

            grid_layout.addWidget(col_checkbox, idx, 0)
            grid_layout.addWidget(width_spinner, idx, 1)
            grid_layout.addWidget(wrap_checkbox, idx, 2)

        scroll_widget.setLayout(grid_layout)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)

        # Add Select All and Unselect All buttons
        button_layout = QHBoxLayout()
        self.select_all_button = QPushButton("Select All")
        self.unselect_all_button = QPushButton("Unselect All")
        button_layout.addWidget(self.select_all_button)
        button_layout.addWidget(self.unselect_all_button)

        self.select_all_button.clicked.connect(self.select_all)
        self.unselect_all_button.clicked.connect(self.unselect_all)

        # Add Proceed button
        self.proceed_button = QPushButton("Proceed")
        self.proceed_button.clicked.connect(self.accept)

        layout.addWidget(scroll_area)
        layout.addLayout(button_layout)
        layout.addWidget(self.proceed_button)
        self.setLayout(layout)

    def select_all(self):
        for checkbox in self.checkboxes:
            checkbox.setChecked(True)

    def unselect_all(self):
        for checkbox in self.checkboxes:
            checkbox.setChecked(False)

    def accept(self):
        self.selected_columns = [idx for idx, cb in enumerate(self.checkboxes) if cb.isChecked()]
        self.column_widths = {idx: sp.value() for idx, sp in enumerate(self.width_spinners)}
        self.wrap_text_columns = {idx: cb.isChecked() for idx, cb in enumerate(self.wrap_checkboxes)}
        super().accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Column Filter")
        self.setGeometry(200, 200, 400, 200)

        layout = QVBoxLayout()
        self.label = QLabel("Drag and Drop Excel File Here")
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)

        self.process_button = QPushButton("Process")
        self.process_button.setEnabled(False)
        self.process_button.clicked.connect(self.process_file)
        layout.addWidget(self.process_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.file_path = None
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            self.file_path = urls[0].toLocalFile()
            if self.file_path.endswith('.xlsx'):
                self.label.setText(f"File Selected: {os.path.basename(self.file_path)}")
                self.process_button.setEnabled(True)
            else:
                QMessageBox.warning(self, "Invalid File", "Please drop a valid Excel file.")

    def process_file(self):
        if not self.file_path:
            QMessageBox.warning(self, "Error", "No file selected.")
            return

        try:
            df = pd.read_excel(self.file_path, header=None)
            columns = df.iloc[4, :].tolist()

            dialog = ColumnSelectionDialog(columns, self)
            if dialog.exec_() == QDialog.Accepted:
                selected_columns = dialog.selected_columns
                column_widths = dialog.column_widths
                wrap_text_columns = dialog.wrap_text_columns

                output_path = os.path.join(
                    os.path.dirname(self.file_path),
                    f"{QDateTime.currentDateTime().toString('yyyyMMdd_HHmmss')}_{os.path.basename(self.file_path)}"
                )

                process_excel(
                    self.file_path,
                    selected_columns,
                    column_widths,
                    wrap_text_columns,
                    output_path
                )

                QMessageBox.information(self, "Success", f"Processed file saved as: {output_path}")

        except Exception as e:
            with open('error_log.txt', 'a') as error_log:
                error_log.write(f"Error processing file {self.file_path}: {str(e)}\n")
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
