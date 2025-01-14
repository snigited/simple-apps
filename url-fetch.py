import sys
import os
import requests
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QCheckBox
)
from PyQt5.QtCore import Qt, QSettings
from PyQt5.QtGui import QTextCharFormat, QColor, QTextCursor
from datetime import datetime
from pathlib import Path
import difflib


class HTMLFetcher(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.htmls_dir = Path("htmls")
        self.htmls_dir.mkdir(exist_ok=True)
        self.settings = QSettings("MyCompany", "HTMLFetcherApp")
        self.load_settings()

    def initUI(self):
        self.setWindowTitle("HTML Fetcher and Comparator")

        layout = QVBoxLayout()

        # URL input
        url_layout = QHBoxLayout()
        url_label = QLabel("URL:")
        self.url_input = QLineEdit()
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.url_input)
        layout.addLayout(url_layout)

        # SSL Verification Checkbox
        ssl_layout = QHBoxLayout()
        self.ssl_checkbox = QCheckBox("Ignore SSL Certificate Errors")
        self.ssl_checkbox.setChecked(False)  # Default is to verify SSL
        ssl_layout.addWidget(self.ssl_checkbox)
        ssl_layout.addStretch()
        layout.addLayout(ssl_layout)

        # Fetch button
        self.fetch_button = QPushButton("Fetch and Save HTML")
        self.fetch_button.clicked.connect(self.fetch_and_save)
        layout.addWidget(self.fetch_button)

        # Output box
        self.output_box = QTextEdit()
        self.output_box.setReadOnly(True)
        layout.addWidget(self.output_box)

        # Optional: Clear Output Button
        clear_layout = QHBoxLayout()
        self.clear_button = QPushButton("Clear Output")
        self.clear_button.clicked.connect(self.clear_output)
        clear_layout.addStretch()
        clear_layout.addWidget(self.clear_button)
        layout.addLayout(clear_layout)

        self.setLayout(layout)
        self.resize(800, 600)

    def load_settings(self):
        # Load saved URL
        saved_url = self.settings.value("url", "")
        self.url_input.setText(saved_url)

        # Load SSL checkbox state
        saved_ssl = self.settings.value("ignore_ssl", "false").lower()
        self.ssl_checkbox.setChecked(saved_ssl == "true")

    def save_settings(self):
        # Save URL
        url = self.url_input.text().strip()
        self.settings.setValue("url", url)

        # Save SSL checkbox state
        ignore_ssl = self.ssl_checkbox.isChecked()
        self.settings.setValue("ignore_ssl", str(ignore_ssl).lower())

    def fetch_and_save(self):
        url = self.url_input.text().strip()
        ignore_ssl = self.ssl_checkbox.isChecked()

        if not url:
            self.append_output("Error: Please enter a URL.", "red")
            return

        self.save_settings()

        try:
            response = requests.get(url, verify=not ignore_ssl, timeout=10)
            response.raise_for_status()
            current_html = response.text
        except requests.exceptions.SSLError as ssl_err:
            self.append_output(f"SSL Error: {ssl_err}", "red")
            return
        except requests.exceptions.RequestException as e:
            self.append_output(f"Request Error: Failed to fetch URL:\n{e}", "red")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = self.htmls_dir / f"{timestamp}.html"

        # Get the last saved HTML
        last_file = self.get_last_saved_file()

        if last_file:
            try:
                with open(last_file, 'r', encoding='utf-8') as f:
                    last_html = f.read()
            except Exception as e:
                self.append_output(f"File Error: Failed to read last HTML file:\n{e}", "red")
                return

            if current_html == last_html:
                # No difference, do not save the new HTML
                self.append_output("No differences found. New HTML not saved.", "green")
                return
            else:
                # Differences found, process and display them
                diff = difflib.ndiff(last_html.splitlines(), current_html.splitlines())
                added = []
                removed = []
                for line in diff:
                    # Ignore lines that are only whitespace or empty
                    content = line[2:].strip()
                    if not content:
                        continue

                    if line.startswith('+ ') and not line.startswith('+++'):
                        added.append(line[2:])
                    elif line.startswith('- ') and not line.startswith('---'):
                        removed.append(line[2:])

                if not added and not removed:
                    self.append_output("Differences found, but could not parse them.", "orange")
                    return

                # Save the new HTML
                try:
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(current_html)
                except Exception as e:
                    self.append_output(f"File Error: Failed to save new HTML file:\n{e}", "red")
                    return

                # Display differences
                self.display_differences(added, removed, filename.name)
        else:
            # No previous file, save the current HTML
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(current_html)
                self.append_output(f"HTML saved as {filename.name}.", "green")
            except Exception as e:
                self.append_output(f"File Error: Failed to save HTML file:\n{e}", "red")
                return

    def get_last_saved_file(self):
        files = sorted(self.htmls_dir.glob("*.html"), key=os.path.getmtime)
        if files:
            return files[-1]
        return None

    def append_output(self, message, color="black"):
        """
        Appends a message to the output box with the specified color.
        """
        format = QTextCharFormat()
        format.setForeground(QColor(color))
        cursor = self.output_box.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.output_box.setTextCursor(cursor)
        self.output_box.setCurrentCharFormat(format)
        self.output_box.append(message)
        self.output_box.setCurrentCharFormat(QTextCharFormat())  # Reset to default

    def display_differences(self, added, removed, filename):
        """
        Displays added and removed lines with color highlighting.
        Added lines are shown in green.
        Removed lines are shown in red.
        """
        self.append_output(f"Differences found compared to {filename}:", "blue")

        if added:
            self.append_output("\nAdded Lines:", "green")
            for line in added:
                self.append_output(f"+ {line}", "green")

        if removed:
            self.append_output("\nRemoved Lines:", "red")
            for line in removed:
                self.append_output(f"- {line}", "red")

        self.append_output("\n", "black")  # Add an empty line for separation

    def clear_output(self):
        """
        Clears the output textbox.
        """
        self.output_box.clear()


def main():
    app = QApplication(sys.argv)
    fetcher = HTMLFetcher()
    fetcher.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
