import pandas as pd
import sys
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QGroupBox, QLabel,
    QVBoxLayout, QPushButton, QFileDialog
)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Corbion Badge System")

        name = QGroupBox('Corbion')
        layout = QVBoxLayout()

        # Create a button to open the file dialog
        self.open_button = QPushButton("Open Excel File")
        self.open_button.clicked.connect(self.open_file_dialog)
        layout.addWidget(self.open_button)

        # Create a label to display the selected file
        self.file_label = QLabel("No file selected")
        layout.addWidget(self.file_label)

        # Set the layout for the window
        name.setLayout(layout)
        self.setLayout(layout)

    def open_file_dialog(self):
        # Show file dialog with filter for Excel files
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx; *.xls)")

        if file:
            # Display the selected file path in the label
            self.file_label.setText(f"Selected file: {file}")
            self.load_excel_file(file)

    def load_excel_file(self, file):
        # Load the Excel file using pandas
        try:
            df = pd.read_excel(file)
            print(df.head())  # Just print the first few rows of the DataFrame for demo
        except Exception as e:
            print(f"Error loading Excel file: {e}")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

