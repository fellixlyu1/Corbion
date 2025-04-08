import sys
import openpyxl
from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout, QPushButton, QLineEdit
)

def update(material, quantity, size, date):
    file_path = 'Expiration_Tracker.xlsx'
    workbook = openpyxl.load_workbook(file_path)

    sheet_name = 'Sheet1'
    worksheet = workbook[sheet_name]

    data_inserted = False

    for row in worksheet.iter_rows():
        # Check if there is an empty cell in the row (i.e., the row is not completely filled)
        if any(cell.value is None for cell in row):
            # Find the first empty cell and update it with the corresponding values
            for i, cell in enumerate(row):
                if cell.value is None:
                    if i == 0:
                        cell.value = material
                    elif i == 1:
                        cell.value = quantity
                    elif i == 2:
                        cell.value = size
                    elif i == 3:
                        cell.value = date
                    data_inserted = True
                    break

        if data_inserted:
            break

    # Save the changes to the workbook
    workbook.save(file_path)

class ExpirationTracker(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Corbion Chemicals Expiration Tracker")
        self.material_field = QLineEdit(self)
        self.material_field.setPlaceholderText("Material: ")

        self.quantity = QLineEdit(self)
        self.quantity.setPlaceholderText("Quantity: ")

        self.size = QLineEdit(self)
        self.size.setPlaceholderText("Size: ")

        self.date = QLineEdit(self)
        self.date.setPlaceholderText("Expiry Date: ")

        self.add_button = QPushButton("Add New Material", self)

        layout = QVBoxLayout()
        layout.addWidget(self.material_field)
        layout.addWidget(self.quantity)
        layout.addWidget(self.size)
        layout.addWidget(self.date)

        self.setLayout(layout)

        self.add_button.clicked.connect(self.on_click)

    def on_click(self):
        update(self.material_field.text(), self.quantity.text(), self.size.text(), self.date.text())

if __name__ == '__main__':
    app = QApplication(sys.argv)