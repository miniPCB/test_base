import sys
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem,
    QWidget, QFileDialog, QMessageBox, QDialog, QLabel, QLineEdit, QComboBox, QFormLayout, QDialogButtonBox
)

class TestManager(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Test Manager")
        self.setGeometry(100, 100, 800, 600)

        self.tests = []
        self.current_file = "tests.json"

        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.table = QTableWidget()
        self.table.setColumnCount(5)  # Display limited columns for brevity
        self.table.setHorizontalHeaderLabels(["Test Name", "Test No", "UUT PN", "Test Type", "Purpose"])
        layout.addWidget(self.table)

        self.load_button = QPushButton("Load JSON")
        self.load_button.clicked.connect(self.load_json)
        layout.addWidget(self.load_button)

        self.save_button = QPushButton("Save JSON")
        self.save_button.clicked.connect(self.save_json)
        layout.addWidget(self.save_button)

        self.new_test_button = QPushButton("Create New Test")
        self.new_test_button.clicked.connect(self.create_new_test)
        layout.addWidget(self.new_test_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        self.load_json()  # Load the default file initially

    def load_json(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Open JSON File", "", "JSON Files (*.json)", options=options)
        if file_name:
            try:
                with open(file_name, "r") as file:
                    self.tests = json.load(file)
                    self.current_file = file_name
                    self.populate_table()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load JSON file: {e}")

    def save_json(self):
        try:
            with open(self.current_file, "w") as file:
                json.dump(self.tests, file, indent=4)
                QMessageBox.information(self, "Success", f"JSON saved to {self.current_file}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save JSON file: {e}")

    def populate_table(self):
        self.table.setRowCount(0)
        for test in self.tests:
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(test.get("test_name", "")))
            self.table.setItem(row_position, 1, QTableWidgetItem(test.get("test_no", "")))
            self.table.setItem(row_position, 2, QTableWidgetItem(test.get("uut_pn", "")))
            self.table.setItem(row_position, 3, QTableWidgetItem(test.get("test_type", "")))
            self.table.setItem(row_position, 4, QTableWidgetItem(test.get("purpose", "")))

    def create_new_test(self):
        dialog = NewTestDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            new_test = dialog.get_test_data()
            self.tests.append(new_test)
            self.populate_table()

class NewTestDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Create New Test")

        self.layout = QFormLayout()

        self.test_name_input = QLineEdit()
        self.layout.addRow("Test Name:", self.test_name_input)

        self.test_no_input = QLineEdit()
        self.layout.addRow("Test No:", self.test_no_input)

        self.uut_pn_input = QLineEdit()
        self.layout.addRow("UUT PN:", self.uut_pn_input)

        self.last_test_no_input = QLineEdit()
        self.layout.addRow("Last Test No:", self.last_test_no_input)

        self.test_type_input = QComboBox()
        self.test_type_input.addItems(["STATIC", "PERFORMANCE", "FINAL"])
        self.layout.addRow("Test Type:", self.test_type_input)

        self.purpose_input = QLineEdit()
        self.layout.addRow("Purpose:", self.purpose_input)

        self.scope_input = QLineEdit()
        self.layout.addRow("Scope:", self.scope_input)

        self.setup_input = QLineEdit()
        self.layout.addRow("Setup:", self.setup_input)

        self.procedure_input = QLineEdit()
        self.layout.addRow("Procedure:", self.procedure_input)

        self.measurement_input = QLineEdit()
        self.layout.addRow("Measurement:", self.measurement_input)

        self.parameter_input = QLineEdit()
        self.layout.addRow("Parameter:", self.parameter_input)

        self.ll_input = QLineEdit()
        self.layout.addRow("Lower Limit:", self.ll_input)

        self.tv_input = QLineEdit()
        self.layout.addRow("Target Value:", self.tv_input)

        self.ul_input = QLineEdit()
        self.layout.addRow("Upper Limit:", self.ul_input)

        self.units_input = QLineEdit()
        self.layout.addRow("Units:", self.units_input)

        self.buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)

        self.layout.addWidget(self.buttons)
        self.setLayout(self.layout)

    def get_test_data(self):
        return {
            "test_name": self.test_name_input.text(),
            "test_no": self.test_no_input.text(),
            "uut_pn": self.uut_pn_input.text(),
            "last_test_no": self.last_test_no_input.text(),
            "test_type": self.test_type_input.currentText(),
            "purpose": self.purpose_input.text(),
            "scope": self.scope_input.text(),
            "setup": self.setup_input.text(),
            "procedure": self.procedure_input.text(),
            "measurement": self.measurement_input.text(),
            "parameter": self.parameter_input.text(),
            "ll": self.ll_input.text(),
            "tv": self.tv_input.text(),
            "ul": self.ul_input.text(),
            "units": self.units_input.text()
        }

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TestManager()
    window.show()
    sys.exit(app.exec_())
