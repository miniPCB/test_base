import sys
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QFormLayout, QLabel, QLineEdit,
    QPushButton, QWidget, QFileDialog, QMessageBox
)

class TestPlanApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Test Plan Editor")
        self.test_plan = {}

        # Main Layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)

        # Form Layout for Test Plan Fields
        self.form_layout = QFormLayout()
        self.fields = {}
        
        field_names = [
            "test_name", "test_no", "last_test_no", "purpose",
            "scope", "setup", "procedure", "measurement",
            "acceptancecriteria", "conclusion"
        ]

        for field in field_names:
            label = QLabel(field.replace("_", " ").capitalize())
            line_edit = QLineEdit()
            self.fields[field] = line_edit
            self.form_layout.addRow(label, line_edit)

        self.layout.addLayout(self.form_layout)

        # Buttons
        self.load_button = QPushButton("Load JSON")
        self.load_button.clicked.connect(self.load_json)
        self.layout.addWidget(self.load_button)

        self.save_button = QPushButton("Save JSON")
        self.save_button.clicked.connect(self.save_json)
        self.layout.addWidget(self.save_button)

        self.new_button = QPushButton("New Test Plan")
        self.new_button.clicked.connect(self.new_test_plan)
        self.layout.addWidget(self.new_button)

    def load_json(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Test Plan JSON", "", "JSON Files (*.json)", options=options
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r') as file:
                self.test_plan = json.load(file)
            
            for key, line_edit in self.fields.items():
                value = self.test_plan.get(key, "")
                line_edit.setText(value)

            QMessageBox.information(self, "Success", "Test plan loaded successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load JSON: {str(e)}")

    def save_json(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Test Plan JSON", "", "JSON Files (*.json)", options=options
        )

        if not file_path:
            return

        try:
            for key, line_edit in self.fields.items():
                self.test_plan[key] = line_edit.text()

            with open(file_path, 'w') as file:
                json.dump(self.test_plan, file, indent=4)

            QMessageBox.information(self, "Success", "Test plan saved successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save JSON: {str(e)}")

    def new_test_plan(self):
        self.test_plan = {}
        for line_edit in self.fields.values():
            line_edit.clear()
        QMessageBox.information(self, "New Test Plan", "New test plan created.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TestPlanApp()
    window.show()
    sys.exit(app.exec_())
