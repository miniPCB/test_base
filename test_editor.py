import sys
import json
from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QSplitter,
    QListWidget,
    QListWidgetItem,
    QTabWidget,
    QFormLayout,
    QLineEdit,
    QTextEdit,
    QPlainTextEdit,
    QVBoxLayout,
    QFileDialog,
    QMessageBox,
)


def build_markdown_from_test(test: dict) -> str:
    """Render a Markdown-ish view of a test dict."""
    def g(key, default=""):
        return test.get(key, default) or ""

    lines = []
    lines.append(f"# {g('test_name', 'Untitled Test')} (Test {g('test_no', '?')})")
    lines.append("")
    lines.append("## Purpose")
    lines.append(g("purpose"))
    lines.append("")
    lines.append("## Scope")
    lines.append(g("scope"))
    lines.append("")
    lines.append("## Setup")
    lines.append(g("setup"))
    lines.append("")
    lines.append("## Procedure")
    lines.append(g("procedure"))
    lines.append("")
    lines.append("## Measurement")
    lines.append(g("measurement"))
    lines.append("")
    lines.append("## Acceptance Criteria")
    lines.append(g("acceptancecriteria"))
    lines.append("")
    lines.append("## Conclusion")
    lines.append(g("conclusion"))
    lines.append("")
    # Revision history (simple table)
    rh = test.get("revision_history", [])
    if rh:
        lines.append("---")
        lines.append("")
        lines.append("## Revision History")
        lines.append("")
        lines.append("| Rev | Date | Description | By |")
        lines.append("|-----|------|-------------|----|")
        for entry in rh:
            rev = entry.get("rev", "")
            date = entry.get("rev_date", "")
            desc = entry.get("description", "")
            by = entry.get("rev_by", "")
            lines.append(f"| {rev} | {date} | {desc} | {by} |")
        lines.append("")

    return "\n".join(lines)


def build_variables_summary(test: dict) -> str:
    """Build a human-readable text summary of test variables for the right pane."""
    lines = []

    # Testpoints
    tps = test.get("testpoints", [])
    lines.append("TESTPOINTS")
    lines.append("----------")
    if tps:
        for tp in tps:
            name = tp.get("name", "")
            role = tp.get("role", "")
            ref = tp.get("schematic_ref", "")
            net = tp.get("net", "")
            desc = tp.get("description", "")
            lines.append(f"- {name} ({role}) [{ref}/{net}] - {desc}")
    else:
        lines.append("(none defined)")
    lines.append("")

    # Measurement equipment
    eqs = test.get("measurement_equipment", [])
    lines.append("MEASUREMENT EQUIPMENT")
    lines.append("---------------------")
    if eqs:
        for eq in eqs:
            eid = eq.get("id", "")
            etype = eq.get("type", "")
            model = eq.get("model", "")
            serial = eq.get("serial", "")
            loc = eq.get("location", "")
            notes = eq.get("notes", "")
            lines.append(f"- {eid}: {etype} {model} (SN {serial}), {loc} - {notes}")
    else:
        lines.append("(none defined)")
    lines.append("")

    # Measurement settings
    msets = test.get("measurement_settings", [])
    lines.append("MEASUREMENT SETTINGS")
    lines.append("---------------------")
    if msets:
        for ms in msets:
            eid = ms.get("equipment_id", "")
            mode = ms.get("mode", "")
            r = ms.get("range", "")
            samp = ms.get("sampling", "")
            other = ms.get("other_settings", "")
            lines.append(f"- {eid}: mode={mode}, range={r}, sampling={samp}, {other}")
    else:
        lines.append("(none defined)")
    lines.append("")

    # Acceptance thresholds
    ths = test.get("acceptance_thresholds", [])
    lines.append("ACCEPTANCE THRESHOLDS")
    lines.append("----------------------")
    if ths:
        for th in ths:
            p = th.get("parameter", "")
            tgt = th.get("target_value", "")
            lo = th.get("lower_limit", "")
            hi = th.get("upper_limit", "")
            units = th.get("units", "")
            notes = th.get("notes", "")
            lines.append(f"- {p}: target={tgt} {units}, [{lo}, {hi}] {units} - {notes}")
    else:
        lines.append("(none defined)")
    lines.append("")

    # Example data
    exs = test.get("example_data", [])
    lines.append("EXAMPLE DATA")
    lines.append("------------")
    if exs:
        for ex in exs:
            t = ex.get("type", "")
            d = ex.get("description", "")
            v = ex.get("data_value", "")
            f = ex.get("file_ref", "")
            notes = ex.get("notes", "")
            lines.append(f"- {t}: {d} value={v} file={f} - {notes}")
    else:
        lines.append("(none defined)")
    lines.append("")

    return "\n".join(lines)


class TestEditorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("TestBASE Plan Editor (v0)")
        self.resize(1400, 800)

        self.current_folder: Path | None = None
        self.current_path: Path | None = None
        self.current_test: dict = {}
        self._suppress_updates = False

        self._create_actions()
        self._create_menu()
        self._create_ui()

    # ----- Menu / actions -----

    def _create_actions(self):
        self.open_folder_act = QAction("Open Test Folder…", self)
        self.open_folder_act.triggered.connect(self.open_folder)

        self.save_act = QAction("Save", self)
        self.save_act.triggered.connect(self.save_current_test)

        self.exit_act = QAction("Exit", self)
        self.exit_act.triggered.connect(self.close)

    def _create_menu(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("&File")
        file_menu.addAction(self.open_folder_act)
        file_menu.addSeparator()
        file_menu.addAction(self.save_act)
        file_menu.addSeparator()
        file_menu.addAction(self.exit_act)

    # ----- UI -----

    def _create_ui(self):
        splitter = QSplitter(Qt.Orientation.Horizontal, self)

        # Left: test list
        self.test_list = QListWidget()
        self.test_list.itemSelectionChanged.connect(self.on_test_selection_changed)
        splitter.addWidget(self.test_list)
        splitter.setStretchFactor(0, 1)

        # Center: form + markdown preview
        center_tabs = QTabWidget()
        self._create_center_form(center_tabs)
        splitter.addWidget(center_tabs)
        splitter.setStretchFactor(1, 3)

        # Right: variables, keywords
        right_tabs = QTabWidget()
        self._create_right_panels(right_tabs)
        splitter.addWidget(right_tabs)
        splitter.setStretchFactor(2, 2)

        self.setCentralWidget(splitter)

    def _create_center_form(self, tabs: QTabWidget):
        # Form tab
        form_widget = QWidget()
        form_layout = QFormLayout(form_widget)

        self.field_edits = {}

        def add_line_field(key: str, label: str):
            edit = QLineEdit()
            self.field_edits[key] = edit
            edit.textChanged.connect(lambda text, k=key: self.on_field_changed(k, text))
            form_layout.addRow(label, edit)

        def add_text_field(key: str, label: str):
            edit = QTextEdit()
            self.field_edits[key] = edit
            edit.textChanged.connect(lambda k=key: self.on_field_changed(k, self.field_edits[k].toPlainText()))
            form_layout.addRow(label, edit)

        add_line_field("test_name", "Test Name")
        add_line_field("test_no", "Test No.")
        add_line_field("last_test_no", "Last Test No.")
        add_text_field("purpose", "Purpose")
        add_text_field("scope", "Scope")
        add_text_field("setup", "Setup")
        add_text_field("procedure", "Procedure")
        add_text_field("measurement", "Measurement")
        add_text_field("acceptancecriteria", "Acceptance Criteria")
        add_text_field("conclusion", "Conclusion")

        tabs.addTab(form_widget, "Form")

        # Markdown preview tab
        self.markdown_view = QPlainTextEdit()
        self.markdown_view.setReadOnly(True)
        tabs.addTab(self.markdown_view, "Markdown Preview")

    def _create_right_panels(self, tabs: QTabWidget):
        # Variables summary (read-only)
        var_widget = QWidget()
        vlayout = QVBoxLayout(var_widget)
        self.variables_summary = QPlainTextEdit()
        self.variables_summary.setReadOnly(True)
        vlayout.addWidget(self.variables_summary)
        tabs.addTab(var_widget, "Variables")

        # Keywords editor
        kw_widget = QWidget()
        kw_layout = QVBoxLayout(kw_widget)
        self.keywords_edit = QPlainTextEdit()
        self.keywords_edit.setPlaceholderText("One keyword per line…")
        self.keywords_edit.textChanged.connect(self.on_keywords_changed)
        kw_layout.addWidget(self.keywords_edit)
        tabs.addTab(kw_widget, "Keywords")

    # ----- Folder / file handling -----

    def open_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Test Folder")
        if not folder:
            return
        self.current_folder = Path(folder)
        self.load_tests_from_folder(self.current_folder)

    def load_tests_from_folder(self, folder: Path):
        self.test_list.clear()
        self.current_path = None
        self.current_test = {}

        for path in sorted(folder.glob("*.json")):
            try:
                with path.open("r", encoding="utf-8") as f:
                    data = json.load(f)
                # Heuristic: treat as test if it has test_name and test_no
                if "test_name" in data and "test_no" in data:
                    label = f"{data.get('test_no', '?')} – {data.get('test_name', 'Untitled')}"
                    item = QListWidgetItem(label)
                    item.setData(Qt.ItemDataRole.UserRole, str(path))
                    self.test_list.addItem(item)
            except Exception as e:
                print(f"Skipping {path}: {e}")

        if self.test_list.count() == 0:
            QMessageBox.information(self, "No Tests Found", "No test JSON files found in this folder.")
        else:
            self.test_list.setCurrentRow(0)

    def on_test_selection_changed(self):
        items = self.test_list.selectedItems()
        if not items:
            return
        item = items[0]
        path_str = item.data(Qt.ItemDataRole.UserRole)
        if not path_str:
            return
        path = Path(path_str)
        self.load_test_file(path)

    def load_test_file(self, path: Path):
        try:
            with path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load {path}:\n{e}")
            return

        self.current_path = path
        self.current_test = data
        self._suppress_updates = True

        # Populate fields
        for key, widget in self.field_edits.items():
            value = data.get(key, "") or ""
            if isinstance(widget, QLineEdit):
                widget.setText(value)
            elif isinstance(widget, QTextEdit):
                widget.setPlainText(value)

        # Keywords
        keywords = data.get("keywords", [])
        self.keywords_edit.setPlainText("\n".join(keywords))

        # Variables summary
        self.variables_summary.setPlainText(build_variables_summary(data))

        # Markdown preview
        self.markdown_view.setPlainText(build_markdown_from_test(data))

        self._suppress_updates = False

    def save_current_test(self):
        if not self.current_path or not self.current_test:
            QMessageBox.information(self, "Nothing to Save", "No test is currently loaded.")
            return
        try:
            with self.current_path.open("w", encoding="utf-8") as f:
                json.dump(self.current_test, f, indent=2, ensure_ascii=False)
            QMessageBox.information(self, "Saved", f"Saved {self.current_path.name}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save {self.current_path}:\n{e}")

    # ----- Data binding -----

    def on_field_changed(self, key: str, value: str):
        if self._suppress_updates:
            return
        self.current_test[key] = value
        # Live update of markdown
        self.markdown_view.setPlainText(build_markdown_from_test(self.current_test))

    def on_keywords_changed(self):
        if self._suppress_updates:
            return
        text = self.keywords_edit.toPlainText()
        keywords = [line.strip() for line in text.splitlines() if line.strip()]
        self.current_test["keywords"] = keywords


def main():
    app = QApplication(sys.argv)
    win = TestEditorWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
