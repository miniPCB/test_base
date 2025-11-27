import sys
import json
import re
import os
import platform
import ctypes
from ctypes import wintypes
from pathlib import Path

from openai import OpenAI

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction, QKeySequence, QPalette, QColor
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
    QHBoxLayout,
    QFileDialog,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QMenu,
    QStyleFactory,
    QLabel,
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


# Column definitions for variable tables
TP_COLUMNS = ["name", "role", "schematic_ref", "net", "description"]
EQ_COLUMNS = ["id", "type", "model", "serial", "location", "notes"]
MS_COLUMNS = ["equipment_id", "mode", "range", "sampling", "other_settings"]
TH_COLUMNS = ["parameter", "target_value", "lower_limit", "upper_limit", "units", "notes"]
EX_COLUMNS = ["type", "description", "data_value", "file_ref", "notes"]
RH_COLUMNS = ["rev", "rev_date", "description", "rev_by"]


class TestEditorWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("TestBASE Plan Editor (v0)")
        self.resize(1600, 900)

        self.current_folder: Path | None = None
        self.current_path: Path | None = None  # None = new/unsaved test
        self.current_test: dict = {}
        self._suppress_updates = False

        # Remember which field AI is targeting
        # (kind, key, widget)
        self.ai_target_context = None

        self._create_actions()
        self._create_menu()
        self._create_ui()

    # ----- Menu / actions -----

    def _create_actions(self):
        self.open_folder_act = QAction("Open Test Folder…", self)
        self.open_folder_act.triggered.connect(self.open_folder)

        self.new_test_act = QAction("New Test", self)
        self.new_test_act.triggered.connect(self.new_test)

        self.save_act = QAction("Save", self)
        self.save_act.triggered.connect(self.save_current_test)
        self.save_act.setShortcut(QKeySequence.StandardKey.Save)

        self.exit_act = QAction("Exit", self)
        self.exit_act.triggered.connect(self.close)

        # AI: Alt + A
        self.ai_generate_act = QAction("Generate AI Content", self)
        self.ai_generate_act.setShortcut(QKeySequence("Alt+A"))
        self.ai_generate_act.triggered.connect(self.on_generate_ai_content)

    def _create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("&File")
        file_menu.addAction(self.open_folder_act)
        file_menu.addSeparator()
        file_menu.addAction(self.new_test_act)
        file_menu.addAction(self.save_act)
        file_menu.addSeparator()
        file_menu.addAction(self.exit_act)

        tools_menu = menubar.addMenu("&Tools")
        tools_menu.addAction(self.ai_generate_act)

    # ----- UI -----

    def _create_ui(self):
        splitter = QSplitter(Qt.Orientation.Horizontal, self)

        # Left: test list
        self.test_list = QListWidget()
        self.test_list.itemSelectionChanged.connect(self.on_test_selection_changed)
        # Context menu (right-click) for duplicate/delete
        self.test_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.test_list.customContextMenuRequested.connect(self.on_test_list_context_menu)
        splitter.addWidget(self.test_list)
        splitter.setStretchFactor(0, 1)

        # Center: form + markdown preview
        center_tabs = QTabWidget()
        self._create_center_form(center_tabs)
        splitter.addWidget(center_tabs)
        splitter.setStretchFactor(1, 3)

        # Right: revision history + variable editors + keywords + AI content
        right_tabs = QTabWidget()
        self._create_right_panels(right_tabs)
        splitter.addWidget(right_tabs)
        splitter.setStretchFactor(2, 3)

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
            edit.textChanged.connect(
                lambda k=key: self.on_field_changed(k, self.field_edits[k].toPlainText())
            )
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

    def _create_table_tab(
        self,
        parent_tabs: QTabWidget,
        title: str,
        columns: list[str],
        cell_changed_handler,
    ):
        """Helper to create a tab with a table and add/remove buttons."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        table = QTableWidget()
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.horizontalHeader().setStretchLastSection(True)
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QTableWidget.EditTrigger.AllEditTriggers)

        # Buttons
        btn_row = QHBoxLayout()
        add_btn = QPushButton("Add Row")
        del_btn = QPushButton("Delete Row")
        btn_row.addWidget(add_btn)
        btn_row.addWidget(del_btn)
        btn_row.addStretch()

        layout.addWidget(table)
        layout.addLayout(btn_row)

        parent_tabs.addTab(widget, title)

        # Connect signals
        table.cellChanged.connect(cell_changed_handler)
        add_btn.clicked.connect(
            lambda: self._add_row_to_table(table, columns, cell_changed_handler)
        )
        del_btn.clicked.connect(
            lambda: self._delete_row_from_table(table, columns, cell_changed_handler)
        )

        return table

    def _create_right_panels(self, tabs: QTabWidget):
        # Revision history
        self.rh_table = self._create_table_tab(
            tabs, "Revisions", RH_COLUMNS, self.on_rh_cell_changed
        )

        # Variable tables
        self.tp_table = self._create_table_tab(
            tabs, "Testpoints", TP_COLUMNS, self.on_tp_cell_changed
        )
        self.eq_table = self._create_table_tab(
            tabs, "Equipment", EQ_COLUMNS, self.on_eq_cell_changed
        )
        self.ms_table = self._create_table_tab(
            tabs, "Settings", MS_COLUMNS, self.on_ms_cell_changed
        )
        self.th_table = self._create_table_tab(
            tabs, "Thresholds", TH_COLUMNS, self.on_th_cell_changed
        )
        self.ex_table = self._create_table_tab(
            tabs, "Examples", EX_COLUMNS, self.on_ex_cell_changed
        )

        # Keywords editor
        kw_widget = QWidget()
        kw_layout = QVBoxLayout(kw_widget)
        self.keywords_edit = QPlainTextEdit()
        self.keywords_edit.setPlaceholderText("One keyword per line…")
        self.keywords_edit.textChanged.connect(self.on_keywords_changed)
        kw_layout.addWidget(self.keywords_edit)
        tabs.addTab(kw_widget, "Keywords")

        # AI Content editor
        ai_widget = QWidget()
        ai_layout = QVBoxLayout(ai_widget)

        self.ai_target_label = QLabel("Target: (none)")
        self.ai_content_edit = QPlainTextEdit()
        self.ai_content_edit.setPlaceholderText("AI suggestions will appear here…")

        btn_row = QHBoxLayout()
        self.ai_replace_btn = QPushButton("Replace Field with AI")
        self.ai_append_btn = QPushButton("Append to Field")
        btn_row.addWidget(self.ai_replace_btn)
        btn_row.addWidget(self.ai_append_btn)
        btn_row.addStretch()

        ai_layout.addWidget(self.ai_target_label)
        ai_layout.addWidget(self.ai_content_edit)
        ai_layout.addLayout(btn_row)
        tabs.addTab(ai_widget, "AI Content")

        self.ai_replace_btn.clicked.connect(self.on_ai_replace_clicked)
        self.ai_append_btn.clicked.connect(self.on_ai_append_clicked)

    # ----- Helpers for table row management -----

    def _add_row_to_table(self, table: QTableWidget, columns: list[str], handler):
        if self._suppress_updates:
            return
        self._suppress_updates = True
        row = table.rowCount()
        table.insertRow(row)
        for col in range(len(columns)):
            table.setItem(row, col, QTableWidgetItem(""))
        self._suppress_updates = False
        # Trigger a sync manually after adding a blank row
        handler(row, 0)

    def _delete_row_from_table(self, table: QTableWidget, columns: list[str], handler):
        if self._suppress_updates:
            return
        row = table.currentRow()
        if row < 0:
            return
        self._suppress_updates = True
        table.removeRow(row)
        self._suppress_updates = False
        # After deletion we resync whole corresponding array from the table
        if table is self.rh_table:
            self._sync_rh_from_table()
        elif table is self.tp_table:
            self._sync_tp_from_table()
        elif table is self.eq_table:
            self._sync_eq_from_table()
        elif table is self.ms_table:
            self._sync_ms_from_table()
        elif table is self.th_table:
            self._sync_th_from_table()
        elif table is self.ex_table:
            self._sync_ex_from_table()

    # ----- Helpers for naming / numbering -----

    def _suggest_next_test_no(self) -> str:
        """Look at the list labels and suggest the next numeric test number."""
        max_no = 0
        for i in range(self.test_list.count()):
            item = self.test_list.item(i)
            label = item.text()
            # Label format: "NNN – Name"
            no_part = label.split("–", 1)[0].strip()
            if no_part.isdigit():
                max_no = max(max_no, int(no_part))
        if max_no <= 0:
            return "001"
        return f"{max_no + 1:03d}"

    def _slugify_name(self, text: str) -> str:
        """Create a filesystem-friendly name fragment from test_name."""
        text = text.strip().lower()
        text = re.sub(r"\s+", "_", text)
        text = re.sub(r"[^a-z0-9_]+", "", text)
        return text or "new_test"

    def _add_or_update_list_item_for_current_test(self):
        """Ensure the left list has an item for current_path and update its label."""
        if not self.current_path:
            return
        path_str = str(self.current_path)
        label = f"{self.current_test.get('test_no', '?')} – {self.current_test.get('test_name', 'Untitled')}"
        # Try to find an existing item with this path
        for i in range(self.test_list.count()):
            item = self.test_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == path_str:
                item.setText(label)
                self.test_list.setCurrentItem(item)
                return
        # Not found → add new
        item = QListWidgetItem(label)
        item.setData(Qt.ItemDataRole.UserRole, path_str)
        self.test_list.addItem(item)
        self.test_list.setCurrentItem(item)

    def _clear_current_test_view(self):
        """Clear all form fields, tables, and preview."""
        self._suppress_updates = True
        for widget in self.field_edits.values():
            if isinstance(widget, QLineEdit):
                widget.clear()
            elif isinstance(widget, QTextEdit):
                widget.clear()
        self.keywords_edit.clear()
        for table in (
            self.rh_table,
            self.tp_table,
            self.eq_table,
            self.ms_table,
            self.th_table,
            self.ex_table,
        ):
            table.setRowCount(0)
        self.markdown_view.clear()
        self._suppress_updates = False

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
        self._clear_current_test_view()
        self._suppress_updates = True

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

        self._suppress_updates = False

        if self.test_list.count() == 0:
            QMessageBox.information(self, "No Tests Found", "No test JSON files found in this folder.")
        else:
            self.test_list.setCurrentRow(0)

    def on_test_selection_changed(self):
        if self._suppress_updates:
            return
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

        # Populate core fields
        for key, widget in self.field_edits.items():
            value = data.get(key, "") or ""
            if isinstance(widget, QLineEdit):
                widget.setText(value)
            elif isinstance(widget, QTextEdit):
                widget.setPlainText(value)

        # Keywords
        keywords = data.get("keywords", [])
        self.keywords_edit.setPlainText("\n".join(keywords))

        # Revision + variable tables
        self._populate_tables_from_test(data)

        # Markdown preview
        self.markdown_view.setPlainText(build_markdown_from_test(data))

        self._suppress_updates = False

    # ----- Context menu on test list -----

    def on_test_list_context_menu(self, pos):
        """Right-click context menu for duplicate/delete a test + file."""
        item = self.test_list.itemAt(pos)
        if not item:
            return

        menu = QMenu(self)
        dup_act = menu.addAction("Duplicate Test")
        del_act = menu.addAction("Delete Test")
        action = menu.exec(self.test_list.viewport().mapToGlobal(pos))
        if action == del_act:
            self._delete_test_item(item)
        elif action == dup_act:
            self._duplicate_test_item(item)

    def _delete_test_item(self, item: QListWidgetItem):
        path_str = item.data(Qt.ItemDataRole.UserRole)
        label = item.text()
        if not path_str:
            return
        path = Path(path_str)

        reply = QMessageBox.question(
            self,
            "Delete Test",
            f"Delete test:\n{label}\n\nThis will delete the JSON file:\n{path.name}\n\nAre you sure?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return

        # Delete JSON file
        try:
            if path.exists():
                path.unlink()
        except Exception as e:
            QMessageBox.warning(
                self,
                "Error",
                f"Failed to delete file:\n{path}\n\n{e}",
            )
            return

        # Remove from list
        row = self.test_list.row(item)
        self.test_list.takeItem(row)

        # If this was the current test, clear editor
        if self.current_path and path == self.current_path:
            self.current_path = None
            self.current_test = {}
            self._clear_current_test_view()

        # Select a neighbor, if any remain
        count = self.test_list.count()
        if count > 0:
            new_row = min(row, count - 1)
            self.test_list.setCurrentRow(new_row)

    def _duplicate_test_item(self, item: QListWidgetItem):
        """Duplicate the selected test into a new JSON file and load it."""
        path_str = item.data(Qt.ItemDataRole.UserRole)
        if not path_str or not self.current_folder:
            return
        src_path = Path(path_str)

        try:
            with src_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to read source test:\n{src_path}\n\n{e}")
            return

        # Deep copy
        new_test = json.loads(json.dumps(data))

        old_no = (new_test.get("test_no") or "").strip()
        new_no = self._suggest_next_test_no()
        new_test["last_test_no"] = old_no
        new_test["test_no"] = new_no

        # Append a revision note about duplication
        rh = new_test.get("revision_history")
        if not isinstance(rh, list):
            rh = []
        rh.append(
            {
                "rev": "-",
                "rev_date": "",
                "description": f"Duplicated from test {old_no or 'N/A'}",
                "rev_by": "",
            }
        )
        new_test["revision_history"] = rh

        # Determine new filename
        base_name = self._slugify_name(new_test.get("test_name", "new_test"))
        candidate = f"TB_{new_no}_{base_name}.json"
        dst_path = self.current_folder / candidate
        i = 2
        while dst_path.exists():
            candidate = f"TB_{new_no}_{base_name}_{i}.json"
            dst_path = self.current_folder / candidate
            i += 1

        # Write new file
        try:
            with dst_path.open("w", encoding="utf-8") as f:
                json.dump(new_test, f, indent=2, ensure_ascii=False)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to write duplicated test:\n{dst_path}\n\n{e}")
            return

        # Reload folder and select the new item
        self.load_tests_from_folder(self.current_folder)
        for i in range(self.test_list.count()):
            it = self.test_list.item(i)
            if it and it.data(Qt.ItemDataRole.UserRole) == str(dst_path):
                self.test_list.setCurrentItem(it)
                break

    # ----- New Test -----

    def new_test(self):
        """Create a new in-memory test using the TestBASE template."""
        if not self.current_folder:
            QMessageBox.information(
                self,
                "No Folder Open",
                "Please open a test folder first (File → Open Test Folder…).",
            )
            return

        next_no = self._suggest_next_test_no()
        new_test = {
            "test_name": "",
            "test_no": next_no,
            "last_test_no": "",
            "purpose": "",
            "scope": "",
            "setup": "",
            "procedure": "",
            "measurement": "",
            "acceptancecriteria": "",
            "conclusion": "",
            "revision_history": [
                {
                    "rev": "-",
                    "rev_date": "",
                    "description": "Initial release",
                    "rev_by": "",
                }
            ],
            "keywords": [],
            "testpoints": [],
            "measurement_equipment": [],
            "measurement_settings": [],
            "acceptance_thresholds": [],
            "example_data": [],
        }

        self.current_test = new_test
        self.current_path = None  # unsaved
        self._suppress_updates = True

        # Clear selection to show we are editing a new, unsaved test
        self.test_list.clearSelection()

        # Populate UI from template
        for key, widget in self.field_edits.items():
            value = new_test.get(key, "") or ""
            if isinstance(widget, QLineEdit):
                widget.setText(value)
            elif isinstance(widget, QTextEdit):
                widget.setPlainText(value)

        self.keywords_edit.setPlainText("")
        self._populate_tables_from_test(new_test)
        self.markdown_view.setPlainText(build_markdown_from_test(new_test))

        self._suppress_updates = False

    # ----- Save -----

    def save_current_test(self):
        if not self.current_test:
            QMessageBox.information(self, "Nothing to Save", "No test is currently loaded.")
            return

        if not self.current_folder:
            QMessageBox.information(
                self,
                "No Folder",
                "Please open a test folder first (File → Open Test Folder…).",
            )
            return

        # Sync arrays from tables into the dict
        self._sync_rh_from_table()
        self._sync_tp_from_table()
        self._sync_eq_from_table()
        self._sync_ms_from_table()
        self._sync_th_from_table()
        self._sync_ex_from_table()

        # If this is a new/unsaved test, choose a filename
        if self.current_path is None:
            test_no = (self.current_test.get("test_no") or "").strip()
            if not test_no.isdigit():
                test_no = self._suggest_next_test_no()
                self.current_test["test_no"] = test_no
            base_name = self._slugify_name(self.current_test.get("test_name", "new_test"))
            candidate = f"TB_{test_no}_{base_name}.json"
            path = self.current_folder / candidate

            i = 2
            while path.exists():
                candidate = f"TB_{test_no}_{base_name}_{i}.json"
                path = self.current_folder / candidate
                i += 1

            self.current_path = path

        try:
            with self.current_path.open("w", encoding="utf-8") as f:
                json.dump(self.current_test, f, indent=2, ensure_ascii=False)
            self._add_or_update_list_item_for_current_test()
            QMessageBox.information(self, "Saved", f"Saved {self.current_path.name}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save {self.current_path}:\n{e}")

    # ----- Data binding for core + keywords -----

    def on_field_changed(self, key: str, value: str):
        if self._suppress_updates or not self.current_test:
            return
        self.current_test[key] = value
        # Live update of markdown
        self.markdown_view.setPlainText(build_markdown_from_test(self.current_test))

    def on_keywords_changed(self):
        if self._suppress_updates or not self.current_test:
            return
        text = self.keywords_edit.toPlainText()
        keywords = [line.strip() for line in text.splitlines() if line.strip()]
        self.current_test["keywords"] = keywords

    # ----- Populate tables from test dict -----

    def _populate_tables_from_test(self, test: dict):
        # Helper to fill one table from an array of dicts
        def fill_table(table: QTableWidget, cols: list[str], rows: list[dict]):
            table.blockSignals(True)
            table.setRowCount(0)
            for row_data in rows:
                row = table.rowCount()
                table.insertRow(row)
                for c, key in enumerate(cols):
                    value = row_data.get(key, "")
                    table.setItem(row, c, QTableWidgetItem(str(value)))
            table.blockSignals(False)

        fill_table(self.rh_table, RH_COLUMNS, test.get("revision_history", []))
        fill_table(self.tp_table, TP_COLUMNS, test.get("testpoints", []))
        fill_table(self.eq_table, EQ_COLUMNS, test.get("measurement_equipment", []))
        fill_table(self.ms_table, MS_COLUMNS, test.get("measurement_settings", []))
        fill_table(self.th_table, TH_COLUMNS, test.get("acceptance_thresholds", []))
        fill_table(self.ex_table, EX_COLUMNS, test.get("example_data", []))

    # ----- Sync helpers: tables -> dict -----

    def _sync_array_from_table(self, table: QTableWidget, cols: list[str]) -> list[dict]:
        rows_data: list[dict] = []
        for row in range(table.rowCount()):
            row_dict: dict = {}
            empty = True
            for c, key in enumerate(cols):
                item = table.item(row, c)
                text = item.text().strip() if item is not None else ""
                if text:
                    empty = False
                row_dict[key] = text
            if not empty:
                rows_data.append(row_dict)
        return rows_data

    def _sync_rh_from_table(self):
        if not self.current_test:
            return
        self.current_test["revision_history"] = self._sync_array_from_table(
            self.rh_table, RH_COLUMNS
        )

    def _sync_tp_from_table(self):
        if not self.current_test:
            return
        self.current_test["testpoints"] = self._sync_array_from_table(
            self.tp_table, TP_COLUMNS
        )

    def _sync_eq_from_table(self):
        if not self.current_test:
            return
        self.current_test["measurement_equipment"] = self._sync_array_from_table(
            self.eq_table, EQ_COLUMNS
        )

    def _sync_ms_from_table(self):
        if not self.current_test:
            return
        self.current_test["measurement_settings"] = self._sync_array_from_table(
            self.ms_table, MS_COLUMNS
        )

    def _sync_th_from_table(self):
        if not self.current_test:
            return
        self.current_test["acceptance_thresholds"] = self._sync_array_from_table(
            self.th_table, TH_COLUMNS
        )

    def _sync_ex_from_table(self):
        if not self.current_test:
            return
        self.current_test["example_data"] = self._sync_array_from_table(
            self.ex_table, EX_COLUMNS
        )

    # ----- Handlers for cell changes (tables -> dict, live) -----

    def on_rh_cell_changed(self, row: int, col: int):
        if self._suppress_updates or not self.current_test:
            return
        self._sync_rh_from_table()

    def on_tp_cell_changed(self, row: int, col: int):
        if self._suppress_updates or not self.current_test:
            return
        self._sync_tp_from_table()

    def on_eq_cell_changed(self, row: int, col: int):
        if self._suppress_updates or not self.current_test:
            return
        self._sync_eq_from_table()

    def on_ms_cell_changed(self, row: int, col: int):
        if self._suppress_updates or not self.current_test:
            return
        self._sync_ms_from_table()

    def on_th_cell_changed(self, row: int, col: int):
        if self._suppress_updates or not self.current_test:
            return
        self._sync_th_from_table()

    def on_ex_cell_changed(self, row: int, col: int):
        if self._suppress_updates or not self.current_test:
            return
        self._sync_ex_from_table()

    # ----- AI integration -----

    def _get_active_field(self):
        """Figure out which field currently has focus."""
        w = self.focusWidget()
        if not w:
            return None

        # Core form fields
        for key, widget in self.field_edits.items():
            if w is widget:
                return ("core", key, widget)

        # Keywords editor
        if w is self.keywords_edit:
            return ("keywords", "keywords", self.keywords_edit)

        # You could later add support for table cells here

        return None

    def on_generate_ai_content(self):
        """Triggered by Alt + A or Tools → Generate AI Content."""
        if not self.current_test:
            QMessageBox.information(
                self, "No Test Loaded", "Open or create a test first."
            )
            return

        ctx = self._get_active_field()
        if not ctx:
            QMessageBox.information(
                self,
                "No Field Selected",
                "Click into a text field (e.g. Purpose, Scope, etc.) and try again.",
            )
            return

        kind, key, widget = ctx
        self.ai_target_context = ctx
        self.ai_target_label.setText(f"Target: {key}")

        # Get current value of the field
        if isinstance(widget, QLineEdit):
            current_val = widget.text().strip()
        elif isinstance(widget, (QTextEdit, QPlainTextEdit)):
            current_val = widget.toPlainText().strip()
        else:
            current_val = ""

        test_json = json.dumps(self.current_test, indent=2)

        prompt = (
            "You are helping complete an engineering hardware test plan for a circuit.\n"
            "I will give you the current JSON for a single test, and the field I want you to help with.\n"
            "Return ONLY suggested text for that single field. No markdown, no JSON, no extra explanation.\n\n"
            f"Field to write: {key}\n"
            f"Current value (may be empty): {current_val!r}\n\n"
            f"Test JSON:\n{test_json}\n"
        )

        self.ai_content_edit.setPlainText("Generating AI content…")

        try:
            api_key = os.getenv("OPENAI_API_KEY")
            if not api_key:
                raise RuntimeError(
                    "OPENAI_API_KEY environment variable is not set."
                )

            # New-style client (openai>=1.0)
            client = OpenAI(api_key=api_key)

            response = client.chat.completions.create(
                model="gpt-4o-mini",
                messages=[
                    {
                        "role": "system",
                        "content": "You write concise, technically accurate test plan text.",
                    },
                    {"role": "user", "content": prompt},
                ],
                temperature=0.2,
            )

            text = response.choices[0].message.content.strip()
            self.ai_content_edit.setPlainText(text)
        except Exception as e:
            self.ai_content_edit.setPlainText("")
            QMessageBox.warning(
                self, "AI Error", f"Failed to generate AI content:\n{e}"
            )


    def on_ai_replace_clicked(self):
        """Replace the selected field with the AI suggestion."""
        if not self.ai_target_context:
            QMessageBox.information(
                self,
                "No Target",
                "Use Alt+A in a field first to choose a target.",
            )
            return

        kind, key, widget = self.ai_target_context
        text = self.ai_content_edit.toPlainText().strip()
        if not text:
            return

        self._suppress_updates = True
        if isinstance(widget, QLineEdit):
            widget.setText(text)
        elif isinstance(widget, (QTextEdit, QPlainTextEdit)):
            widget.setPlainText(text)
        self._suppress_updates = False

        # Update backing dict
        if kind == "core":
            self.on_field_changed(key, text)
        elif kind == "keywords":
            self.keywords_edit.setPlainText(text)
            self.on_keywords_changed()

    def on_ai_append_clicked(self):
        """Append the AI suggestion to the selected field."""
        if not self.ai_target_context:
            QMessageBox.information(
                self,
                "No Target",
                "Use Alt+A in a field first to choose a target.",
            )
            return

        kind, key, widget = self.ai_target_context
        text = self.ai_content_edit.toPlainText().strip()
        if not text:
            return

        self._suppress_updates = True
        if isinstance(widget, QLineEdit):
            current = widget.text().strip()
            if current:
                widget.setText(current + " " + text)
            else:
                widget.setText(text)
            new_val = widget.text()
        elif isinstance(widget, (QTextEdit, QPlainTextEdit)):
            current = widget.toPlainText().rstrip()
            if current:
                widget.setPlainText(current + "\n\n" + text)
            else:
                widget.setPlainText(text)
            new_val = widget.toPlainText()
        else:
            new_val = text
        self._suppress_updates = False

        if kind == "core":
            self.on_field_changed(key, new_val)
        elif kind == "keywords":
            self.keywords_edit.setPlainText(new_val)
            self.on_keywords_changed()


# ----- Dark mode helpers -----

def apply_dark_palette(app: QApplication):
    """Apply a dark theme to the whole application."""
    app.setStyle(QStyleFactory.create("Fusion"))

    palette = QPalette()
    # Try to align with Windows dark title bar
    bg = QColor(32, 32, 32)   # Window background
    base = QColor(24, 24, 24) # Text/edit background

    palette.setColor(QPalette.ColorRole.Window, bg)
    palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Base, base)
    palette.setColor(QPalette.ColorRole.AlternateBase, bg)
    palette.setColor(QPalette.ColorRole.ToolTipBase, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.ToolTipText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.Button, bg)
    palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
    palette.setColor(QPalette.ColorRole.BrightText, Qt.GlobalColor.red)
    palette.setColor(QPalette.ColorRole.Link, QColor(90, 160, 255))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(90, 160, 255))
    palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)

    app.setPalette(palette)


def enable_windows_dark_titlebar(window: QWidget):
    """Ask Windows 10/11 to use a dark title bar for this window."""
    if platform.system() != "Windows":
        return

    try:
        hwnd = wintypes.HWND(int(window.winId()))

        # Windows 10 1809+ / 11: DWMWA_USE_IMMERSIVE_DARK_MODE
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20

        value = ctypes.c_int(1)
        dwmapi = ctypes.windll.dwmapi

        # Try attribute 20, fall back to 19 if needed
        res = dwmapi.DwmSetWindowAttribute(
            hwnd,
            ctypes.c_int(DWMWA_USE_IMMERSIVE_DARK_MODE),
            ctypes.byref(value),
            ctypes.sizeof(value),
        )
        if res != 0:
            DWMWA_USE_IMMERSIVE_DARK_MODE = 19
            dwmapi.DwmSetWindowAttribute(
                hwnd,
                ctypes.c_int(DWMWA_USE_IMMERSIVE_DARK_MODE),
                ctypes.byref(value),
                ctypes.sizeof(value),
            )
    except Exception:
        # If anything fails, just ignore and keep running
        pass


def main():
    app = QApplication(sys.argv)
    apply_dark_palette(app)

    win = TestEditorWindow()
    win.show()
    enable_windows_dark_titlebar(win)

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
