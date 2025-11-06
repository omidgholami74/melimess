import sys
import pandas as pd
import random
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QLineEdit, QLabel, QCheckBox, QMessageBox, QSplitter
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QBrush

class DataProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Data Processor with PyQt6")
        self.setGeometry(100, 100, 1200, 800)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QHBoxLayout(self.central_widget)

        # Left panel for controls
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)

        # Buttons
        self.load_button = QPushButton("Load CSV/Excel File")
        self.load_button.clicked.connect(self.load_file)
        left_layout.addWidget(self.load_button)

        nav_layout = QHBoxLayout()
        self.prev_column_button = QPushButton("Previous Column")
        self.prev_column_button.clicked.connect(self.prev_column)
        self.prev_column_button.setEnabled(False)
        nav_layout.addWidget(self.prev_column_button)

        self.next_column_button = QPushButton("Next Column")
        self.next_column_button.clicked.connect(self.next_column)
        nav_layout.addWidget(self.next_column_button)
        self.next_column_button.setEnabled(False)
        left_layout.addLayout(nav_layout)

        # Control panel for filling
        control_layout = QVBoxLayout()
        control_layout.addWidget(QLabel("Fill Controls:"))
        min_max_layout = QHBoxLayout()
        min_max_layout.addWidget(QLabel("Min:"))
        self.min_edit = QLineEdit("0.9")
        min_max_layout.addWidget(self.min_edit)
        min_max_layout.addWidget(QLabel("Max:"))
        self.max_edit = QLineEdit("1.1")
        min_max_layout.addWidget(self.max_edit)
        control_layout.addLayout(min_max_layout)

        offset_ratio_layout = QHBoxLayout()
        offset_ratio_layout.addWidget(QLabel("Offset:"))
        self.offset_edit = QLineEdit("0")
        offset_ratio_layout.addWidget(self.offset_edit)
        offset_ratio_layout.addWidget(QLabel("Ratio:"))
        self.ratio_edit = QLineEdit("1.0")
        offset_ratio_layout.addWidget(self.ratio_edit)
        control_layout.addLayout(offset_ratio_layout)

        checkboxes_layout = QHBoxLayout()
        self.apply_filled_checkbox = QCheckBox("Apply to filled cells")
        checkboxes_layout.addWidget(self.apply_filled_checkbox)
        self.apply_ratio_checkbox = QCheckBox("Apply ratio to filled cells")
        checkboxes_layout.addWidget(self.apply_ratio_checkbox)
        control_layout.addLayout(checkboxes_layout)

        self.fill_button = QPushButton("Fill Empty Cells")
        self.fill_button.clicked.connect(self.fill_empty_cells)
        control_layout.addWidget(self.fill_button)
        left_layout.addLayout(control_layout)

        # Duplicate handling
        dup_layout = QVBoxLayout()
        dup_layout.addWidget(QLabel("Duplicate Handling:"))
        dup_range_layout = QHBoxLayout()
        dup_range_layout.addWidget(QLabel("Duplicate Range:"))
        self.dup_range_edit = QLineEdit("0.05")
        dup_range_layout.addWidget(self.dup_range_edit)
        dup_layout.addLayout(dup_range_layout)

        dup_buttons_layout = QHBoxLayout()
        self.check_dup_button = QPushButton("Check Duplicates")
        self.check_dup_button.clicked.connect(self.check_duplicates)
        dup_buttons_layout.addWidget(self.check_dup_button)

        self.fix_dup_button = QPushButton("Fix Duplicates")
        self.fix_dup_button.clicked.connect(self.fix_duplicates)
        dup_buttons_layout.addWidget(self.fix_dup_button)
        dup_layout.addLayout(dup_buttons_layout)
        left_layout.addLayout(dup_layout)

        # CRM handling
        crm_layout = QVBoxLayout()
        crm_layout.addWidget(QLabel("CRM Handling:"))
        crm_range_layout = QHBoxLayout()
        crm_range_layout.addWidget(QLabel("CRM Range:"))
        self.crm_range_edit = QLineEdit("0.1")
        crm_range_layout.addWidget(self.crm_range_edit)
        crm_layout.addLayout(crm_range_layout)

        crm_buttons_layout = QHBoxLayout()
        self.select_crm_button = QPushButton("Select CRM Row")
        self.select_crm_button.clicked.connect(self.select_crm_row)
        crm_buttons_layout.addWidget(self.select_crm_button)

        self.compare_crm_button = QPushButton("Compare with CRM 901")
        self.compare_crm_button.clicked.connect(self.compare_with_crm)
        crm_buttons_layout.addWidget(self.compare_crm_button)

        self.fix_crm_button = QPushButton("Fix CRM Differences")
        self.fix_crm_button.clicked.connect(self.fix_crm_differences)
        crm_buttons_layout.addWidget(self.fix_crm_button)
        crm_layout.addLayout(crm_buttons_layout)
        left_layout.addLayout(crm_layout)

        # Apply limits from row 4
        self.apply_limits_button = QPushButton("Apply Limits < >")
        self.apply_limits_button.clicked.connect(self.apply_limits)
        left_layout.addWidget(self.apply_limits_button)

        # Finalize
        self.finalize_button = QPushButton("Finalize and Save")
        self.finalize_button.clicked.connect(self.finalize_data)
        left_layout.addWidget(self.finalize_button)

        left_layout.addStretch()

        # Splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(left_panel)
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked | QTableWidget.EditTrigger.AnyKeyPressed)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ContiguousSelection)
        splitter.addWidget(self.table)
        splitter.setSizes([300, 900])

        main_layout.addWidget(splitter)

        # Data storage
        self.df = None
        self.reserved_rows = {}
        self.processed_columns = {}
        self.current_column_index = 0
        self.current_column_data = None
        self.fixed_column = None
        self.crm_row = None
        self.crm_901 = [18267.30, 11648.70, 11416.50, 11280.40, 11322.10, 10765.30, 9095.06, 6273.45, 8994.77, 9797.85, 9803.39, 9959.60, 10553.30, 10484.60, 10183.60, 11909.60, 10976.70, 10962.00, 12918.10, 10035.60, 9265.05, 11652.10, 12520.20, 12584.60, 11720.20, 10161.40, 10931.30, 10729.50, 10235.60, 10530.40, 6040.80, 13430.70]

        # Install event filter for global Ctrl+V
        self.installEventFilter(self)

    def eventFilter(self, source, event):
        if event.type() == event.Type.KeyPress:
            key_event = event
            if key_event.modifiers() == Qt.KeyboardModifier.ControlModifier and key_event.key() == Qt.Key.Key_V:
                self.paste_from_clipboard()
                return True
        return super().eventFilter(source, event)

    def paste_from_clipboard(self):
        clipboard = QApplication.clipboard()
        mime_data = clipboard.mimeData()
        if not mime_data.hasText():
            return

        text = mime_data.text()
        rows = [row.split('\t') for row in text.split('\n') if row.strip()]
        if not rows:
            return

        flat_values = []
        for row in rows:
            flat_values.extend([v.strip() for v in row if v.strip()])
        
        if not flat_values:
            return

        current = self.table.currentIndex()
        start_row = current.row() if current.isValid() and current.column() == 2 else 0
        col_idx = 2
        num_rows = self.table.rowCount()

        for i, val_str in enumerate(flat_values):
            row = start_row + i
            if row >= num_rows:
                break
            try:
                val = float(val_str)
            except ValueError:
                val = val_str
            item = self.table.item(row, col_idx)
            if not item:
                item = QTableWidgetItem()
                self.table.setItem(row, col_idx, item)
            item.setText(str(val))
            self.current_column_data.at[row, 'Modified'] = val

    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "CSV/Excel (*.csv *.xlsx)")
        if file_path:
            if file_path.endswith('.csv'):
                self.df = pd.read_csv(file_path, header=None)
            else:
                self.df = pd.read_excel(file_path, header=None)
            
            self.reserved_rows = {
                0: self.df.iloc[0].copy(),
                2: self.df.iloc[2].copy(),
                3: self.df.iloc[3].copy(),
                4: self.df.iloc[4].copy(),
                5: self.df.iloc[5].copy()
            }
            
            self.processing_df = self.df.iloc[6:].reset_index(drop=True)
            
            for col in self.processing_df.columns:
                self.processing_df[col] = self.processing_df[col].apply(self.clean_cell)
            
            self.fixed_column = self.processing_df.iloc[:, 0].values
            self.next_column_button.setEnabled(True)
            self.load_column(1)

    def clean_cell(self, cell):
        if isinstance(cell, str):
            if cell.startswith('<') or cell.startswith('>'):
                return float(cell[1:])
            elif cell.endswith('<') or cell.endswith('>'):
                return float(cell[:-1])
        return cell

    def load_column(self, col_index):
        if col_index < 1 or col_index >= len(self.processing_df.columns):
            return
        
        col_data = self.processing_df.iloc[:, col_index].values
        num_rows = len(self.fixed_column)
        
        modified = self.processed_columns.get(col_index, [None] * num_rows)
        self.current_column_data = pd.DataFrame({
            'Original': col_data,
            'Modified': modified
        })
        
        self.table.setRowCount(num_rows)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Fixed', 'Original', 'Modified'])
        
        for i in range(num_rows):
            fixed_item = QTableWidgetItem(str(self.fixed_column[i]))
            fixed_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(i, 0, fixed_item)
            
            orig_item = QTableWidgetItem(str(self.current_column_data.at[i, 'Original']))
            orig_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(i, 1, orig_item)
            
            mod_val = self.current_column_data.at[i, 'Modified']
            mod_item = QTableWidgetItem(str(mod_val) if mod_val is not None else "")
            mod_item.setFlags(Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(i, 2, mod_item)
        
        self.current_column_index = col_index
        self.prev_column_button.setEnabled(col_index > 1)
        self.next_column_button.setEnabled(col_index < len(self.processing_df.columns) - 1)

    def next_column(self):
        self.save_current_modified()
        self.load_column(self.current_column_index + 1)

    def prev_column(self):
        self.save_current_modified()
        self.load_column(self.current_column_index - 1)

    def save_current_modified(self):
        modified = []
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 2)
            text = item.text().strip() if item else ""
            try:
                val = float(text) if text else None
            except ValueError:
                val = text if text else None
            modified.append(val)
        self.processed_columns[self.current_column_index] = modified
        self.current_column_data['Modified'] = modified

    def fill_empty_cells(self):
        min_val = float(self.min_edit.text())
        max_val = float(self.max_edit.text())
        offset = float(self.offset_edit.text())
        ratio = float(self.ratio_edit.text())
        apply_filled = self.apply_filled_checkbox.isChecked()
        apply_ratio_filled = self.apply_ratio_checkbox.isChecked()

        for i in range(len(self.current_column_data)):
            original = self.current_column_data.at[i, 'Original']
            if pd.isna(original) or not isinstance(original, (int, float)):
                continue
            
            modified = self.current_column_data.at[i, 'Modified']
            if modified is None or apply_filled:
                rand_factor = random.uniform(min_val, max_val)
                new_val = (original * rand_factor) + offset
                if apply_ratio_filled or modified is None:
                    new_val *= ratio
                new_val = round(new_val, 2)
                self.current_column_data.at[i, 'Modified'] = new_val
                item = self.table.item(i, 2)
                if not item:
                    item = QTableWidgetItem()
                    self.table.setItem(i, 2, item)
                item.setText(str(new_val))

    def check_duplicates(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        selected_rows = set(item.row() for item in selected_items)
        
        values = [self.current_column_data.at[row, 'Original'] for row in selected_rows 
                 if self.current_column_data.at[row, 'Original'] is not None and isinstance(self.current_column_data.at[row, 'Original'], (int, float))]
        
        if not values:
            return
        
        mean_val = sum(values) / len(values)
        dup_range = float(self.dup_range_edit.text())

        light_yellow = QColor(255, 255, 150)
        light_red = QColor(255, 180, 180)

        # First: highlight all selected rows in light yellow
        for row in selected_rows:
            for col in [0, 1, 2]:
                item = self.table.item(row, col)
                if item:
                    item.setBackground(QBrush(light_yellow))

        # Then: mark outliers in red (only based on Original)
        for row in selected_rows:
            val = self.current_column_data.at[row, 'Original']
            if val is not None and isinstance(val, (int, float)) and abs(val - mean_val) > mean_val * dup_range:
                for col in [0, 1, 2]:
                    item = self.table.item(row, col)
                    if item:
                        item.setBackground(QBrush(light_red))

    def fix_duplicates(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        selected_rows = set(item.row() for item in selected_items)
        values = [self.current_column_data.at[row, 'Original'] for row in selected_rows 
                 if self.current_column_data.at[row, 'Original'] is not None and isinstance(self.current_column_data.at[row, 'Original'], (int, float))]
        if not values:
            return
        mean_val = sum(values) / len(values)
        
        min_val = float(self.min_edit.text())
        max_val = float(self.max_edit.text())
        
        # Only fix cells that are currently red
        light_red = QColor(255, 180, 180)
        light_green = QColor(180, 255, 180)  # Light green for fixed

        for row in selected_rows:
            # Check if row is red
            item = self.table.item(row, 1)  # Original column
            if item and item.background().color() == light_red:
                # Fix only Modified
                rand_factor = random.uniform(min_val, max_val)
                new_val = round(mean_val * rand_factor, 2)
                self.current_column_data.at[row, 'Modified'] = new_val
                mod_item = self.table.item(row, 2)
                if not mod_item:
                    mod_item = QTableWidgetItem()
                    self.table.setItem(row, 2, mod_item)
                mod_item.setText(str(new_val))
                # Highlight fixed row in light green
                for col in [0, 1, 2]:
                    bg_item = self.table.item(row, col)
                    if bg_item:
                        bg_item.setBackground(QBrush(light_green))
            else:
                # For yellow rows: copy Original to Modified (if Modified is empty or None)
                mod_val = self.current_column_data.at[row, 'Modified']
                orig_val = self.current_column_data.at[row, 'Original']
                if mod_val is None and orig_val is not None:
                    self.current_column_data.at[row, 'Modified'] = orig_val
                    mod_item = self.table.item(row, 2)
                    if not mod_item:
                        mod_item = QTableWidgetItem()
                        self.table.setItem(row, 2, mod_item)
                    mod_item.setText(str(orig_val))

    def select_crm_row(self):
        selected_items = self.table.selectedItems()
        if len(selected_items) != 1 or selected_items[0].column() != 2:
            QMessageBox.warning(self, "Error", "Select one modified cell as CRM.")
            return
        row = selected_items[0].row()
        self.crm_row = row
        QMessageBox.information(self, "Selected", f"CRM row selected: {row}")

    def compare_with_crm(self):
        if self.crm_row is None:
            QMessageBox.warning(self, "Error", "Select CRM row first.")
            return
        
        if self.current_column_index >= len(self.crm_901):
            return
        
        crm_901_val = self.crm_901[self.current_column_index]
        crm_range = float(self.crm_range_edit.text())
        
        for i in range(len(self.current_column_data)):
            val = self.current_column_data.at[i, 'Modified']
            if val is not None and abs(val - crm_901_val) > crm_901_val * crm_range:
                item = self.table.item(i, 2)
                item.setBackground(QBrush(QColor(255, 180, 180)))

    def fix_crm_differences(self):
        if self.crm_row is None:
            return
        
        if self.current_column_index >= len(self.crm_901):
            return
        
        crm_901_val = self.crm_901[self.current_column_index]
        min_val = float(self.min_edit.text())
        max_val = float(self.max_edit.text())
        
        for i in range(len(self.current_column_data)):
            item = self.table.item(i, 2)
            if item and item.background().color() == QColor(255, 180, 180):
                rand_factor = random.uniform(min_val, max_val)
                new_val = round(crm_901_val * rand_factor, 2)
                self.current_column_data.at[i, 'Modified'] = new_val
                item.setText(str(new_val))
                item.setBackground(QBrush(QColor("white")))

    def apply_limits(self):
        limit_row = self.reserved_rows[3]
        col_index = self.current_column_index
        limit = limit_row[col_index] if not pd.isna(limit_row[col_index]) else 0
        
        for i in range(len(self.current_column_data)):
            val = self.current_column_data.at[i, 'Modified']
            if val is not None and isinstance(val, (int, float)):
                if val < limit:
                    new_val = f"<{limit}"
                elif val > limit * 10:
                    new_val = f">{limit}"
                else:
                    new_val = val
                self.current_column_data.at[i, 'Modified'] = new_val
                item = self.table.item(i, 2)
                item.setText(str(new_val))

    def finalize_data(self):
        self.save_current_modified()
        if len(self.processed_columns) != (len(self.processing_df.columns) - 1):
            QMessageBox.warning(self, "Error", "Process all columns first.")
            return
        
        for col_index, col_data in self.processed_columns.items():
            self.processing_df.iloc[:, col_index] = col_data
        
        full_df = pd.DataFrame(columns=self.df.columns, index=range(len(self.df)))
        full_df.iloc[0] = self.reserved_rows[0]
        full_df.iloc[2] = self.reserved_rows[2]
        full_df.iloc[3] = self.reserved_rows[3]
        full_df.iloc[4] = self.reserved_rows[4]
        full_df.iloc[5] = self.reserved_rows[5]
        full_df.iloc[6:] = self.processing_df.values
        
        save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "CSV (*.csv)")
        if save_path:
            full_df.to_csv(save_path, index=False, header=False)
            QMessageBox.information(self, "Saved", "File saved successfully.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataProcessor()
    window.show()
    sys.exit(app.exec())