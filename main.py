import sys
import pandas as pd
import random
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QLineEdit, QLabel, QCheckBox, QMessageBox, QSplitter,
    QGroupBox, QFormLayout, QDoubleSpinBox, QStatusBar, QComboBox
)
from PyQt6.QtCore import Qt, QEvent
from PyQt6.QtGui import QColor, QBrush, QFont
class DataProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Data Processor with PyQt6")
        self.setGeometry(100, 100, 1200, 800)
        # Set global stylesheet for better UI
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QLabel {
                font-size: 12px;
                color: #333;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 6px 12px;
                border: none;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
            QLineEdit, QDoubleSpinBox {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 4px;
                background-color: white;
            }
            QCheckBox {
                color: #333;
            }
            QTableWidget {
                background-color: white;
                alternate-background-color: #f9f9f9;
                gridline-color: #ddd;
                selection-background-color: #a8d1ff;
            }
            QTableWidget::item {
                padding: 4px;
            }
            QHeaderView::section {
                background-color: #e0e0e0;
                padding: 4px;
                border: 1px solid #ccc;
                font-weight: bold;
            }
            QGroupBox {
                border: 1px solid #ccc;
                border-radius: 6px;
                margin-top: 10px;
                padding: 10px;
                font-weight: bold;
                color: #2c3e50;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
            }
            QComboBox {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 4px;
                background-color: white;
            }
        """)
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QHBoxLayout(self.central_widget)
        # Left panel for controls
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setSpacing(10)
        # File Group
        file_group = QGroupBox("File Operations")
        file_layout = QVBoxLayout(file_group)
        self.load_button = QPushButton("Load CSV/Excel File")
        self.load_button.clicked.connect(self.load_file)
        self.load_button.setToolTip("Load a CSV or Excel file to process")
        file_layout.addWidget(self.load_button)
        left_layout.addWidget(file_group)
        # Navigation Group
        nav_group = QGroupBox("Column Navigation")
        nav_layout = QVBoxLayout(nav_group)
        nav_buttons_layout = QHBoxLayout()
        self.prev_column_button = QPushButton("Previous Column")
        self.prev_column_button.clicked.connect(self.prev_column)
        self.prev_column_button.setEnabled(False)
        nav_buttons_layout.addWidget(self.prev_column_button)
        self.next_column_button = QPushButton("Next Column")
        self.next_column_button.clicked.connect(self.next_column)
        nav_buttons_layout.addWidget(self.next_column_button)
        self.next_column_button.setEnabled(False)
        nav_layout.addLayout(nav_buttons_layout)
        # Current element display
        self.element_label = QLabel("Element: -")
        self.element_label.setStyleSheet("font-weight: bold; font-size: 14px; color: #2c3e50;")
        nav_layout.addWidget(self.element_label)
        left_layout.addWidget(nav_group)
        # Fill Controls Group
        fill_group = QGroupBox("Fill Empty Cells")
        fill_layout = QFormLayout(fill_group)
        fill_layout.setSpacing(6)
        # Min/Max
        min_max_layout = QHBoxLayout()
        self.min_spin = QDoubleSpinBox()
        self.min_spin.setValue(0.9)
        self.min_spin.setMinimum(0.0)
        self.min_spin.setMaximum(10.0)
        self.min_spin.setSingleStep(0.1)
        min_max_layout.addWidget(QLabel("Min:"))
        min_max_layout.addWidget(self.min_spin)
        self.max_spin = QDoubleSpinBox()
        self.max_spin.setValue(1.1)
        self.max_spin.setMinimum(0.0)
        self.max_spin.setMaximum(10.0)
        self.max_spin.setSingleStep(0.1)
        min_max_layout.addWidget(QLabel("Max:"))
        min_max_layout.addWidget(self.max_spin)
        fill_layout.addRow(min_max_layout)
        # Offset/Ratio
        offset_ratio_layout = QHBoxLayout()
        self.offset_spin = QDoubleSpinBox()
        self.offset_spin.setValue(0.0)
        self.offset_spin.setMinimum(-1000.0)
        self.offset_spin.setMaximum(1000.0)
        self.offset_spin.setSingleStep(1.0)
        offset_ratio_layout.addWidget(QLabel("Offset:"))
        offset_ratio_layout.addWidget(self.offset_spin)
        self.ratio_spin = QDoubleSpinBox()
        self.ratio_spin.setValue(1.0)
        self.ratio_spin.setMinimum(0.0)
        self.ratio_spin.setMaximum(10.0)
        self.ratio_spin.setSingleStep(0.1)
        offset_ratio_layout.addWidget(QLabel("Ratio:"))
        offset_ratio_layout.addWidget(self.ratio_spin)
        fill_layout.addRow(offset_ratio_layout)
        # Checkboxes
        checkboxes_layout = QHBoxLayout()
        self.apply_filled_checkbox = QCheckBox("Apply to filled cells")
        checkboxes_layout.addWidget(self.apply_filled_checkbox)
        self.apply_ratio_checkbox = QCheckBox("Apply ratio to filled cells")
        checkboxes_layout.addWidget(self.apply_ratio_checkbox)
        fill_layout.addRow(checkboxes_layout)
        self.fill_button = QPushButton("Fill Empty Cells")
        self.fill_button.clicked.connect(self.fill_empty_cells)
        self.fill_button.setToolTip("Fill empty or selected cells with random values in range")
        fill_layout.addRow(self.fill_button)
        left_layout.addWidget(fill_group)
        # Global Operations Group (initially disabled)
        self.global_group = QGroupBox("Global Duplicate & CRM Handling")
        self.global_group.setEnabled(False)
        global_layout = QVBoxLayout(self.global_group)
       
        # Column selector
        column_select_layout = QHBoxLayout()
        self.column_combo = QComboBox()
        self.column_combo.setToolTip("Select a column to apply operations")
        column_select_layout.addWidget(QLabel("Select Column:"))
        column_select_layout.addWidget(self.column_combo)
        global_layout.addLayout(column_select_layout)
        # Duplicate Handling for selected column
        global_dup_group = QGroupBox("Duplicate Handling")
        global_dup_layout = QFormLayout(global_dup_group)
        self.global_dup_range_spin = QDoubleSpinBox()
        self.global_dup_range_spin.setValue(0.05)
        self.global_dup_range_spin.setMinimum(0.0)
        self.global_dup_range_spin.setMaximum(1.0)
        self.global_dup_range_spin.setSingleStep(0.01)
        global_dup_layout.addRow("Duplicate Range:", self.global_dup_range_spin)
        global_dup_buttons_layout = QHBoxLayout()
        self.global_check_dup_button = QPushButton("Check Duplicates")
        self.global_check_dup_button.clicked.connect(self.global_check_duplicates)
        self.global_check_dup_button.setToolTip("Highlight duplicates in selected rows for selected column")
        global_dup_buttons_layout.addWidget(self.global_check_dup_button)
        self.global_fix_dup_button = QPushButton("Fix Duplicates")
        self.global_fix_dup_button.clicked.connect(self.global_fix_duplicates)
        self.global_fix_dup_button.setToolTip("Fix highlighted duplicates with average-based values for selected column")
        global_dup_buttons_layout.addWidget(self.global_fix_dup_button)
        global_dup_layout.addRow(global_dup_buttons_layout)
        global_layout.addWidget(global_dup_group)
        # CRM Handling for selected column
        global_crm_group = QGroupBox("CRM Handling")
        global_crm_layout = QFormLayout(global_crm_group)
        self.global_crm_range_spin = QDoubleSpinBox()
        self.global_crm_range_spin.setValue(0.1)
        self.global_crm_range_spin.setMinimum(0.0)
        self.global_crm_range_spin.setMaximum(1.0)
        self.global_crm_range_spin.setSingleStep(0.01)
        global_crm_layout.addRow("CRM Range:", self.global_crm_range_spin)
        global_crm_buttons_layout = QVBoxLayout()
        self.global_compare_crm_button = QPushButton("Compare with CRM 903")
        self.global_compare_crm_button.clicked.connect(self.global_compare_with_crm)
        self.global_compare_crm_button.setToolTip("Compare selected row with CRM 903 value for selected column")
        global_crm_buttons_layout.addWidget(self.global_compare_crm_button)
        self.global_fix_crm_button = QPushButton("Fix CRM Differences")
        self.global_fix_crm_button.clicked.connect(self.global_fix_crm_differences)
        self.global_fix_crm_button.setToolTip("Adjust selected CRM row to match within range for selected column")
        global_crm_buttons_layout.addWidget(self.global_fix_crm_button)
        self.global_clear_crm_button = QPushButton("Clear CRM Row")
        self.global_clear_crm_button.clicked.connect(self.global_clear_crm_row)
        self.global_clear_crm_button.setEnabled(False)
        self.global_clear_crm_button.setToolTip("Remove CRM reference row and clear highlights for selected column")
        global_crm_buttons_layout.addWidget(self.global_clear_crm_button)
        global_crm_layout.addRow(global_crm_buttons_layout)
        global_layout.addWidget(global_crm_group)
        left_layout.addWidget(self.global_group)
        # Actions Group
        actions_group = QGroupBox("Actions")
        actions_layout = QVBoxLayout(actions_group)
        self.apply_limits_button = QPushButton("Apply Limits to All Columns")
        self.apply_limits_button.clicked.connect(self.apply_limits)
        self.apply_limits_button.setToolTip("Apply limits from row 4 to modified values in all columns")
        self.apply_limits_button.setEnabled(False)
        actions_layout.addWidget(self.apply_limits_button)
        self.finalize_button = QPushButton("Finalize and Save")
        self.finalize_button.clicked.connect(self.finalize_data)
        self.finalize_button.setToolTip("Save processed data to Excel")
        actions_layout.addWidget(self.finalize_button)
        left_layout.addWidget(actions_group)
        left_layout.addStretch()
        # Splitter
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(left_panel)
        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.EditTrigger.DoubleClicked | QTableWidget.EditTrigger.AnyKeyPressed)
        self.table.setSelectionMode(QTableWidget.SelectionMode.ContiguousSelection)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setVisible(False)
        splitter.addWidget(self.table)
        splitter.setSizes([350, 850]) # Slightly wider left panel for better UI
        main_layout.addWidget(splitter)
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")
        # Data storage
        self.df = None
        self.header_row = None # Row 1: element names (هدر اصلی)
        self.reserved_rows = {}
        self.processed_columns = {}
        self.current_column_index = 0
        self.current_column_data = None
        self.fixed_column = None
        self.crm_row = None
        self.crm_reference_row = None
        self.all_processed_mode = False
        # CRM 903 - OREAS 903 values by element name
        self.crm_903 = {
            'Ag': 0.432001, 'Al': 58903.5, 'As': 49.6558, 'Au': 0.00495, 'Ba': 197.464,
            'Be': 4.42199, 'Bi': 8.94184, 'Ca': 6250.92, 'Cd': 0.203073, 'Ce': 82.2075,
            'Co': 130.622, 'Cr': 72.8517, 'Cs': 3.56726, 'Cu': 6516.13, 'Cu-Sol(H2SO4)': 4340.59,
            'Fe': 41573.8, 'Ga': 15.0037, 'Ge': 0.0976435, 'Hf': 4.55988, 'In': 0.162135,
            'K': 33078.5, 'La': 40.2266, 'Li': 18.3245, 'Lu': 0.364898, 'Mg': 7139.92,
            'Mn': 689.554, 'Mo': 4.31947, 'Na': 300.694, 'Ni': 53.9221, 'P': 1068.18,
            'Pb': 11.2861, 'Rb': 136.573, 'S': 4996.88, 'Sb': 1.57009, 'Sc': 10.2376,
            'Se': 6.06318, 'Sn': 2.62936, 'Sr': 77.1271, 'Ta': 0.536122, 'Tb': 0.834551,
            'Te': 0.0344281, 'Th': 13.643, 'Ti': 1924.92, 'Tl': 0.621858, 'U': 7.58033,
            'V': 73.9117, 'W': 0.531139, 'Y': 22.4734, 'Yb': 2.36446, 'Zn': 24.2974, 'Zr': 151.863
        }
        # Install event filter for global Ctrl+V
        self.installEventFilter(self)
    def eventFilter(self, source, event):
        if event.type() == QEvent.Type.KeyPress:
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
        if not current.isValid():
            return
        start_row = current.row()
        if self.all_processed_mode:
            start_col = current.column()
            if start_col < 1:
                return
            col_index = start_col
        else:
            start_col = 2
            col_index = self.current_column_index
        num_rows = self.table.rowCount()
        for i, val_str in enumerate(flat_values):
            row = start_row + i
            if row >= num_rows:
                break
            try:
                val = float(val_str)
            except ValueError:
                val = val_str
            item = self.table.item(row, start_col)
            if not item:
                item = QTableWidgetItem()
                self.table.setItem(row, start_col, item)
            item.setText(str(val))
            actual_row = row
            if self.crm_reference_row is not None and row > self.crm_reference_row:
                actual_row -= 1
            if actual_row < len(self.processed_columns.get(col_index, [])):
                self.processed_columns[col_index][actual_row] = val
        self.status_bar.showMessage("Pasted from clipboard")
    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "CSV/Excel (*.csv *.xlsx)")
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    self.df = pd.read_csv(file_path, header=None)
                else:
                    self.df = pd.read_excel(file_path, header=None)
               
                # Row 1: element names (هدر اصلی)
                self.header_row = self.df.iloc[1].copy()
               
                # Reserved rows: 2,3,4,5
                self.reserved_rows = {
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
                self.status_bar.showMessage(f"Loaded file: {file_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to load file: {str(e)}")
                self.status_bar.showMessage("Error loading file")
    def clean_cell(self, cell):
        if isinstance(cell, str):
            if cell.startswith('<') or cell.startswith('>'):
                return float(cell[1:])
            elif cell.endswith('<') or cell.endswith('>'):
                return float(cell[:-1])
        return cell
    def get_current_element_name(self):
        """دریافت نام عنصر از سطر 1 (هدر)"""
        if self.header_row is None or self.current_column_index >= len(self.header_row):
            return "Unknown"
        name = str(self.header_row.iloc[self.current_column_index]).strip()
        return name if name else "Unknown"
    def get_element_name(self, col_index):
        if self.header_row is None or col_index >= len(self.header_row):
            return "Unknown"
        name = str(self.header_row.iloc[col_index]).strip()
        return name if name else "Unknown"
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
       
        self.remove_crm_reference_row()
       
        self.table.setRowCount(num_rows)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(['Fixed', 'Original', 'Modified'])
       
        for i in range(num_rows):
            fixed_item = QTableWidgetItem(str(self.fixed_column[i]))
            fixed_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(i, 0, fixed_item)
           
            orig_val = self.current_column_data.at[i, 'Original']
            if isinstance(orig_val, (int, float)) and not pd.isna(orig_val):
                if orig_val == int(orig_val): # عدد صحیح است
                    orig_text = str(int(orig_val))
                else: # اعشار دارد
                    orig_text = f"{orig_val:.2f}"
            else:
                orig_text = "" if pd.isna(orig_val) else str(orig_val)
            orig_item = QTableWidgetItem(orig_text)
           
            orig_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(i, 1, orig_item)
           
            mod_val = self.current_column_data.at[i, 'Modified']
            mod_item = QTableWidgetItem(str(mod_val) if mod_val is not None else "")
            mod_item.setFlags(Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(i, 2, mod_item)
       
        self.current_column_index = col_index
        self.update_navigation_buttons()
       
        # نمایش نام عنصر
        element_name = self.get_current_element_name()
        self.element_label.setText(f"Element: {element_name}")
        self.status_bar.showMessage(f"Loaded column {col_index}: {element_name}")
    def update_navigation_buttons(self):
        if self.all_processed_mode:
            self.prev_column_button.setEnabled(False)
            self.next_column_button.setEnabled(False)
        else:
            self.prev_column_button.setEnabled(self.current_column_index > 1)
            self.next_column_button.setEnabled(True)
    def next_column(self):
        self.save_current_modified()
        next_index = self.current_column_index + 1
        if next_index < len(self.processing_df.columns):
            self.load_column(next_index)
        else:
            self.check_all_columns_processed()
    def prev_column(self):
        self.save_current_modified()
        self.load_column(self.current_column_index - 1)
    def save_current_modified(self):
        if self.all_processed_mode:
            return
        modified = []
        for i in range(self.table.rowCount()):
            if self.crm_reference_row is not None and i == self.crm_reference_row:
                continue
            item = self.table.item(i, 2)
            text = item.text().strip() if item else ""
            try:
                val = float(text) if text else None
            except ValueError:
                val = text if text else None
            modified.append(val)
        self.processed_columns[self.current_column_index] = modified
        self.current_column_data['Modified'] = modified
    def check_all_columns_processed(self):
        num_columns = len(self.processing_df.columns) - 1 # excluding fixed column
        if len(self.processed_columns) == num_columns:
            self.global_group.setEnabled(True)
            self.apply_limits_button.setEnabled(True)
            self.column_combo.clear()
            for col_index in range(1, len(self.processing_df.columns)):
                element_name = self.get_element_name(col_index)
                self.column_combo.addItem(element_name, col_index)
            self.load_all_processed()
            self.all_processed_mode = True
            self.update_navigation_buttons()
            self.element_label.setText("All Elements")
            self.status_bar.showMessage("All columns processed. Showing all modified columns.")
        else:
            self.global_group.setEnabled(False)
            self.apply_limits_button.setEnabled(False)
    def load_all_processed(self):
        self.table.clear()
        num_rows = len(self.fixed_column)
        num_cols = len(self.processing_df.columns) # fixed + modified columns (col 0 fixed, 1..n-1 modified)
        self.table.setRowCount(num_rows)
        self.table.setColumnCount(num_cols)
        headers = ['Fixed']
        for col_index in range(1, num_cols):
            headers.append(self.get_element_name(col_index))
        self.table.setHorizontalHeaderLabels(headers)
        for i in range(num_rows):
            fixed_item = QTableWidgetItem(str(self.fixed_column[i]))
            fixed_item.setFlags(Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
            self.table.setItem(i, 0, fixed_item)
            for col_index in range(1, num_cols):
                mod_val = self.processed_columns.get(col_index, [None] * num_rows)[i]
                mod_item = QTableWidgetItem(str(mod_val) if mod_val is not None else "")
                mod_item.setFlags(Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                self.table.setItem(i, col_index, mod_item)
    def fill_empty_cells(self):
        if self.all_processed_mode:
            return # Fill not available in all mode
        min_val = self.min_spin.value()
        max_val = self.max_spin.value()
        offset = self.offset_spin.value()
        ratio = self.ratio_spin.value()
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
        self.status_bar.showMessage("Filled empty cells")
    def global_check_duplicates(self):
        col_index = self.column_combo.currentData()
        if col_index is None:
            return
        if not self.all_processed_mode:
            self.load_column(col_index)
        self.check_duplicates(col_index)
    def check_duplicates(self, col_index):
        table_col = 2 if not self.all_processed_mode else col_index
        selected_items = self.table.selectedItems()
        if not selected_items:
            self.status_bar.showMessage("No rows selected for duplicates")
            return
       
        selected_rows = set(item.row() for item in selected_items if item.column() == table_col and item.row() != self.crm_reference_row)
       
        originals = self.processing_df.iloc[:, col_index]
        values = [originals[row] for row in selected_rows
                  if originals[row] is not None and isinstance(originals[row], (int, float))]
       
        if not values:
            self.status_bar.showMessage("No valid values in selected rows")
            return
       
        mean_val = sum(values) / len(values)
        dup_range = self.global_dup_range_spin.value()
        light_yellow = QColor(255, 255, 150)
        light_red = QColor(255, 180, 180)
        for row in selected_rows:
            self.table.item(row, table_col).setBackground(QBrush(light_yellow))
            self.table.item(row, 0).setBackground(QBrush(light_yellow))
        for row in selected_rows:
            val = originals[row]
            if val is not None and isinstance(val, (int, float)) and abs(val - mean_val) > mean_val * dup_range:
                self.table.item(row, table_col).setBackground(QBrush(light_red))
                self.table.item(row, 0).setBackground(QBrush(light_red))
        self.status_bar.showMessage("Checked duplicates")
    def global_fix_duplicates(self):
        col_index = self.column_combo.currentData()
        if col_index is None:
            return
        if not self.all_processed_mode:
            self.load_column(col_index)
        self.fix_duplicates(col_index)
    def fix_duplicates(self, col_index):
        table_col = 2 if not self.all_processed_mode else col_index
        selected_items = self.table.selectedItems()
        if not selected_items:
            self.status_bar.showMessage("No rows selected for fixing duplicates")
            return
       
        selected_rows = set(item.row() for item in selected_items if item.column() == table_col and item.row() != self.crm_reference_row)
        originals = self.processing_df.iloc[:, col_index]
        values = [originals[row] for row in selected_rows
                  if originals[row] is not None and isinstance(originals[row], (int, float))]
        if not values:
            return
        mean_val = sum(values) / len(values)
       
        min_val = self.min_spin.value()
        max_val = self.max_spin.value()
       
        light_red = QColor(255, 180, 180)
        light_green = QColor(180, 255, 180)
        for row in selected_rows:
            mod_item = self.table.item(row, table_col)
            orig_val = originals[row]
            if mod_item and mod_item.background().color() == light_red:
                rand_factor = random.uniform(min_val, max_val)
                new_val = round(mean_val * rand_factor, 2)
                self.processed_columns[col_index][row] = new_val
                mod_item.setText(str(new_val))
                self.table.item(row, 0).setBackground(QBrush(light_red))
                mod_item.setBackground(QBrush(light_green))
            else:
                mod_val = self.processed_columns[col_index][row]
                if mod_val is None and orig_val is not None:
                    self.processed_columns[col_index][row] = orig_val
                    mod_item.setText(str(orig_val))
        self.status_bar.showMessage("Fixed duplicates")
    def remove_crm_reference_row(self):
        if self.crm_reference_row is not None:
            self.table.removeRow(self.crm_reference_row)
            if self.crm_reference_row <= self.crm_row:
                self.crm_row -= 1
            self.crm_reference_row = None
            self.global_clear_crm_button.setEnabled(False)
    def global_compare_with_crm(self):
        col_index = self.column_combo.currentData()
        if col_index is None:
            return
        if not self.all_processed_mode:
            self.load_column(col_index)
        self.compare_with_crm(col_index)
        self.global_clear_crm_button.setEnabled(True)
    def compare_with_crm(self, col_index):
        table_col = 2 if not self.all_processed_mode else col_index
        selected_items = self.table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Please select at least one row.")
            return
       
        selected_rows = set(item.row() for item in selected_items if item.column() == table_col and item.row() != self.crm_reference_row)
        if len(selected_rows) != 1:
            QMessageBox.warning(self, "Error", "Please select exactly one row for CRM comparison.")
            return
       
        self.crm_row = next(iter(selected_rows))
        originals = self.processing_df.iloc[:, col_index]
        crm_original = originals[self.crm_row]
        if crm_original is None or not isinstance(crm_original, (int, float)):
            QMessageBox.warning(self, "Error", "Selected CRM row has no valid Original value.")
            return
        # نام عنصر از سطر 1
        element_name = self.get_element_name(col_index)
        if element_name not in self.crm_903:
            QMessageBox.warning(self, "Error", f"CRM 903 value not available for element: {element_name}")
            return
       
        crm_903_val = self.crm_903[element_name]
        crm_range = self.global_crm_range_spin.value()
        light_green = QColor(180, 255, 180)
        light_red = QColor(255, 180, 180)
        self.remove_crm_reference_row()
        insert_row = self.crm_row + 1
        self.table.insertRow(insert_row)
        self.crm_reference_row = insert_row
        # For all mode, only set in fixed and the specific column
        fixed_item = QTableWidgetItem("CRM 903")
        fixed_item.setBackground(QBrush(QColor(200, 200, 255)))
        fixed_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
        self.table.setItem(insert_row, 0, fixed_item)
        if not self.all_processed_mode:
            orig_item = QTableWidgetItem(str(crm_903_val))
            orig_item.setBackground(QBrush(QColor(200, 200, 255)))
            orig_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self.table.setItem(insert_row, 1, orig_item)
            mod_item = QTableWidgetItem("")
            mod_item.setBackground(QBrush(QColor(200, 200, 255)))
            mod_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self.table.setItem(insert_row, 2, mod_item)
        else:
            # In all mode, set CRM value in the mod column
            crm_item = QTableWidgetItem(str(crm_903_val))
            crm_item.setBackground(QBrush(QColor(200, 200, 255)))
            crm_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            self.table.setItem(insert_row, col_index, crm_item)
            # Fill other columns with empty items to maintain structure
            for c in range(1, self.table.columnCount()):
                if c != col_index:
                    empty_item = QTableWidgetItem("")
                    empty_item.setBackground(QBrush(QColor(200, 200, 255)))
                    empty_item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                    self.table.setItem(insert_row, c, empty_item)
        mod_item = self.table.item(self.crm_row, table_col)
        if mod_item:
            try:
                mod_val = float(mod_item.text())
                if abs(mod_val - crm_903_val) <= crm_903_val * crm_range:
                    color = light_green
                else:
                    color = light_red
                mod_item.setBackground(QBrush(color))
                self.table.item(self.crm_row, 0).setBackground(QBrush(color))
            except ValueError:
                pass
        self.status_bar.showMessage("Compared with CRM 903")
    def global_fix_crm_differences(self):
        col_index = self.column_combo.currentData()
        if col_index is None:
            return
        if not self.all_processed_mode:
            self.load_column(col_index)
        self.fix_crm_differences(col_index)
    def fix_crm_differences(self, col_index):
        if self.crm_row is None or self.crm_reference_row is None:
            QMessageBox.warning(self, "Error", "No CRM row selected. Use 'Compare with CRM 903' first.")
            return
       
        table_col = 2 if not self.all_processed_mode else col_index
        element_name = self.get_element_name(col_index)
        if element_name not in self.crm_903:
            return
       
        crm_903_val = self.crm_903[element_name]
        crm_range = self.global_crm_range_spin.value()
        min_factor = 1.0 - crm_range
        max_factor = 1.0 + crm_range
        light_green = QColor(180, 255, 180)
        rand_factor = random.uniform(min_factor, max_factor)
        new_val = round(crm_903_val * rand_factor, 6)
        self.processed_columns[col_index][self.crm_row] = new_val
        mod_item = self.table.item(self.crm_row, table_col)
        if mod_item is None:
            mod_item = QTableWidgetItem()
            self.table.setItem(self.crm_row, table_col, mod_item)
        mod_item.setText(str(new_val))
        mod_item.setBackground(QBrush(light_green))
        self.table.item(self.crm_row, 0).setBackground(QBrush(light_green))
        self.remove_crm_reference_row()
        self.status_bar.showMessage("Fixed CRM differences")
    def global_clear_crm_row(self):
        col_index = self.column_combo.currentData()
        if col_index is None:
            return
        if not self.all_processed_mode:
            self.load_column(col_index)
        self.clear_crm_row(col_index)
        self.global_clear_crm_button.setEnabled(False)
    def clear_crm_row(self, col_index):
        table_col = 2 if not self.all_processed_mode else col_index
        self.remove_crm_reference_row()
        if self.crm_row is not None:
            self.table.item(self.crm_row, table_col).setBackground(QBrush(QColor("white")))
            self.table.item(self.crm_row, 0).setBackground(QBrush(QColor("white")))
            if not self.all_processed_mode:
                self.table.item(self.crm_row, 1).setBackground(QBrush(QColor("white")))
            self.crm_row = None
        self.status_bar.showMessage("Cleared CRM row")
    def apply_limits(self):
        limit_row = self.reserved_rows[4]
        for col_index in range(1, len(self.processing_df.columns)):
            limit_val = limit_row[col_index] if not pd.isna(limit_row[col_index]) else None
           
            if limit_val is None or not isinstance(limit_val, (int, float)):
                continue
            mods = self.processed_columns[col_index]
            table_col = 2 if not self.all_processed_mode else col_index
            for i in range(len(mods)):
                mod_val = mods[i]
                if mod_val is not None and isinstance(mod_val, (int, float)):
                    if mod_val < limit_val:
                        new_val = f"<{limit_val}"
                        mods[i] = new_val
                        if self.all_processed_mode or self.current_column_index == col_index:
                            item = self.table.item(i, table_col)
                            if item:
                                item.setText(new_val)
        self.status_bar.showMessage("Applied limits to all columns")
    def finalize_data(self):
        if not self.all_processed_mode:
            self.save_current_modified()
            num_columns = len(self.processing_df.columns) - 1
            if len(self.processed_columns) != num_columns:
                QMessageBox.warning(self, "Error", "Process all columns first.")
                return
       
        for col_index, col_data in self.processed_columns.items():
            self.processing_df.iloc[:, col_index] = col_data
       
        full_df = pd.DataFrame(columns=self.df.columns, index=range(len(self.df)))
        full_df.iloc[1] = self.header_row # بازگرداندن هدر اصلی (سطر 1)
        full_df.iloc[2] = self.reserved_rows[2]
        full_df.iloc[3] = self.reserved_rows[3]
        full_df.iloc[4] = self.reserved_rows[4]
        full_df.iloc[5] = self.reserved_rows[5]
        full_df.iloc[6:] = self.processing_df.values
       
        save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "Excel (*.xlsx)")
        if save_path:
            full_df.to_excel(save_path, index=False, header=False)
            QMessageBox.information(self, "Saved", "File saved successfully.")
            self.status_bar.showMessage(f"Saved file: {save_path}")
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Arial", 10))
    window = DataProcessor()
    window.show()
    sys.exit(app.exec())