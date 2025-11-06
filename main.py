import sys
import pandas as pd
import random
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QLineEdit, QLabel, QCheckBox, QMessageBox, QSlider
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
        self.layout = QVBoxLayout(self.central_widget)

        # Buttons
        self.load_button = QPushButton("Load CSV/Excel File")
        self.load_button.clicked.connect(self.load_file)
        self.layout.addWidget(self.load_button)

        self.next_column_button = QPushButton("Next Column")
        self.next_column_button.clicked.connect(self.next_column)
        self.layout.addWidget(self.next_column_button)
        self.next_column_button.setEnabled(False)

        # Control panel for filling
        control_layout = QHBoxLayout()
        self.min_edit = QLineEdit("0.9")
        self.max_edit = QLineEdit("1.1")
        self.offset_edit = QLineEdit("0")
        self.ratio_edit = QLineEdit("1.0")
        self.apply_filled_checkbox = QCheckBox("Apply to filled cells")
        self.apply_ratio_checkbox = QCheckBox("Apply ratio to filled cells")

        control_layout.addWidget(QLabel("Min:"))
        control_layout.addWidget(self.min_edit)
        control_layout.addWidget(QLabel("Max:"))
        control_layout.addWidget(self.max_edit)
        control_layout.addWidget(QLabel("Offset:"))
        control_layout.addWidget(self.offset_edit)
        control_layout.addWidget(QLabel("Ratio:"))
        control_layout.addWidget(self.ratio_edit)
        control_layout.addWidget(self.apply_filled_checkbox)
        control_layout.addWidget(self.apply_ratio_checkbox)

        self.fill_button = QPushButton("Fill Empty Cells")
        self.fill_button.clicked.connect(self.fill_empty_cells)
        control_layout.addWidget(self.fill_button)
        self.layout.addLayout(control_layout)

        # Duplicate handling
        dup_layout = QHBoxLayout()
        self.dup_range_edit = QLineEdit("0.05")  # Range for duplicates
        dup_layout.addWidget(QLabel("Duplicate Range:"))
        dup_layout.addWidget(self.dup_range_edit)

        self.check_dup_button = QPushButton("Check Duplicates")
        self.check_dup_button.clicked.connect(self.check_duplicates)
        dup_layout.addWidget(self.check_dup_button)

        self.fix_dup_button = QPushButton("Fix Duplicates")
        self.fix_dup_button.clicked.connect(self.fix_duplicates)
        dup_layout.addWidget(self.fix_dup_button)

        self.layout.addLayout(dup_layout)

        # CRM handling
        crm_layout = QHBoxLayout()
        self.crm_range_edit = QLineEdit("0.1")  # Range for CRM comparison
        crm_layout.addWidget(QLabel("CRM Range:"))
        crm_layout.addWidget(self.crm_range_edit)

        self.select_crm_button = QPushButton("Select CRM Row")
        self.select_crm_button.clicked.connect(self.select_crm_row)
        crm_layout.addWidget(self.select_crm_button)

        self.compare_crm_button = QPushButton("Compare with CRM 901")
        self.compare_crm_button.clicked.connect(self.compare_with_crm)
        crm_layout.addWidget(self.compare_crm_button)

        self.fix_crm_button = QPushButton("Fix CRM Differences")
        self.fix_crm_button.clicked.connect(self.fix_crm_differences)
        crm_layout.addWidget(self.fix_crm_button)

        self.layout.addLayout(crm_layout)

        # Apply limits from row 4
        self.apply_limits_button = QPushButton("Apply Limits < >")
        self.apply_limits_button.clicked.connect(self.apply_limits)
        self.layout.addWidget(self.apply_limits_button)

        # Finalize
        self.finalize_button = QPushButton("Finalize and Save")
        self.finalize_button.clicked.connect(self.finalize_data)
        self.layout.addWidget(self.finalize_button)

        # Table for current column
        self.table = QTableWidget()
        self.layout.addWidget(self.table)

        # Data storage
        self.df = None
        self.reserved_rows = {}
        self.processed_columns = []
        self.current_column_index = 0
        self.current_column_data = None  # For the two-column table
        self.crm_row = None
        self.crm_901 = {}  # Hardcoded CRM 901 data, assume dict with column names as keys
        # Example: self.crm_901 = {'col1': 100, 'col2': 200, ...}  # Replace with actual data

    def load_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open File", "", "CSV/Excel (*.csv *.xlsx)")
        if file_path:
            if file_path.endswith('.csv'):
                self.df = pd.read_csv(file_path, header=None)
            else:
                self.df = pd.read_excel(file_path, header=None)
            
            # Reserve rows 1,3,4,5,6 (0-indexed: 0,2,3,4,5)
            self.reserved_rows = {
                0: self.df.iloc[0].copy(),
                2: self.df.iloc[2].copy(),
                3: self.df.iloc[3].copy(),
                4: self.df.iloc[4].copy(),
                5: self.df.iloc[5].copy()
            }
            
            # Processing data starts from row 6 (index 6) onwards, drop reserved rows
            self.processing_df = self.df.iloc[6:].reset_index(drop=True)
            
            # Clean cells like <2 or >2 to 2
            for col in self.processing_df.columns:
                self.processing_df[col] = self.processing_df[col].apply(self.clean_cell)
            
            self.next_column_button.setEnabled(True)
            self.load_column(0)

    def clean_cell(self, cell):
        if isinstance(cell, str):
            if cell.startswith('<') or cell.startswith('>'):
                return float(cell[1:])
            elif cell.endswith('<') or cell.endswith('>'):
                return float(cell[:-1])
        return cell

    def load_column(self, col_index):
        if col_index >= len(self.processing_df.columns):
            QMessageBox.information(self, "Done", "All columns processed.")
            return
        
        col_data = self.processing_df.iloc[:, col_index].dropna().values  # Assume no NaN for simplicity
        self.current_column_data = pd.DataFrame({
            'Original': col_data,
            'Modified': [None] * len(col_data)
        })
        
        self.table.setRowCount(len(col_data))
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(['Original', 'Modified'])
        
        for i, val in enumerate(col_data):
            self.table.setItem(i, 0, QTableWidgetItem(str(val)))
        
        self.current_column_index = col_index

    def next_column(self):
        self.processed_columns.append(self.current_column_data['Modified'].values)
        self.load_column(self.current_column_index + 1)

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
                continue  # Skip invalid
            
            modified = self.current_column_data.at[i, 'Modified']
            if modified is None or apply_filled:
                rand_factor = random.uniform(min_val, max_val)
                new_val = (original * rand_factor) + offset
                if apply_ratio_filled or modified is None:
                    new_val *= ratio
                self.current_column_data.at[i, 'Modified'] = new_val
                self.table.setItem(i, 1, QTableWidgetItem(str(new_val)))

    def check_duplicates(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        selected_rows = set(item.row() for item in selected_items if item.column() == 1)
        values = [self.current_column_data.at[row, 'Modified'] for row in selected_rows if self.current_column_data.at[row, 'Modified'] is not None]
        
        if not values:
            return
        
        mean_val = sum(values) / len(values)
        dup_range = float(self.dup_range_edit.text())
        
        for row in selected_rows:
            val = self.current_column_data.at[row, 'Modified']
            if abs(val - mean_val) > mean_val * dup_range:
                item = self.table.item(row, 1)
                item.setBackground(QBrush(QColor("red")))

    def fix_duplicates(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        
        selected_rows = set(item.row() for item in selected_items if item.column() == 1)
        values = [self.current_column_data.at[row, 'Modified'] for row in selected_rows if self.current_column_data.at[row, 'Modified'] is not None]
        mean_val = sum(values) / len(values)
        
        min_val = float(self.min_edit.text())  # Reuse min/max for random
        max_val = float(self.max_edit.text())
        
        for row in selected_rows:
            rand_factor = random.uniform(min_val, max_val)
            new_val = mean_val * rand_factor
            self.current_column_data.at[row, 'Modified'] = new_val
            self.table.setItem(row, 1, QTableWidgetItem(str(new_val)))
            item = self.table.item(row, 1)
            item.setBackground(QBrush(QColor("white")))

    def select_crm_row(self):
        selected_items = self.table.selectedItems()
        if len(selected_items) != 1 or selected_items[0].column() != 1:
            QMessageBox.warning(self, "Error", "Select one modified cell as CRM.")
            return
        row = selected_items[0].row()
        self.crm_row = row
        QMessageBox.information(self, "Selected", f"CRM row selected: {row}")

    def compare_with_crm(self):
        if self.crm_row is None:
            QMessageBox.warning(self, "Error", "Select CRM row first.")
            return
        
        col_name = self.processing_df.columns[self.current_column_index]  # Assume columns have names if needed
        crm_901_val = self.crm_901.get(col_name, 0)  # Get from hardcoded
        
        crm_val = self.current_column_data.at[self.crm_row, 'Modified']
        crm_range = float(self.crm_range_edit.text())
        
        for i in range(len(self.current_column_data)):
            val = self.current_column_data.at[i, 'Modified']
            if val is not None and abs(val - crm_901_val) > crm_901_val * crm_range:
                item = self.table.item(i, 1)
                item.setBackground(QBrush(QColor("red")))

    def fix_crm_differences(self):
        if self.crm_row is None:
            return
        
        col_name = self.processing_df.columns[self.current_column_index]
        crm_901_val = self.crm_901.get(col_name, 0)
        
        min_val = float(self.min_edit.text())
        max_val = float(self.max_edit.text())
        
        for i in range(len(self.current_column_data)):
            item = self.table.item(i, 1)
            if item.background().color() == QColor("red"):
                rand_factor = random.uniform(min_val, max_val)
                new_val = crm_901_val * rand_factor
                self.current_column_data.at[i, 'Modified'] = new_val
                self.table.setItem(i, 1, QTableWidgetItem(str(new_val)))
                item.setBackground(QBrush(QColor("white")))

    def apply_limits(self):
        limit_row = self.reserved_rows[3]  # Row 4 (0-index 3) for limits per column
        col_index = self.current_column_index
        limit = limit_row[col_index] if not pd.isna(limit_row[col_index]) else 0
        
        for i in range(len(self.current_column_data)):
            val = self.current_column_data.at[i, 'Modified']
            if val is not None and val < limit:
                new_val = f"<{limit}"
            elif val > limit * 10:  # Arbitrary large, adjust as needed
                new_val = f">{limit}"
            else:
                new_val = val
            self.current_column_data.at[i, 'Modified'] = new_val
            self.table.setItem(i, 1, QTableWidgetItem(str(new_val)))

    def finalize_data(self):
        if len(self.processed_columns) != len(self.processing_df.columns):
            QMessageBox.warning(self, "Error", "Process all columns first.")
            return
        
        # Reassemble processing_df with modified columns
        for i, col_data in enumerate(self.processed_columns):
            self.processing_df.iloc[:, i] = col_data
        
        # Reinsert reserved rows
        full_df = pd.DataFrame(columns=self.df.columns)
        full_df.loc[0] = self.reserved_rows[0]
        full_df.loc[2] = self.reserved_rows[2]
        full_df.loc[3] = self.reserved_rows[3]
        full_df.loc[4] = self.reserved_rows[4]
        full_df.loc[5] = self.reserved_rows[5]
        full_df = pd.concat([full_df, self.processing_df], ignore_index=True)
        
        # Save to file
        save_path, _ = QFileDialog.getSaveFileName(self, "Save File", "", "CSV (*.csv)")
        if save_path:
            full_df.to_csv(save_path, index=False, header=False)
            QMessageBox.information(self, "Saved", "File saved successfully.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DataProcessor()
    window.show()
    sys.exit(app.exec())