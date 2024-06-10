import os
import shutil
import sys
import itertools
import re
from PyQt5.QtWidgets import (QApplication, QWidget, QLabel, QLineEdit, QPushButton, 
                             QMessageBox, QTextEdit, QFileDialog, QCheckBox, QComboBox, QListWidget,
                             QVBoxLayout, QHBoxLayout)
from PyQt5.QtGui import QFont
from pandas import read_csv

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('CSV File Management For Gladiator')
        self.resize(1200, 800)
        self.setStyleSheet('background-color: #e1d0ba;')

        # Source Root Folder
        source_label = QLabel('Source Root Folder:', self)
        source_label.setGeometry(10, 10, 2000, 30)
        source_label.setFont(QFont('Helvetica Bold', 12))
        self.source_var = QLineEdit(self)
        self.source_var.setGeometry(180, 10, 700, 30)
        source_button = QPushButton('Browse', self)
        source_button.setGeometry(890, 10, 100, 30)
        source_button.clicked.connect(self.browseSource)

        # Destination Folder
        destination_label = QLabel('Destination Folder:', self)
        destination_label.setGeometry(10, 50, 150, 30)
        destination_label.setFont(QFont('Helvetica Bold', 12))
        self.destination_var = QLineEdit(self)
        self.destination_var.setGeometry(180, 50, 700, 30)
        destination_button = QPushButton('Browse', self)
        destination_button.setGeometry(890, 50, 100, 30)
        destination_button.clicked.connect(self.browseDestination)

        # Source Directory
        source_directory_label = QLabel('Source Directory:', self)
        source_directory_label.setGeometry(10, 90, 150, 30)
        source_directory_label.setFont(QFont('Helvetica Bold', 12))
        self.source_directory_var = QLineEdit(self)
        self.source_directory_var.setGeometry(180, 90, 700, 30)
        source_directory_button = QPushButton('Browse', self)
        source_directory_button.setGeometry(890, 90, 100, 30)
        source_directory_button.clicked.connect(self.browseSourceDirectory)

        # File Type Selection
        self.file_type_csv = QCheckBox('CSV', self)
        self.file_type_csv.setGeometry(890, 150, 60, 30)
        self.file_type_csv.setFont(QFont('Helvetica Bold', 12))
        self.file_type_csv.setChecked(True)  # Set CSV as default

        self.file_type_excel = QCheckBox('Excel', self)
        self.file_type_excel.setGeometry(890, 120, 80, 30)
        self.file_type_excel.setFont(QFont('Helvetica Bold', 12))

        self.file_type_csv.stateChanged.connect(self.handleFileTypeChange)
        self.file_type_excel.stateChanged.connect(self.handleFileTypeChange)

        # Number of Files
        num_files_label = QLabel('Number of Files:', self)
        num_files_label.setGeometry(650, 180, 150, 30)
        num_files_label.setFont(QFont('Helvetica Bold', 12))
        self.num_files_var = QLineEdit(self)
        self.num_files_var.setGeometry(770, 180, 100, 30)

        # Add Mapping Button
        add_mapping_button = QPushButton('Mapping', self)
        add_mapping_button.setGeometry(890, 180, 100, 30)
        add_mapping_button.clicked.connect(self.addMappingFields)

        # Folder Names
        folder_names_label = QLabel('Folder Names:', self)
        folder_names_label.setGeometry(10, 130, 150, 30)
        folder_names_label.setFont(QFont('Helvetica Bold', 12))
        self.folder_names_var = QLineEdit(self)
        self.folder_names_var.setGeometry(180, 130, 700, 30)

        # Number of Files per Folder
        num_files_per_folder_label = QLabel('Number of Files per Folder:', self)
        num_files_per_folder_label.setGeometry(10, 180, 250, 30)
        num_files_per_folder_label.setFont(QFont('Helvetica Bold', 12))
        self.num_files_per_folder_var = QLineEdit(self)
        self.num_files_per_folder_var.setGeometry(220, 180, 100, 30)

        # Number of Repeats
        num_repeats_label = QLabel('Number of Repeats:', self)
        num_repeats_label.setGeometry(370, 180, 150, 30)
        num_repeats_label.setFont(QFont('Helvetica Bold', 12))
        self.num_repeats_var = QLineEdit(self)
        self.num_repeats_var.setGeometry(520, 180, 100, 30)

        # Log
        log_label = QLabel('Log:', self)
        log_label.setGeometry(540, 340, 250, 30)
        log_label.setFont(QFont('Helvetica Bold', 12))
        self.log_text = QTextEdit(self)
        self.log_text.setGeometry(375, 390, 450, 320)

        # Original File Names
        original_names_label = QLabel('Original File Names:', self)
        original_names_label.setGeometry(10, 340, 250, 30)
        original_names_label.setFont(QFont('Helvetica Bold', 12))
        self.original_names_text = QTextEdit(self)
        self.original_names_text.setGeometry(10, 390, 350, 320)

        # Desired New Names
        new_names_label = QLabel('Desired New Names:', self)
        new_names_label.setGeometry(980, 340, 250, 30)
        new_names_label.setFont(QFont('Helvetica Bold', 12))
        self.new_names_text = QTextEdit(self)
        self.new_names_text.setGeometry(840, 390, 350, 320)

        # Clean Missing Values
       

        self.column_selection_label = QLabel('Columns to Clean:', self)
        self.column_selection_label.setGeometry(10, 240, 200, 30)
        self.column_selection_label.setFont(QFont('Helvetica Bold', 12))
        self.column_selection_var = QComboBox(self)
        self.column_selection_var.setGeometry(150, 240, 300, 30)
        self.column_selection_var.currentIndexChanged.connect(self.addSelectedColumn)

        self.selected_columns_label = QLabel('', self)
        self.selected_columns_label.setGeometry(320, 280, 200, 30)
        self.selected_columns_label.setFont(QFont('Helvetica Bold', 12))
        self.selected_columns_list = QListWidget(self)
        self.selected_columns_list.setGeometry(450, 240, 300, 60)

        self.remove_column_button = QPushButton('Remove', self)
        self.remove_column_button.setGeometry(760, 240, 80, 30)
        self.remove_column_button.clicked.connect(self.removeSelectedColumn)

        # Desired Value Selection
        desired_value_label = QLabel('Desired Value:', self)
        desired_value_label.setGeometry(10, 280, 150, 30)
        desired_value_label.setFont(QFont('Helvetica Bold', 12))
        self.desired_value_var = QLineEdit(self)
        self.desired_value_var.setGeometry(150, 280, 200, 30)

        # Buttons
        copy_button = QPushButton('Copy CSV', self)
        copy_button.setGeometry(10, 730, 100, 30)
        copy_button.clicked.connect(self.copyAndRenameCSVFiles)

        create_button = QPushButton('Create Folders', self)
        create_button.setGeometry(540, 730, 150, 30)
        create_button.clicked.connect(self.createAndRenameFolders)

        rename_button = QPushButton('Rename CSV', self)
        rename_button.setGeometry(1040, 730, 100, 30)
        rename_button.clicked.connect(self.renameCSVFilesInFolders)

        # Populate columns in combobox
        columns = ['Longitude', 'Latitude', 'RSSI (0)', 'R0 RSRP (0)', 'R0 RS CINR (0)', 'Cell ID (0)']
        self.column_selection_var.addItems(columns)

    def browseSource(self):
        folder_path = QFileDialog.getExistingDirectory(self, 'Select Source Root Folder')
        if folder_path:
            self.source_var.setText(folder_path)

    def browseDestination(self):
        folder_path = QFileDialog.getExistingDirectory(self, 'Select Destination Folder')
        if folder_path:
            self.destination_var.setText(folder_path)

    def browseSourceDirectory(self):
        folder_path = QFileDialog.getExistingDirectory(self, 'Select Source Directory')
        if folder_path:
            self.source_directory_var.setText(folder_path)

    def handleFileTypeChange(self):
        if self.sender() == self.file_type_csv:
            self.file_type_excel.setChecked(False)
        elif self.sender() == self.file_type_excel:
            self.file_type_csv.setChecked(False)

    def addMappingFields(self):
        try:
            num_files = int(self.num_files_var.text())
            self.original_names_text.clear()
            self.new_names_text.clear()
            for i in range(1, num_files + 1):
                original_name = f"Original File {i}\n"
                new_name = f"New File {i}\n"
                self.original_names_text.insertPlainText(original_name)
                self.new_names_text.insertPlainText(new_name)
        except ValueError:
            self.log_text.append("Please enter a valid number of files.")

    def copyAndRenameCSVFiles(self):
        source_root_folder = self.source_var.text()
        destination_folder = self.destination_var.text()

        if not os.path.exists(source_root_folder):
            self.log_text.append(f"Source root folder '{source_root_folder}' does not exist.")
            return

        if not os.path.exists(destination_folder):
            self.log_text.append(f"Destination folder '{destination_folder}' does not exist.")
            return

        file_extension = '.csv' if self.file_type_csv.isChecked() else '.xlsx'

        try:
            columns_to_clean = self.getColumnsToClean()
            desired_value = self.desired_value_var.text()

            for folder_name in os.listdir(source_root_folder):
                folder_path = os.path.join(source_root_folder, folder_name)
                if os.path.isdir(folder_path):
                    csv_files = [f for f in os.listdir(folder_path) if f.endswith('.csv')]
                    if len(csv_files) == 1:
                        csv_file_path = os.path.join(folder_path, csv_files[0])
                        new_file_name = f"{folder_name}_data{file_extension}"
                        destination_file_path = os.path.join(destination_folder, new_file_name)
                        
                        df = read_csv(csv_file_path)
                        for column in columns_to_clean:
                            if desired_value == "missing":
                                df = df.dropna(subset=[column])
                            else:
                                df[column] = df[column].replace(desired_value, float('nan'))
                        df.to_csv(destination_file_path, index=False)
                        
                        self.log_text.append(f"File '{csv_file_path}' copied and cleaned values for selected columns, then saved as '{new_file_name}'.")
        except Exception as e:
            self.log_text.append(f"An error occurred while copying and cleaning CSV files: {e}")

    def createAndRenameFolders(self):
        source_dir = self.source_directory_var.text()

        if not os.path.exists(source_dir):
            self.log_text.append(f"Source directory '{source_dir}' does not exist.")
            return

        try:
            csv_files = [file for file in os.listdir(source_dir) if file.endswith(".csv")]
            csv_files.sort(key=self.extract_number)

            total_files = len(csv_files)
            folder_names = self.folder_names_var.text().split(",")
            num_folders = len(folder_names)
            folder_count = 0
            file_index = 0

            while file_index < total_files:
                folder_name = folder_names[folder_count].strip()
                folder_path = os.path.join(source_dir, folder_name)
                if not os.path.exists(folder_path):
                    os.mkdir(folder_path)

                for _ in range(int(self.num_files_per_folder_var.text())):
                    if file_index < total_files:
                        old_file_path = os.path.join(source_dir, csv_files[file_index])
                        new_file_path = os.path.join(folder_path, csv_files[file_index])
                        os.rename(old_file_path, new_file_path)
                        file_index += 1

                self.log_text.append(f"Moved {int(self.num_files_per_folder_var.text())} CSV files to folder '{folder_name}'")
                folder_count = (folder_count + 1) % num_folders

            self.log_text.append("CSV files moved to folders successfully!")

        except Exception as e:
            self.log_text.append(f"An error occurred: {e}")

    def extract_number(self, filename):
        match = re.search(r'\d+', filename)
        return int(match.group()) if match else 0

    def renameCSVFilesInFolders(self):
        source_directory = self.source_directory_var.text()

        if not os.path.exists(source_directory):
            self.log_text.append(f"Source directory '{source_directory}' does not exist.")
            return

        file_extension = '.csv' if self.file_type_csv.isChecked() else '.xlsx'
        desired_names = self.new_names_text.toPlainText().splitlines()

        for folder_name in self.folder_names_var.text().split(","):
            folder_name = folder_name.strip()
            source_folder = os.path.join(source_directory, folder_name)

            if not os.path.exists(source_folder):
                self.log_text.append(f"Folder '{source_folder}' does not exist.")
                continue

            desired_names_cycle = itertools.cycle(desired_names)

            for file_name in os.listdir(source_folder):
                if file_name.endswith('.csv'):
                    desired_name = next(desired_names_cycle)
                    new_file_name = f"{folder_name}_{desired_name}{file_extension}"  # Prepend the folder name to the desired name
                    old_file_path = os.path.join(source_folder, file_name)
                    new_file_path = os.path.join(source_folder, new_file_name)
                    try:
                        os.rename(old_file_path, new_file_path)
                        self.log_text.append(f"File '{file_name}' renamed to '{new_file_name}'.")
                    except Exception as e:
                        self.log_text.append(f"Failed to rename '{file_name}' to '{new_file_name}': {e}")

    def addSelectedColumn(self):
        column = self.column_selection_var.currentText()
        self.selected_columns_list.addItem(column)

    def removeSelectedColumn(self):
        selected_items = self.selected_columns_list.selectedItems()
        for item in selected_items:
            self.selected_columns_list.takeItem(self.selected_columns_list.row(item))

    def getColumnsToClean(self):
        columns_to_clean = []
        for index in range(self.selected_columns_list.count()):
            columns_to_clean.append(self.selected_columns_list.item(index).text())
        return columns_to_clean

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
