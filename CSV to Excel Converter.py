import sys
import pandas as pd
import os
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QLineEdit, QLabel, QFileDialog, QMessageBox

class CSVtoExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('CSV to Excel Converter')
        self.setGeometry(100, 100, 400, 200)
        
        # Layout
        layout = QVBoxLayout()
        
        # CSV File Path Entry
        self.csv_file_path = QLineEdit(self)
        self.csv_file_path.setReadOnly(True)
        self.csv_file_path.setPlaceholderText("Click to select CSV file...")
        self.csv_file_path.mousePressEvent = self.load_csv
        layout.addWidget(self.csv_file_path)
        
        # Directory Path Entry
        self.folder_path = QLineEdit(self)
        self.folder_path.setReadOnly(True)
        self.folder_path.setPlaceholderText("Click to select folder to save Excel file...")
        self.folder_path.mousePressEvent = self.select_folder
        layout.addWidget(self.folder_path)
        
        # File Name Entry
        fileNameLayout = QVBoxLayout()
        label = QLabel("Output File Name:")
        fileNameLayout.addWidget(label)
        self.file_name = QLineEdit(self)
        fileNameLayout.addWidget(self.file_name)
        layout.addLayout(fileNameLayout)
        
        # Convert Button
        self.convert_button = QPushButton('Convert to Excel', self)
        self.convert_button.clicked.connect(self.convert)
        layout.addWidget(self.convert_button)
        
        self.setLayout(layout)
        self.show()

    def load_csv(self, event):
        file_path, _ = QFileDialog.getOpenFileName(self, "Open CSV File", "", "CSV Files (*.csv)")
        if file_path:
            self.csv_file_path.setText(file_path)

    def select_folder(self, event):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.folder_path.setText(folder_path)

    def convert(self):
        csv_path = self.csv_file_path.text()
        folder_path = self.folder_path.text()
        file_name = self.file_name.text()
        
        if not csv_path or not folder_path or not file_name:
            QMessageBox.warning(self, "Warning", "Please select a CSV file, specify a folder, and enter a file name.")
            return
        
        excel_path = os.path.join(folder_path, file_name + '.xlsx')
        try:
            df = pd.read_csv(csv_path, encoding='utf-8')
            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            QMessageBox.information(self, "Success", "The file has been converted and saved successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = CSVtoExcelApp()
    sys.exit(app.exec_())