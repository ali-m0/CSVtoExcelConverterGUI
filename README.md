# CSV to Excel Converter GUI

A simple and user-friendly desktop application built with Python and PyQt5 that allows users to convert CSV files to Excel (XLSX) format. With an intuitive graphical user interface (GUI), users can easily select a CSV file, choose an output folder, specify the output file name, and perform the conversion with just a few clicks.

## Features

- **Easy-to-Use GUI**: The application is built with PyQt5, providing a simple and intuitive interface for users to interact with.
- **CSV to Excel Conversion**: Convert CSV files to Excel format (.xlsx) seamlessly.
- **Custom Output Location**: Choose the folder where the converted Excel file will be saved.
- **Custom File Naming**: Specify the name of the output Excel file.
- **Error Handling**: User-friendly error messages are displayed for missing input, file issues, or any other conversion errors.
- **Standalone Executable**: Convert the Python script to a standalone executable (.exe) using PyInstaller.

## Requirements

To run this project, you'll need the following:

- **Python 3.x**
- **Libraries**:
  - `pandas`
  - `openpyxl`
  - `PyQt5`

### Installation

First, install the required libraries using `pip`:

```
pip install pandas openpyxl PyQt5
```
## Python Code:

```
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
```
## Running the Application
To run the application, execute the following command in your terminal or command prompt:

```
python csv_to_excel_converter.py
```

## Usage
1. **Select a CSV File**: Click on the text box labeled "Click to select CSV file..." to open a file dialog and select the CSV file you want to convert.
2. **Choose an Output Folder**: Click on the text box labeled "Click to select folder to save Excel file..." to open a folder dialog and select the directory where you want the converted Excel file to be saved.
3. **Specify Output File Name**: Enter the desired name for the output Excel file in the "Output File Name" text box.
4. **Convert**: Click the "Convert to Excel" button to convert the selected CSV file to an Excel file. If the conversion is successful, a success message will be displayed.

## Example
1. **Launch the application.**
2. **Select example.csv from your computer.**
3. **Choose an output folder such as "C:/Documents/ConvertedFiles."**
4. **Enter output_file as the file name.**
5. **Click "Convert to Excel," and the file output_file.xlsx will be saved in the selected folder.**

## Converting to an Executable (.exe)
You can convert the Python script into a standalone executable (.exe) file using PyInstaller. This allows users to run the application without needing to have Python installed.

### Steps to Convert to .exe

1. ***Install PyInstaller:***

```
pip install pyinstaller
```
2. ***Navigate to the directory containing your Python script (csv_to_excel_converter.py) and run the following command:***
```
- pyinstaller --onefile --noconsole csv_to_excel_converter.py
```
- *--onefile: Packages everything into a single executable file.*
- *--noconsole: Hides the console window when running the executable.*

3. ***After the process completes, you will find the generated .exe file in the dist folder. You can distribute this file to users, allowing them to run the application without needing to install Python or any dependencies.***

## About PyQt5

PyQt5 is a set of Python bindings for the Qt application framework and is used to build cross-platform applications with a robust graphical user interface. PyQt5 provides modules for developing interactive and feature-rich applications.

## Known Issues

**Performance**: Converting very large CSV files may take some time depending on system performance.
**Excel Compatibility**: The resulting Excel file is in .xlsx format and should be compatible with most modern spreadsheet applications.

## Contribution

If you have suggestions or want to contribute to this project, please contact me directly.

### Contact Information:

- **Author**: Ali Mansouri
- **LinkedIn**: [Ali Mansouri](https://www.linkedin.com/in/ali-mansouri-a7984215b/)
- **Email**: [ali.mansouri1998@gmail.com](mailto:ali.mansouri1998@gmail.com)

## License

For more information and understand more about this process or write some product like this, please contact me directly with email or LinkedIn that provided above.