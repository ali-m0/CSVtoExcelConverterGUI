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

```bash```
pip install pandas openpyxl PyQt5

## Running the Application
To run the application, execute the following command in your terminal or command prompt:

```bash```
python csv_to_excel_converter.py

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

```bash```
pip install pyinstaller
2. ***Navigate to the directory containing your Python script (csv_to_excel_converter.py) and run the following command:***
```bash```
- *pyinstaller --onefile --noconsole csv_to_excel_converter.py*
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