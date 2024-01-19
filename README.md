
# Excel File Updater

### Description
This script updates an Excel file using the `openpyxl` library. It fills empty cells in columns G and H with zeros for rows 1 to 2340 in a specified worksheet.

### Requirements
- Python
- openpyxl library

### Installation
Install the required `openpyxl` library using pip:
```bash
pip install openpyxl
```

### Usage
1. Set the `file_path` variable to your Excel file's path.
2. Set the `sheet_name` variable to the name of the sheet you want to update.
3. Run the script to update the file:
```bash
python your_script_name.py
```

### Functionality
- Loads an Excel workbook from a given path.
- Checks for the existence of the specified sheet.
- Fills empty cells in columns G and H with 0 for rows 1 to 2340.
- Saves the updated workbook.

### Example
```python
file_path = "path/to/your/excelfile.xlsx"
sheet_name = "Sheet1"
update_excel_file(file_path, sheet_name)
```

*Note*: Replace `path/to/your/excelfile.xlsx` and `Sheet1` with your actual file path and sheet name.
