# Inventory Comparison Tool

## Overview
This tool allows you to compare two inventory tables (Primary and Secondary) and identify gaps or discrepancies between them. It's particularly useful for warehouse inventory reconciliation, stock audits, and inventory management processes where you need to compare data from different sources.

The tool processes Excel files with specific worksheets, performs the comparison based on configurable parameters, and generates a comprehensive results file highlighting the differences.

## Instructions

### Quick Start
1. Copy the **template.xlsx** file as a starting point
2. Place your Excel file in the **Excel** folder
3. Fill in all required worksheets with your inventory data (see Data Format section below)
4. Run **compare.exe**
5. Results will be generated in the root directory as **results.xlsx**

### Data Format
The template Excel file should contain the following worksheets:

* **Primary**: Your main inventory data
  * Must include a 'Warehouse' column (or the column specified in config.py)
  * All other columns will be used in the comparison

* **Secondary**: Your secondary inventory data for comparison
  * Must have the same column structure as the Primary worksheet

* **Warehouse Conversion**: (Optional) Used to map warehouse names/codes between systems
  * Useful when the same warehouse has different identifiers in different systems

### Example Data Structure

**Primary Worksheet:**
```
Warehouse | Item    | Quantity
----------|---------|----------
WH001     | Item-A  | 100
WH002     | Item-B  | 50
```

**Secondary Worksheet:**
```
Warehouse | Item    | Quantity
----------|---------|----------
WH001     | Item-A  | 95
WH002     | Item-C  | 75
```

**Results will show:**
- Item-A has a quantity difference of 5
- Item-B is missing from Secondary
- Item-C is in Secondary but not in Primary

## Running via Python

### Requirements
- Python 3.8 or higher
- Required packages listed in requirements.txt

### Installation
```bash
pip install -r requirements.txt
```

### Execution
```bash
python main.py
```

### Building Executable
```bash
# Using Nuitka (recommended)
python -m nuitka main.py

# Alternative: Using PyInstaller
pip install pyinstaller
pyinstaller --onefile main.py
```

## Configuration Options

All configuration options can be modified in the `config.py` file:

```python
# Folder for Excel files to process
EXCEL_FOLDER = 'Excel'

# Name of the result file
RESULTS_FILE = 'results.xlsx'

# Column to sort for the primary and secondary tables
# This column must exist in both Primary and Secondary worksheets
SORT_BY_COLUMN = 'Warehouse'

# Name of the primary worksheet in your Excel file
PRIMARY_COLUMN = 'Primary'

# Name of the secondary worksheet in your Excel file
SECONDARY_COLUMN = 'Secondary'

# Name of the conversion worksheet (for mapping warehouse names)
CONVERSION_WORKSHEET = 'Warehouse Conversion'
```

## Troubleshooting

### Common Issues

1. **Missing Columns**: Ensure both Primary and Secondary worksheets have identical column headers

2. **File Access Error**: Close any open Excel files before running the comparison

3. **Empty Results**: Check that your data follows the expected format and that the worksheet names match the configuration

4. **Excel Format Issues**: Save your Excel file in .xlsx format (not .xls)

5. **Warehouse Mapping**: If using different warehouse codes between systems, ensure the Warehouse Conversion sheet is properly configured

### Getting Help
If you encounter issues not covered in this documentation, please check the source code in the `lib` directory for more detailed implementation details.