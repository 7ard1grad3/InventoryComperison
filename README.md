# Inventory comparison

Use it to compare gaps between 2 inventory tables

## Instructions
* Place xlsx file (*copy **template.xlsx** as a template*) to the **Excel** folder
* Fill all the needed worksheets with your data
* Run **compare.exe**
* Results file will be stored in the root directory with the name **results.xlsx**

### via python
Install all the required packages from the requirements.txt

```bash
pip install -r requirements.txt
python3 main.py
# compile to exe
python -m nuitka main.py
```

## Configurations

Can be changed in the config.py file

```python
# Folder for Excel files to process
EXCEL_FOLDER = 'Excel'
# Name of the result file
RESULTS_FILE = 'results.xlsx'
# Column to sort for the primary and secondary tables
SORT_BY_COLUMN = 'Warehouse'
# Name of the primary worksheet
PRIMARY_COLUMN = 'Primary'
# Name of the secondary worksheet
SECONDARY_COLUMN = 'Secondary'
# Name of the conversion worksheet
CONVERSION_WORKSHEET = 'Warehouse Conversion'

```