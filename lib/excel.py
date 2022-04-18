import glob
import os

import pandas as pd
from rich.console import Console

from config import EXCEL_FOLDER, PRIMARY_COLUMN, SECONDARY_COLUMN, CONVERSION_WORKSHEET
from lib.filelogic import FileLogic
from lib.inventory_excel import InventoryExcel
from lib.warehouse_conversion_excel import WarehouseConversionExcel


class Excel(FileLogic):
    def __init__(self):
        super().__init__(Excel.get_excel_file())

    @staticmethod
    def get_excel_file():
        try:
            os.chdir(f"./{EXCEL_FOLDER}")
            for file in glob.glob("*.xlsx"):
                return f"../{EXCEL_FOLDER}/{file}"
        except FileNotFoundError:
            return None
        return None

    def start_comparison(self):
        if self.file is not None:
            Console().print(f"✅ Found excel file in [bold magenta]{self.file}[/bold magenta]")
        else:
            self.add_error(
                f"❌ Missing file. [bold magenta]Make sure to place file in {EXCEL_FOLDER} folder "
                f"with .xlsx format[/bold magenta]", "error")
        if self.is_valid():
            # Get the list of conversions
            warehouse_conversion = WarehouseConversionExcel(self.file, CONVERSION_WORKSHEET)
            if warehouse_conversion.is_valid():
                # Validate primary table
                primary_inventory = InventoryExcel(self.file, PRIMARY_COLUMN, warehouse_conversion)
                if primary_inventory.is_valid():
                    # Validate secondary table
                    secondary_inventory = InventoryExcel(self.file, SECONDARY_COLUMN, warehouse_conversion)
                    if secondary_inventory.is_valid():
                        Console().print(f"✅ Validation complete 😎 - Starting comparison "
                                        f"\n🔃 Checking {PRIMARY_COLUMN} worksheet by serial...")
                        primary_inventory.check_serial_items(secondary_inventory.to_data_frame())
                        Console().print(f"🔃 Checking {SECONDARY_COLUMN} worksheet by serial...")
                        secondary_inventory.check_serial_items(primary_inventory.to_data_frame())
                        Console().print(f"🔃 Checking {PRIMARY_COLUMN} worksheet by quantity...")
                        primary_inventory.check_non_serial_items(secondary_inventory.to_data_frame())
                        Console().print(f"🔃 Checking {SECONDARY_COLUMN} worksheet by quantity...")
                        secondary_inventory.check_non_serial_items(primary_inventory.to_data_frame())

                        df = pd.DataFrame([primary_inventory.fields + ["Issue"]] + primary_inventory.invalid_rows +
                                          secondary_inventory.invalid_rows)
                        writer = pd.ExcelWriter('../results.xlsx',
                                                engine='xlsxwriter')
                        df.to_excel(writer, sheet_name='Results', index=False, header=False)
                        writer.save()
                        Console().print(f"✅ Results saved to [bold magenta]results.xlsx[/bold magenta] in the root "
                                        f"folder 😃")
