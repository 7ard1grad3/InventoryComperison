import pandas as pd
from pandas import DataFrame
from config import SORT_BY_COLUMN, PRIMARY_COLUMN, SECONDARY_COLUMN
from lib.filelogic import FileLogic
from lib.warehouse_conversion_excel import WarehouseConversionExcel
from rich.console import Console
import re

class InventoryExcel(FileLogic):
    def __init__(self, file: str, sheet_name: str, warehouse_conversion: WarehouseConversionExcel):
        super().__init__(file, sheet_name)
        self.fields = ['Part Number', 'Serial', 'Quantity', 'Warehouse', 'Sub Inventory']
        self.warehouse_conversion = warehouse_conversion
        self.validate_and_clean_excel()

    def validate_and_clean_excel(self):
        data = self.validate_worksheet()
        if data is not False:
            for id_n, row in enumerate(data.sort_values(by=SORT_BY_COLUMN).values):
                valid_row = []
                id_n += 1
                # month must be 1-12
                valid_row.append(row[0])
                valid_row.append(row[1])

                if type(row[2]) != int and row[2] > 0:
                    self.add_error(
                        f"'Quantity' must be number that above 0. failed at row {id_n}",
                        'error')
                else:
                    valid_row.append(row[2])

                # Check that warehouse listed
                conversion_row = self.warehouse_conversion.find_conversion(self.sheet_name, row[3], row[4])
                if conversion_row is None or conversion_row.empty:
                    self.add_error(
                        f"Missing conversion for warehouse {row[3]} {row[4]}. failed at row {id_n}",
                        'error')
                else:
                    valid_row.append(row[3])
                    valid_row.append(row[4])
                self.valid_rows.append(valid_row)
        else:
            return False

    def check_serial_items(self, compare_df: DataFrame):
        console = Console()

        # ✅ Ensure we use the updated DataFrame (self.df), NOT reload it from the file
        df = self.df if hasattr(self, "df") else self.to_data_frame()

        # Replace NaN with empty string, then convert to uppercase for case-insensitive comparison
        df['Serial'] = df['Serial'].fillna("").astype(str).str.upper()
        compare_df['Serial'] = compare_df['Serial'].fillna("").astype(str).str.upper()


        for _, serial_row in df.loc[df['Serial'].notnull()].iterrows():
            serial_value = serial_row['Serial']

            # ✅ Skip empty serials (in case clearing failed)
            if serial_value.strip() == "":
                continue  # Ignore blank serials

            # Check warehouse conversion
            conversion_row = self.warehouse_conversion.find_conversion(
                self.sheet_name, serial_row['Warehouse'], serial_row['Sub Inventory']
            )

            opposite_sheet = PRIMARY_COLUMN if self.sheet_name == SECONDARY_COLUMN else SECONDARY_COLUMN
            conversion_warehouse = conversion_row[opposite_sheet + " Warehouse"]
            conversion_sub_inventory = conversion_row[opposite_sheet + " Sub Inventory"]

            opposite_df = compare_df.loc[
                (compare_df['Warehouse'] == conversion_warehouse) &
                (compare_df['Sub Inventory'] == conversion_sub_inventory) &
                (compare_df['Serial'] == serial_value)
                ]

            if opposite_df.empty:
                opposite_df = compare_df.loc[compare_df['Serial'] == serial_value]
                if opposite_df.empty:

                    self.add_invalid(serial_row.values.tolist(),
                                     f"Missing serial {serial_value} in {opposite_sheet} "
                                     f"worksheet warehouse: '{conversion_warehouse} "
                                     f"{conversion_sub_inventory}'")
                else:
                    opposite_df = opposite_df.iloc[0]

                    self.add_invalid(serial_row.values.tolist(),
                                     f"Mismatch serial {serial_value} in {opposite_sheet} "
                                     f"worksheet expected to be in '{conversion_warehouse} {conversion_sub_inventory}'"
                                     f" but found in warehouse: "
                                     f"'{opposite_df['Warehouse']} {opposite_df['Sub Inventory']}'")

    def check_non_serial_items(self, compare_df: DataFrame):
        console = Console()

        # ✅ Ensure we use the updated DataFrame (self.df), NOT reload it from the file
        df = self.df if hasattr(self, "df") else self.to_data_frame()

        # Convert all to string to avoid type mismatches
        compare_df = compare_df.astype(str)
        df = df.astype(str)

        # Preserve original formatting but create a case-insensitive matching column
        compare_df['Part Number (Lower)'] = compare_df['Part Number'].str.strip().str.lower()
        df['Part Number (Lower)'] = df['Part Number'].str.strip().str.lower()

        # Convert quantity column to integer (handling possible NaN or non-numeric values)
        compare_df["Quantity"] = pd.to_numeric(compare_df["Quantity"], errors='coerce').fillna(0).astype(int)
        df["Quantity"] = pd.to_numeric(df["Quantity"], errors='coerce').fillna(0).astype(int)

        # Group non-serial items by warehouse, sub-inventory, and case-insensitive part number
        query = df.query('Serial.isnull() or Serial == ""', engine='python') \
            .groupby(['Warehouse', 'Sub Inventory', 'Part Number (Lower)'])['Quantity'].sum()

        for quantity_based_row in query.items():
            # Search for warehouse conversion
            conversion_row = self.warehouse_conversion.find_conversion(
                self.sheet_name, quantity_based_row[0][0], quantity_based_row[0][1])

            # If no conversion found, skip
            if conversion_row is None:
                continue

            # Identify opposite warehouse details
            opposite_sheet = PRIMARY_COLUMN if self.sheet_name == SECONDARY_COLUMN else SECONDARY_COLUMN
            conversion_warehouse = conversion_row[opposite_sheet + " Warehouse"]
            conversion_sub_inventory = conversion_row[opposite_sheet + " Sub Inventory"]

            # Search for the part number in the opposite sheet (case-insensitive)
            opposite_df = compare_df.query(
                f'`Part Number (Lower)` == "{quantity_based_row[0][2]}" and '
                f'`Warehouse` == "{conversion_warehouse}" and '
                f'`Sub Inventory` == "{conversion_sub_inventory}"'
            ).groupby(['Warehouse', 'Sub Inventory', 'Part Number (Lower)'])['Quantity'].sum().reset_index()

            if opposite_df.empty:
                # ✅ Preserve original part number in the results, even though comparison was case-insensitive
                original_part_number = \
                df.loc[df['Part Number (Lower)'] == quantity_based_row[0][2], 'Part Number'].values[0]

                self.add_invalid([
                    original_part_number, None, quantity_based_row[1],
                    quantity_based_row[0][0], quantity_based_row[0][1]
                ], f"Missing item {original_part_number} in {opposite_sheet} "
                   f"worksheet warehouse: '{conversion_warehouse} {conversion_sub_inventory}'")
            else:
                # Preserve original part number format in reporting
                original_part_number = \
                df.loc[df['Part Number (Lower)'] == quantity_based_row[0][2], 'Part Number'].values[0]

                if opposite_df['Quantity'][0] != quantity_based_row[1]:
                    self.add_invalid([
                        original_part_number, None, quantity_based_row[1],
                        quantity_based_row[0][0], quantity_based_row[0][1]
                    ], f"Quantity mismatch in item {original_part_number} in {opposite_sheet} "
                       f"worksheet warehouse: '{conversion_warehouse} {conversion_sub_inventory}' "
                       f"expected {quantity_based_row[1]} actual {opposite_df['Quantity'][0]}")

    def ignore_serials_for_non_serialized_items(self, non_serialized_items):
        """
        Temporarily removes serials for items that are marked as non-serialized.
        This is done in-memory without modifying the actual Excel file.
        """
        console = Console()  # Initialize Console for printing

        # Load the DataFrame
        df = self.to_data_frame()

        # Ensure "Part Number" is formatted consistently
        df["Part Number"] = df["Part Number"].astype(str).str.strip().str.upper()
        non_serialized_items = {str(item).strip().upper() for item in non_serialized_items}

        # 🔍 Debugging: Print cleaned part numbers


        # 🔍 Check if special characters are causing mismatches
        df["Part Number"] = df["Part Number"].apply(lambda x: x.strip().upper())  # Preserve dashes
        non_serialized_items = {item.strip().upper() for item in non_serialized_items}

        # Identify unmatched part numbers
        unmatched_items = set(df["Part Number"]) - non_serialized_items


        # Create mask for filtering non-serialized items
        mask = df["Part Number"].isin(non_serialized_items)


        # Clear serials
        df.loc[mask, "Serial"] = ""


        # Update self.df to ensure the modified DataFrame is used in future processing
        self.df = df.copy()

        console.print(f"✅ Serials ignored for non-serialized items in {self.sheet_name} worksheet.")
