from pandas import DataFrame

from config import SORT_BY_COLUMN, PRIMARY_COLUMN, SECONDARY_COLUMN
from lib.filelogic import FileLogic
from lib.warehouse_conversion_excel import WarehouseConversionExcel


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
        df = self.to_data_frame()
        for _, serial_row in df.loc[df['Serial'].notnull()].iterrows():
            # Search for conversion
            conversion_row = self.warehouse_conversion.find_conversion(self.sheet_name,
                                                                       serial_row['Warehouse'],
                                                                       serial_row['Sub Inventory'])
            # Search for serial in opposite DF and conversion warehouse
            opposite_sheet = PRIMARY_COLUMN if self.sheet_name == SECONDARY_COLUMN else SECONDARY_COLUMN
            conversion_warehouse = conversion_row[opposite_sheet + " Warehouse"]
            conversion_sub_inventory = conversion_row[opposite_sheet + " Sub Inventory"]
            opposite_df = compare_df.loc[
                (compare_df['Warehouse'] == conversion_warehouse) &
                (compare_df['Sub Inventory'] == conversion_sub_inventory) &
                (compare_df['Serial'] == serial_row['Serial'])
                ]
            if opposite_df.empty:
                # Check if exists in another warehouse
                opposite_df = compare_df.loc[
                    (compare_df['Serial'] == serial_row['Serial'])
                ]
                if opposite_df.empty:
                    self.add_invalid(serial_row.values.tolist(),
                                     f"Missing serial {serial_row['Serial']} in {opposite_sheet} "
                                     f"worksheet warehouse: '{conversion_warehouse} "
                                     f"{conversion_sub_inventory}'")
                else:
                    opposite_df = opposite_df.iloc[0]
                    self.add_invalid(serial_row.values.tolist(),
                                     f"Mismatch serial {serial_row['Serial']} in {opposite_sheet} "
                                     f"worksheet expected to be in '{conversion_warehouse} {conversion_sub_inventory}'"
                                     f" but found in warehouse: "
                                     f"'{opposite_df['Warehouse']} {opposite_df['Sub Inventory']}'")

    def check_non_serial_items(self, compare_df: DataFrame):
        compare_df = compare_df.astype(str)
        compare_df[["Quantity"]] = compare_df[["Quantity"]].astype(int)
        df = self.to_data_frame()
        query = df.query('Serial.isnull()', engine='python').groupby(['Warehouse', 'Sub Inventory', 'Part Number'])[
            'Quantity'].sum()
        for quantity_based_row in query.items():
            # Search for conversion
            conversion_row = self.warehouse_conversion.find_conversion(self.sheet_name,
                                                                       quantity_based_row[0][0],
                                                                       quantity_based_row[0][1])
            # Search for serial in opposite DF and conversion warehouse
            opposite_sheet = PRIMARY_COLUMN if self.sheet_name == SECONDARY_COLUMN else SECONDARY_COLUMN
            conversion_warehouse = conversion_row[opposite_sheet + " Warehouse"]
            conversion_sub_inventory = conversion_row[opposite_sheet + " Sub Inventory"]
            # compare_df.columns = [column.replace(" ", "_") for column in compare_df.columns]
            opposite_df = compare_df.query(
                f'`Part Number` == "{str(quantity_based_row[0][2])}" and '
                f'`Warehouse` == "{conversion_warehouse}" and '
                f'`Sub Inventory` == "{str(conversion_sub_inventory)}"').groupby(
                ['Warehouse', 'Sub Inventory', 'Part Number'])[
                'Quantity'].sum().reset_index()

            if opposite_df.empty:
                self.add_invalid([quantity_based_row[0][2], None, quantity_based_row[1], quantity_based_row[0][0],
                                  quantity_based_row[0][1]],
                                 f"Missing item {quantity_based_row[0][2]} in {opposite_sheet} "
                                 f"worksheet warehouse: '{conversion_warehouse} "
                                 f"{conversion_sub_inventory}'")
            else:
                # Item exists but quantity may not match
                if opposite_df['Quantity'][0] != quantity_based_row[1]:
                    self.add_invalid([quantity_based_row[0][2], None, quantity_based_row[1], quantity_based_row[0][0],
                                      quantity_based_row[0][1]],
                                     f"Quantity mismatch in item {quantity_based_row[0][2]} in {opposite_sheet} "
                                     f"worksheet warehouse: '{conversion_warehouse} "
                                     f"{conversion_sub_inventory}' expected {quantity_based_row[1]} actual "
                                     f"{opposite_df['Quantity'][0]}")
