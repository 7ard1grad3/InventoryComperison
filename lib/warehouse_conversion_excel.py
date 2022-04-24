from lib.filelogic import FileLogic


class WarehouseConversionExcel(FileLogic):
    def __init__(self, file: str, sheet_name: str):
        super().__init__(file, sheet_name)
        self.fields = ['Primary Warehouse', 'Primary Sub Inventory', 'Secondary Warehouse', 'Secondary Sub Inventory']
        self.validate_and_clean_excel()

    def validate_and_clean_excel(self):
        # check for required structure of the Excel file
        data = self.validate_worksheet()
        if data is not False:
            for id_n, row in enumerate(data.sort_values(by=self.fields[0]).values):
                valid_row = []
                for id_x, required_field in enumerate(self.fields):
                    if not row[id_x]:
                        self.add_error(
                            f"'{required_field}' is empty. failed at row {id_n}",
                            'error')
                    else:
                        valid_row.append(row[id_x])
                self.valid_rows.append(valid_row)

        else:
            return False

    def find_conversion(self, sheet_name: str, warehouse: str, sub_warehouse: str):
        df = self.to_data_frame()
        try:
            location = df.loc[
                (df[sheet_name + ' Warehouse'] == warehouse) &
                (df[sheet_name + ' Sub Inventory'] == sub_warehouse)].iloc[0]
        except:
            return None
        return location
