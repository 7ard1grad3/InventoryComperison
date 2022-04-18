from typing import List

import pandas as pd
from rich.console import Console
from rich.table import Table


class FileLogic:
    def __init__(self, file: str, sheet_name: str = None):
        self.sheet_name = sheet_name
        self.file = file
        self.valid = True
        self.errors = []
        self.fields = []
        self.valid_rows = []
        self.invalid_rows = []

    def mark_as_invalid(self):
        self.valid = False

    def add_error(self, message: str, message_type: str):
        message_type = message_type.upper()
        self.errors.append({
            "message": message,
            "type": message_type,
        })
        if message_type == 'ERROR':
            self.mark_as_invalid()

    def add_invalid(self, row: List, issue_description: str):
        row.append(issue_description)
        self.invalid_rows.append(row)

    def is_valid(self, print_table=True):
        if print_table:
            self.show_errors()
        return self.valid

    def show_errors(self):
        if len(self.errors) > 0:
            table = Table(show_header=True, header_style="bold magenta")
            table.add_column("Message", style="dim", width=60)
            table.add_column("Type", style="red")
            for error in self.errors:
                table.add_row(error['message'], error['type'])
            Console().print(table)

    def validate_worksheet(self):
        try:
            df = pd.read_excel(self.file, sheet_name=self.sheet_name)
        except FileNotFoundError as e:
            Console().log(f"file {e.filename} not found ", style="red on white")
            return False
        except ValueError:
            self.add_error(f"Worksheet named '{self.sheet_name}' not found", 'error')
            return False
        for required_field in self.fields:
            if required_field not in df.columns:
                self.add_error(f"Missing field {required_field}", 'error')
        return pd.DataFrame(df, columns=self.fields)

    def to_data_frame(self):
        return pd.DataFrame(self.valid_rows, columns=self.fields)
