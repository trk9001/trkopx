"""The new fillsh.py (in development).

This script is intended to eventually replace fillsh.py, since that one hasn't
been updated in ages. Trying to reduce the number of *hacks* present in that
script. Please install `openpyxl` via pip before trying to use this script.
"""

import os.path
import re

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font


class FillSheet:
    
    DEFAULT_FILE = 'Descr.xlsx'

    def __init__(self, file=DEFAULT_FILE, rows=None):
        """Set up all the required variables.

        Args:
            file -- Name of the Excel file
            rows -- Range of rows to work on

        Raises:
            FileNotFoundError
            TypeError
            ValueError
        """

        self.file = None
        self.rows = None
        self.seed = None

        # Validate and set the file name
        if '.xlsx' not in file:
            raise TypeError('INVALID FILE: '
                            'Must be a valid Excel file ending in .xlsx')
        elif not os.path.exists(file):
            raise FileNotFoundError('FILE DOES NOT EXIST')
        else:
            self.file = file

        # A struct for the `self.rows` data
        class Rows:
            def __init__(self):
                self.start = None
                self.end = None

        # Validate and set the range of rows
        if rows is None:
            self.rows = Rows()
            self.rows.start = 2
            self.rows.end = self.get_number_of_rows(self.file)
        elif isinstance(rows, str) and re.match(r'^\d*:\d*$', rows):
            self.rows = Rows()
            row_start, row_end = tuple(rows.split(':'))
            self.rows.start = 2 if row_start == '' else int(row_start)
            self.rows.end = 0 if row_end == '' else int(row_end)
            max_rows = self.get_number_of_rows(self.file)
            if self.rows.end == 0 or self.rows.end > max_rows:
                self.rows.end = max_rows
            if self.rows.start > self.rows.end:
                raise ValueError('INVALID VALUE FOR ROWS: '
                                 'START must be less than END')
        else:
            raise TypeError('INVALID TYPE FOR ROWS: '
                            'Must be empty or of the form START:END')

        # Set the seed column index
        self.seed = self.get_manufacturer_column_index(self.file)
        if self.seed is None:
            raise ValueError('MANUFACTURER COLUMN\'S INDEX NOT FOUND')

    @staticmethod
    def get_number_of_rows(file):
        """Return the number of rows in a worksheet."""
        wb = load_workbook(file)
        return len(list(wb.active.rows))

    @staticmethod
    def get_manufacturer_column_index(file):
        """Return the index of the Manufacturer column in a worksheet."""
        wb = load_workbook(file)
        ws = wb.active

        # The Manufacturer column is usually among the first ones,
        # barring the very first column.
        for col in range(2, 10):
            if ws.cell(row=1, column=col).value == 'MANUFACTURER':
                return col

    def half_fill(self):
        """Generate boilerplate text in the second description column."""

        # Set column numbers sequentially from the seed value, for columns:
        # manufacturer, product, colour, fulldescr1 and fulldescr2.
        mnf, pdt, clr, dc1, dc2 = (self.seed + i for i in range(5))

        wb = load_workbook(self.file)
        ws = wb.active

        # In every iteration of the loop, the next value of the 'product'
        # column is checked against its current value, so set it to None
        # initially.
        pdt_val = None

        i = self.rows.start - 1

        while i < self.rows.end:
            i += 1

            # If there is any text in the second description column, move
            # it to the first description column.
            if ws.cell(i, dc2).value:
                tmp = ws.cell(i, dc1).value
                if tmp is None: tmp = ''
                ws.cell(i, dc1).value = tmp + '; ' + ws.cell(i, dc2).value
                ws.cell(i, dc2).value = None

            # Skip duplicates of the last product
            if ws.cell(i, pdt).value == pdt_val:
                continue

            mnf_val, pdt_val, clr_val = (ws.cell(i, col).value
                                         for col in [mnf, pdt, clr])
            descr = 'The {} from {} comes in {} colour, featuring'.format(
                    pdt_val, mnf_val, clr_val
            )

            self.format_cell(ws.cell(i, dc2))
            ws.cell(i, dc2).value = descr

        wb.save(self.file)

    def full_fill(self):
        """Generate full descriptions in the first column."""
        pass

    @staticmethod
    def format_cell(cell):
        """Formats a cell from a worksheet opened using openpyxl."""
        cell.alignment = Alignment(horizontal='left', wrap_text=True)
        cell.font = Font(name='Calibri', size=8)
