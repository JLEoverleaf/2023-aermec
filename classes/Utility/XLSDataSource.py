# XLSDataSource.py
# Abstraction of an excel sheet
import xlrd


class XLSDataSource:
    """
    Grabs an excel sheet and handles movement through that sheet
    """

    def __init__(
        self,
        sheet: xlrd.sheet,
        min_row: int = 0,
        max_row: int = 500,
        min_column: int = 0,
        max_column: int = 200,
    ):
        self.sheet: xlrd.sheet = sheet
        self.row_num = sheet.nrows
        self.column_num = sheet.ncols
        self._row = 0
        self._column = 0
        self.min_row = min_row
        self.max_row = max_row
        self.min_column = min_column
        self.max_column = max_column

    @classmethod
    def from_blob(cls, blob: bytes, index: int = 0, err_msg: str = None, file_type="xls", **kwargs):
        try:
            book: xlrd.book = xlrd.open_workbook(file_contents=blob)
            if book.nsheets < index + 1:
                raise Exception(f"Worksheet with index {index} is not available")
            return cls(book.sheet_by_index(0), **kwargs)
        except Exception as err:
            if err_msg is None:
                raise err
            else:
                raise Exception(err_msg)

    def cell_stripper(self, row, column):
        val = self.sheet.cell_value(rowx=row, colx=column)
        return val.strip() if isinstance(val, str) else val

    def get_abs_cell_value(self, row, column):
        return self.cell_stripper(row, column)

    def get_rel_cell_value(self, delta_row, delta_column):
        return self.cell_stripper(self.row + delta_row, self.column + delta_column)

    @property
    def cell(self) -> str:
        return self.cell_stripper(self.row, self.column)

    @property
    def cell_right(self) -> str:
        return self.get_rel_cell_value(0, 1)

    @property
    def cell_right_2(self) -> str:
        return self.get_rel_cell_value(0, 2)

    @property
    def cell_right_3(self) -> str:
        return self.get_rel_cell_value(0, 3)

    @property
    def cell_right_4(self) -> str:
        return self.get_rel_cell_value(0, 4)

    @property
    def row(self) -> int:
        return self._row

    @row.setter
    def row(self, proposed_row) -> None:
        if proposed_row > self.max_row or proposed_row > self.row_num - 1:
            raise Exception("Moving down too far in the xls sheet.")
        elif proposed_row < self.min_row:
            raise Exception("Moving up too far in the xls sheet.")
        else:
            self._row = proposed_row

    @property
    def column(self) -> int:
        return self._column

    @column.setter
    def column(self, proposed_column) -> None:
        if proposed_column > self.max_column or proposed_column > self.column_num - 1:
            raise Exception("Moving too far to the right in the xls sheet.")
        elif proposed_column < self.min_column:
            raise Exception("Moving too far to the left in the xls sheet.")
        else:
            self._column = proposed_column

    def move_down_one(self) -> None:
        self.row = self._row + 1

    def move_right_one(self) -> None:
        self.column = self._column + 1

    def move_down(self, delta_row: int) -> None:
        self.row = self._row + abs(delta_row)

    def move_right(self, delta_column: int) -> None:
        self.column = self._column + abs(delta_column)

    def move_abs(self, row: int, column: int) -> None:
        self.row = row
        self.column = column

    def move_relative(self, delta_row: int, delta_column: int) -> None:
        self.row = self._row + delta_row
        self.column = self._column + delta_column
