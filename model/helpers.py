from curses.ascii import isalpha
from collections.abc import Callable

from rules import *
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter


class NotFoundError(Exception):
    pass


class NoValueError(Exception):
    pass


class ValueColumn:
    def __init__(
        self,
        sheet: Worksheet,
        label: str,
        column: str,
        start_cell_row: int,
        end_cell_row: int,
    ):
        self.sheet = sheet
        self.column = column
        self.label = label
        self.start_cell_row = start_cell_row
        self.end_cell_row = end_cell_row

    @staticmethod
    def is_exact_match(value: str, cell_content: str) -> bool:
        return value == cell_content

    def get_row_of(
        self, value: str, matcher: Callable[[str, str], bool] = is_exact_match
    ) -> bool:
        for row in range(self.start_cell_row, self.end_cell_row + 1):
            if matcher(value, self.get_value_at(row)):
                return row
        raise NotFoundError(
            f"couldn't find match of {value}"
            + f"having searched in value column of {self.label} "
            + f"from row number {self.start_cell_row} "
            + f"to row number {self.end_cell_row}"
        )

    def get_value_at(self, row: int) -> str:
        return self.sheet[self.column + row]

    def set_value_at(self, row: int, new_value: str) -> None:
        self.sheet[self.column + row].value = new_value


class CellIdManipulator:
    def __init__(self, cell_id: str):
        self.cell_id = cell_id

    def _get_row_id_start_position(self) -> int:
        position: 0

        while isalpha(self.cell_id[position]):
            position += 1

        return position

    def get_column(self) -> str:
        column_id = ""

        for character in self.cell_id:
            if isalpha(character):
                column_id += character

        return column_id

    def get_row(self) -> int:
        return int(self.cell_id[self._get_row_id_start_position() :])

    def get_below(self, down_steps: int) -> str:
        return str(self.get_column() + (self.get_row() + down_steps))


class Helper:
    def __init__(self, sheet: Worksheet):
        self.sheet = sheet

    def _find(self, cell_content: str) -> str:
        upper_column_range = SPOT_SEARCH_COLUMNS + 1
        upper_row_range = SPOT_SEARCH_ROWS + 1

        for column in range(1, upper_column_range):
            for row in range(1, upper_row_range):

                current_cell_id = get_column_letter(column) + str(row)
                current_cell = self.sheet(current_cell_id)

                if current_cell.value == cell_content:
                    return current_cell_id

        raise NotFoundError(
            f"couldn't find cell with the cell content of {cell_content} "
            + f"after searching {SPOT_SEARCH_ROWS} rows "
            + f"and {SPOT_SEARCH_COLUMNS} columns"
        )

    def get_column(self, label: str) -> str:
        cell_id = self._find(label)

        return CellIdManipulator(cell_id).get_column()

    def create_value_column(self, label: str) -> ValueColumn:
        cell_id = self._find(label)

        last_non_empty_cell = None
        empty_cells_streak = 0

        def update_states(base_cell: str) -> None:
            below_last_visited_cell = CellIdManipulator(base_cell).get_below(1)

            if not self.sheet[below_last_visited_cell].value:
                empty_cells_streak += 1
            else:
                last_non_empty_cell = below_last_visited_cell
                empty_cells_streak = 0

        while empty_cells_streak < MAX_EMPTY_ROWS_FOR_VALUE_COLUMN:
            if not last_non_empty_cell:
                update_states(cell_id)
            else:
                update_states(
                    CellIdManipulator(last_non_empty_cell).get_below(
                        empty_cells_streak + 1
                    )
                )

        if not last_non_empty_cell:
            raise NoValueError(f"the rows below the label {label} are empty")

        column = CellIdManipulator(cell_id).get_column()
        start_row_id = CellIdManipulator(cell_id).get_below(1)
        start_row = CellIdManipulator(start_row_id).get_row()
        end_row = CellIdManipulator(last_non_empty_cell).get_row()

        return ValueColumn(self.sheet, label, column, start_row, end_row)
