from curses.ascii import isalpha
from collections.abc import Callable

import models.rules
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter


class NotFoundError(Exception):
    pass


class NoValueError(Exception):
    pass


class UnacceptableMergedCellError(Exception):
    pass


# REQUIRES: each of the value cells must not be merged
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

        for row in range(self.start_cell_row, self.end_cell_row + 1):
            current_cell_id = column + str(row)
            if type(sheet[current_cell_id]).__name__ == "MergedCell":
                raise UnacceptableMergedCellError(
                    f"cell {current_cell_id} of sheet {sheet}, "
                    + f"but value cells in a value column ({label}) cannot be merged"
                )

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

    def get_start_cell_row(self) -> int:
        return self.start_cell_row

    def get_end_cell_row(self) -> int:
        return self.end_cell_row


class _CellIdManipulator:
    def __init__(self, cell_id: str):
        self.cell_id = cell_id

    def _get_row_id_start_position(self) -> int:
        position = 0

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
        return self.get_column() + str(self.get_row() + down_steps)


class _MergedCellHelper:
    # REQUIRES: cell_id must be a vertically merged cell
    def __init__(self, sheet: Worksheet, cell_id: str):
        self.sheet = sheet
        self.cell_id = cell_id

        if type(sheet[cell_id]).__name__ != "MergedCell":
            raise UnacceptableMergedCellError(
                f"cell_id {cell_id} of sheet {sheet} is not a merged cell"
            )

    def get_bottom(self) -> str:
        unmerged_search = self.cell_id

        while type(self.sheet[unmerged_search]).__name__ == "MergedCell":
            unmerged_search = _CellIdManipulator(unmerged_search).get_below()

        return _CellIdManipulator(unmerged_search).get_column() + str(
            _CellIdManipulator(unmerged_search).get_row() - 1
        )


class Helper:
    def __init__(self, sheet: Worksheet):
        self.sheet = sheet

    # REQUIRES: cell_content may only be merged vertically
    # EFFECTS: if cell_content is merged vertically, the bottom most cell will be returned
    def _find(self, cell_content: str) -> str:
        upper_column_range = models.rules.SPOT_SEARCH_COLUMNS + 1
        upper_row_range = models.rules.SPOT_SEARCH_ROWS + 1

        for column in range(1, upper_column_range):
            for row in range(1, upper_row_range):

                current_cell_id = get_column_letter(column) + str(row)
                current_cell = self.sheet[current_cell_id]

                if current_cell.value == cell_content:
                    if type(current_cell).__name__ != "MergedCell":
                        return current_cell_id
                    return _MergedCellHelper(
                        self.sheet, current_cell_id
                    ).get_bottom()

        raise NotFoundError(
            f"couldn't find cell with the cell content of {cell_content} "
            + f"after searching {models.rules.SPOT_SEARCH_ROWS} rows "
            + f"and {models.rules.SPOT_SEARCH_COLUMNS} columns"
        )

    # REQUIRES: cell containing label may only be merged vertically
    def get_column(self, label: str) -> str:
        cell_id = self._find(label)

        return _CellIdManipulator(cell_id).get_column()

    class _ValueColumnTraverser:
        def __init__(self, sheet, label_cell_id):
            self.sheet = sheet
            self.label_cell_id = label_cell_id

            self.last_non_empty_cell = None
            self.empty_cells_streak = 0

        def _update_states(self, base_cell: str) -> None:
            below_last_visited_cell = _CellIdManipulator(base_cell).get_below(1)

            if not self.sheet[below_last_visited_cell].value:
                self.empty_cells_streak += 1
            else:
                self.last_non_empty_cell = below_last_visited_cell
                self.empty_cells_streak = 0

        def _has_reached_end(self) -> bool:
            self.empty_cells_streak < models.rules.MAX_EMPTY_ROWS_FOR_VALUE_COLUMN

        def _get_last_non_empty_cell(self) -> str:
            return self.last_non_empty_cell

        def _get_current_base_cell(self) -> str:
            if not self.last_non_empty_cell:
                return self.label_cell_id
            return _CellIdManipulator(self.last_non_empty_cell).get_below(
                self.empty_cells_streak + 1
            )

        def get_last_cell_id(self) -> int:
            while (
                self.empty_cells_streak < models.rules.MAX_EMPTY_ROWS_FOR_VALUE_COLUMN
            ):
                self._update_states(self._get_current_base_cell())

            if not self.last_non_empty_cell:
                raise NoValueError(
                    f"the rows below the label {self.sheet[self.label_cell_id]} are empty"
                )

            return self.last_non_empty_cell

    def create_value_column(self, label: str) -> ValueColumn:
        cell_id = self._find(label)

        last_non_empty_cell = self._ValueColumnTraverser(
            self.sheet, cell_id
        ).get_last_cell_id()

        column = _CellIdManipulator(cell_id).get_column()
        start_row_id = _CellIdManipulator(cell_id).get_below(1)
        start_row = _CellIdManipulator(start_row_id).get_row()
        end_row = _CellIdManipulator(last_non_empty_cell).get_row()

        return ValueColumn(self.sheet, label, column, start_row, end_row)
