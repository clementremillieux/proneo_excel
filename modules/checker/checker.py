"""_summary_"""

from typing import List

from modules.cells.schemas import BoxToCheck

from modules.excel import excel_handler

from modules.condition.condition import CellsConditions, ConditionHasToBeFilled

from modules.condition.schemas import CellsConditionReport


class Checker:
    """_summary_
    """

    def __init__(self, cells_conditions: List[CellsConditions]) -> None:
        self.nc_names: List[str] = []

        self.cells_conditions: List[CellsConditions] = cells_conditions

        self.nc_cells_conditions: List[CellsConditions] = []

    def check_cells_conditions(self) -> List[CellsConditionReport]:
        """_summary_"""

        self.handle_nc()

        cells_conditions_reports: List[CellsConditionReport] = []

        for cell_conditions in self.cells_conditions:
            cells_conditions_reports.append(cell_conditions.check())

        for nc_cell_conditions in self.nc_cells_conditions:
            cells_conditions_reports.append(nc_cell_conditions.check())

        return cells_conditions_reports

    def handle_nc(self) -> None:
        """_summary_"""

        self.nc_cells_conditions = []

        all_sheets: List[str] = excel_handler.get_all_sheets()

        nc_sheets: List[str] = [
            sheet for sheet in all_sheets if sheet.startswith("NC ind")
        ]

        self.create_nc_condition(nc_sheets=nc_sheets)

    def create_nc_condition(self, nc_sheets: List[str]) -> None:
        """_summary_"""

        for sheet_name in nc_sheets:
            self.nc_cells_conditions.append(
                self.create_date_filled_condition(sheet_name=sheet_name))

    def create_pass_or_failed_condition(self, sheet_name: str) -> None:
        """_summary_
        """

        pass

    def create_date_filled_condition(self, sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_
        """

        cell = BoxToCheck(sheet_name=sheet_name, cell_address="B46")

        date_filled = ConditionHasToBeFilled(cell=cell,
                                             is_parent_condition=False)

        cells_conditions_date_filled = CellsConditions(
            conditions=[date_filled])

        return cells_conditions_date_filled
