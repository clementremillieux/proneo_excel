"""_summary_"""

from datetime import datetime as dt

from typing import List, Optional

from modules.cells.schemas import BoxToCheck, DateToCheck

from modules.condition.schemas import Condition, ConditionType, CellsConditionReport, CellsConditionState


class CellsConditions:
    """_summary_

    Returns:
        _type_: _description_
    """

    def __init__(self, conditions: List[Condition]) -> None:
        self.conditions: List[Condition] = conditions

    def check(self) -> CellsConditionReport:
        """_summary_

        Returns:
            _type_: _description_
        """

        for condition in self.conditions:
            cells_condition_report = condition.check()

            if cells_condition_report.state == CellsConditionState.NOT_OK and not condition.is_parent_condition:
                break

        return cells_condition_report


class ConditionDateSup(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell_date_start: DateToCheck

    cell_date_stop: DateToCheck

    def __init__(self, cell_date_start: DateToCheck,
                 cell_date_stop: DateToCheck,
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.DATE_SUP,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell_date_start, cell_date_stop])

        self.cell_date_start: DateToCheck = cell_date_start

        self.cell_date_stop: DateToCheck = cell_date_stop

    def check(self) -> CellsConditionReport:
        date_start_cell_value: Optional[dt] = self.cell_date_start.get_value()

        date_stop_cell_value: Optional[dt] = self.cell_date_stop.get_value()

        if not date_start_cell_value or not date_stop_cell_value:

            results: bool = True

        else:

            date_start: dt = date_start_cell_value

            date_stop: dt = date_stop_cell_value

            results = date_stop >= date_start

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        if state == CellsConditionState.OK:
            report_str = f"La date de la cellule {self.cell_date_stop.cell_address} [{self.cell_date_stop.sheet_name}] \
et de la cellule {self.cell_date_start.cell_address} [{self.cell_date_start.sheet_name}] correspondent"

        else:
            report_str = f"La date de la cellule {self.cell_date_start.cell_address} [{self.cell_date_start.sheet_name}] \
doit être antérieur à celle de la cellule {self.cell_date_stop.cell_address} [{self.cell_date_stop.sheet_name}]"

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionHasToBeFilled(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    def __init__(self, cell: BoxToCheck, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.DATE_SUP,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: BoxToCheck = cell

    def check(self) -> CellsConditionReport:
        cell_value: Optional[str] = self.cell.get_value()

        if not cell_value:
            results: bool = False

        else:
            results = True

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        if state == CellsConditionState.NOT_OK:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] doit être remplie"

        else:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] est remplie"

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report
