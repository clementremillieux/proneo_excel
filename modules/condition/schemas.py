"""_summary_"""
from __future__ import annotations

from abc import abstractmethod

from enum import Enum

from typing import List

from pydantic import BaseModel

from modules.cells.schemas import BoxToCheck, CheckBoxToCheck, DateToCheck


class ConditionType(Enum):
    """_summary_

    Args:
        Enum (_type_): _description_
    """

    DATE_SUP = "date_sup"

    DATE_DURATION_BEETWEEN = "date_duration_beetween"

    CELL_HAS_TO_BE_VALUE = "cell_has_to_be_value"

    CELL_HAS_TO_BE_FILLED = "cell_has_to_be_filled"

    CELL_HAS_TO_BE_FILLED_IF_VALUE_FROM_OTHER = "cell_has_to_be_filled_if_value_from_other"

    ALL_CELLS_HAS_TO_BE_CHECKED_IN_LIST = "all_cells_has_to_be_checked_in_list"

    AT_LEAST_ONE_BOX_CHECKED_AMONG_LIST = "at_least_one_box_checked_among_list"

    ONE_BOX_CHECKED_AMONG_LIST = "one_box_checked_among_list"

    MAX_ONE_BOX_CHECKED_AMONG_LIST = "max_one_box_checked_among_list"

    CHECK_ALL_SHEET_DESCRIPTION = "check_all_sheet_description"

    CHECK_ALL_SHEET_REFERENCE = "check_all_sheet_reference"

    CHECK_ALL_SHEET_DROP_DOWN = "check_all_sheet_dropdown"

    CHECK_NC_ALL_J_CHOOSED = "check_nc_all_j_choosed"

    HAS_NC = "has_nc"


class Condition():
    """_summary_

    Args:
        BaseModel (_type_): _description_
    """

    def __init__(
            self, condition_type: ConditionType, is_parent_condition: bool,
            cells_list: List[DateToCheck | CheckBoxToCheck | BoxToCheck]
    ) -> None:

        self.condition_type: ConditionType = condition_type

        self.is_parent_condition: bool = is_parent_condition

        self.cells_list: List[DateToCheck | CheckBoxToCheck
                              | BoxToCheck] = cells_list

    @abstractmethod
    def check(self) -> CellsConditionReport:
        """Check the condition based on string inputs.

        Args:
            *args (str): One or more string arguments.

        Returns:
            bool: True if the condition is met, False otherwise.
        """


class CellsConditionState(Enum):
    """_summary_

    Args:
        Enum (_type_): _description_
    """

    OK = "ok"

    NOT_OK = "not_ok"


class CellsConditionReport(BaseModel):
    """_summary_

    Args:
        BaseModel (_type_): _description_
    """

    condition: Condition

    state: CellsConditionState

    report_str: str

    class Config:
        """_summary_"""

        arbitrary_types_allowed = True
