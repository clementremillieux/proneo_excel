"""_summary_"""

import re

import copy

from datetime import datetime as dt

from typing import Dict, List, Optional

from modules.cells.schemas import BoxToCheck, CheckBoxToCheck, DateToCheck

from modules.condition.schemas import Condition, ConditionType, CellsConditionReport, CellsConditionState

from modules.excel import excel_handler

from modules.excel.schemas import CheckboxParams

from modules.performances.time_counter import time_execution

from modules.sheet.schemas import SheetName

from config.logger_config import logger

checkbox_params = CheckboxParams(
    apple_script_path="modules/excel/apple_script/checkbox.scpt")

REF_TO_NB_INDIC = {
    "V1": 21,
    "V2": 17,
    "V3": 21,
    "V4": 21,
    "V5": 21,
    "V6": 17,
    "V7": 17,
    "V8": 21,
    "V9": 13,
    "V10": 21,
    "V11": 21,
    "V12": 17,
    "V13": 11,
    "V14": 13,
    "V15": 21,
    "V16": 30,
    "V17": 34,
    "V18": 34,
    "V19": 30,
    "V20": 30,
    "V21": 34,
    "V22": 30,
    "V23": 34,
    "V24": 34,
    "V25": 26,
    "V26": 34,
    "V27": 34,
    "V28": 24,
    "V29": 26,
    "V30": 34
}


class CellsConditions:
    """_summary_

    Returns:
        _type_: _description_
    """

    def __init__(self, conditions: List[Condition]) -> None:
        self.conditions: List[Condition] = conditions

        self.parent_condition_ok: str = ""

    def check(
            self
    ) -> Optional[CellsConditionReport | List[CellsConditionReport]]:
        """_summary_

        Returns:
            _type_: _description_
        """

        try:

            for condition in self.conditions:
                cells_condition_report: CellsConditionReport | List[
                    CellsConditionReport] = condition.check()

                if isinstance(
                        cells_condition_report, CellsConditionReport
                ) and cells_condition_report.state == CellsConditionState.NOT_OK:
                    if not condition.is_parent_condition:
                        break

                    return None

                if isinstance(cells_condition_report, List) and any(
                        cell_condition_report.state ==
                        CellsConditionState.NOT_OK
                        for cell_condition_report in cells_condition_report):
                    if not condition.is_parent_condition:
                        break

                    return None

                if condition.is_parent_condition:
                    parent_condition_str: Optional[
                        str] = condition.get_parent_condition_str()

                    if parent_condition_str:

                        self.parent_condition_ok += parent_condition_str

        except Exception as e:
            logger.warning("Error checking condition %s : %s",
                           condition.condition_type.value, e)

            cells_condition_report = CellsConditionReport(
                condition=condition,
                state=CellsConditionState.NOT_OK,
                report_str="Error interne")

        if len(self.parent_condition_ok) > 0:

            if isinstance(cells_condition_report, CellsConditionReport):
                cells_condition_report.report_str += self.parent_condition_ok

            if isinstance(cells_condition_report, List):

                for index, cell_report in enumerate(cells_condition_report):

                    cells_condition_report[
                        index].report_str = cell_report.report_str + self.parent_condition_ok

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

        # logger.info("CONDITION DATE SUP => %s | %s", date_start_cell_value,
        #             date_stop_cell_value)

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

        # logger.info("CONDITION DATE SUP => %s", report_str)

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionDateDurationBetween(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell_date_start: DateToCheck

    cell_date_stop: DateToCheck

    cell_duration: BoxToCheck

    def __init__(self, cell_date_start: DateToCheck,
                 cell_date_stop: DateToCheck, cell_duration: BoxToCheck,
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.DATE_DURATION_BEETWEEN,
            is_parent_condition=is_parent_condition,
            cells_list=[cell_date_start, cell_date_stop, cell_duration])

        self.cell_date_start: DateToCheck = cell_date_start

        self.cell_date_stop: DateToCheck = cell_date_stop

        self.cell_duration: BoxToCheck = cell_duration

    def check(self) -> CellsConditionReport:
        date_start_cell_value: Optional[dt] = self.cell_date_start.get_value()

        date_stop_cell_value: Optional[dt] = self.cell_date_stop.get_value()

        duration_cell_value: Optional[str] = self.cell_duration.get_value()

        if not date_start_cell_value or not date_stop_cell_value or not duration_cell_value:

            results: bool = False

            report_str = f"Une des cellule {self.cell_date_stop.cell_address} [{self.cell_date_stop.sheet_name}] et/ou \
{self.cell_date_start.cell_address} [{self.cell_date_start.sheet_name}] et/ou  {self.cell_duration.cell_address} [{self.cell_duration.sheet_name}] \
n'est pas remplies"

        else:

            try:
                duration_cell_value_int: int = int(duration_cell_value)

                date_start: dt = date_start_cell_value

                date_stop: dt = date_stop_cell_value

                results = (date_stop -
                           date_start).days >= duration_cell_value_int

                if results:
                    report_str = f"La durée de l'audit indiqué à la cellule {self.cell_duration.cell_address} [{self.cell_duration.sheet_name}]\
     correspond aux dates de la cellule {self.cell_date_stop.cell_address} [{self.cell_date_stop.sheet_name}]\
     et de la cellule {self.cell_date_start.cell_address} [{self.cell_date_start.sheet_name}]"

                else:
                    report_str = f"La durée l'audit indiqué à la cellule {self.cell_duration.cell_address} [{self.cell_duration.sheet_name}]\
     ne correspond pas aux dates {self.cell_date_stop.cell_address} [{self.cell_date_stop.sheet_name}]\
     et de la cellule {self.cell_date_start.cell_address} [{self.cell_date_start.sheet_name}]"

            except Exception as e:

                logger.error(e)

                results = False

                report_str = f"La valuer de la cellule {self.cell_duration.cell_address} [{self.cell_duration.sheet_name}] n'est pas un nombre"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionOneCheckBoxAmongList(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cells: List[CheckBoxToCheck]

    only_check: bool

    def __init__(self,
                 cells: List[CheckBoxToCheck],
                 is_parent_condition: bool,
                 only_check: bool = False) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.ONE_BOX_CHECKED_AMONG_LIST,
            is_parent_condition=is_parent_condition,
            cells_list=cells)

        self.cells: List[CheckBoxToCheck] = cells

        self.only_check: bool = only_check

    def check(self) -> CellsConditionReport:
        cells_value: List[Optional[bool]] = [
            cell.get_value() for cell in self.cells
        ]

        checkbox_name: str = " ou ".join(
            f'{cell.alias_name if cell.alias_name else cell.cell_address} [{cell.sheet_name}]'
            for cell in self.cells)

        if not any(cells_value) and not self.only_check:
            results: bool = False

            report_str = f"Une des check box {checkbox_name} doit être cochée"

        elif sum(1 for value in cells_value if value) > 1:
            results = False

            report_str = f"Seule une des check box {checkbox_name} doit être cochée"

        else:
            results = True

            report_str = f"Une des check box {checkbox_name} à bien été cochée"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report

    def get_parent_condition_str(self) -> str:
        """_summary_

        Returns:
            str: _description_
        """

        checkbox_name: str = " ou ".join(
            f'{cell.alias_name if cell.alias_name else cell.cell_address} [{cell.sheet_name}]'
            for cell in self.cells)

        return f". Car la checkbox {checkbox_name} est cochée."


class ConditionAtLeastOneCheckBoxAmongList(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cells: List[CheckBoxToCheck]

    def __init__(self,
                 cells: List[CheckBoxToCheck],
                 is_parent_condition: bool,
                 alias_name: Optional[str] = None) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.AT_LEAST_ONE_BOX_CHECKED_AMONG_LIST,
            is_parent_condition=is_parent_condition,
            cells_list=cells,
            alias_name=alias_name)

        self.cells: List[CheckBoxToCheck] = cells

    def check(self) -> CellsConditionReport:
        cells_value: List[Optional[bool]] = [
            cell.get_value() for cell in self.cells
        ]

        checkbox_name: str = " ou ".join(
            f'{cell.alias_name if cell.alias_name else cell.cell_address} [{cell.sheet_name}]'
            for cell in self.cells)

        if not any(cells_value):
            results: bool = False

            report_str = f"Une des checkbox doit être cochée : {checkbox_name}"

        else:
            results = True

            report_str = f"Une des check box à bien été cochée : {checkbox_name}"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report

    def get_parent_condition_str(self) -> str:
        """_summary_

        Returns:
            str: _description_
        """

        checkbox_name: str = " ou ".join(
            f'{cell.alias_name if cell.alias_name else cell.cell_address} [{cell.sheet_name}]'
            for cell in self.cells)

        return f". Car la checkbox {checkbox_name} est cochée."


class ConditionAtLeastOneCellAmongList(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cells: List[BoxToCheck]

    def __init__(self, cells: List[BoxToCheck],
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.AT_LEAST_ONE_BOX_CHECKED_AMONG_LIST,
            is_parent_condition=is_parent_condition,
            cells_list=cells)

        self.cells: List[BoxToCheck] = cells

    def check(self) -> CellsConditionReport:
        cells_value: List[Optional[bool]] = [
            bool(cell.get_value()) for cell in self.cells
        ]

        checkbox_name: str = " ou ".join(
            f'{cell.cell_address} [{cell.sheet_name}]' for cell in self.cells)

        if not any(cells_value):
            results: bool = False

            report_str = f"Une des cellules doit être remplie : {checkbox_name}"

        else:
            results = True

            report_str = f"Une des cellules à bien été remplie : {checkbox_name}"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionHasToBeFilled(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    size_siren: Optional[int] = 0

    size_nda: Optional[int] = 0

    size_phone: Optional[int] = 0

    sentence_to_remove: Optional[str] = ""

    def __init__(self,
                 cell: BoxToCheck,
                 is_parent_condition: bool,
                 size_siren: Optional[int] = 0,
                 size_nda: Optional[int] = 0,
                 size_phone: Optional[int] = 0,
                 sentence_to_remove: Optional[str] = "") -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_FILLED,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: BoxToCheck = cell

        self.size_siren: Optional[int] = size_siren

        self.size_nda: Optional[int] = size_nda

        self.size_phone: Optional[int] = size_phone

        self.sentence_to_remove: Optional[str] = sentence_to_remove

    def check(self) -> CellsConditionReport:
        cell_value: Optional[str] = self.cell.get_value()

        if not cell_value:
            results: bool = False

        else:
            results = True

            if self.sentence_to_remove:

                if len(cell_value.strip().lower().replace(
                        self.sentence_to_remove.lower(),
                        "").replace(" ", "").replace("\n", "")) == 0:
                    results = False

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        if state == CellsConditionState.NOT_OK:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] doit être remplie"

        else:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] est remplie"

        if self.size_siren and cell_value:

            try:
                int(cell_value)

                if len(str(int(cell_value))) != self.size_siren:
                    state = CellsConditionState.NOT_OK

                    report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] ne correspond pas à un numéro SIREN ({self.size_siren} chiffres)"

            except Exception:
                state = CellsConditionState.NOT_OK

                report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] ne correspond pas à un numéro SIREN ({self.size_siren} chiffres)"

        if self.size_nda and cell_value:

            try:
                int(cell_value)

                if len(str(int(cell_value))) != self.size_nda:
                    state = CellsConditionState.NOT_OK

                    report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] ne correspond pas à un numéro NDA ({self.size_nda} chiffres)"

            except Exception:
                state = CellsConditionState.NOT_OK

                report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] ne correspond pas à un numéro NDA ({self.size_nda} chiffres)"

        if self.size_phone and cell_value:

            cell_value = cell_value.replace(" ", "").replace("_", "").replace(
                "-", "").replace(",", "").replace(";", "")

            try:
                int(cell_value)

                clean_phone = str(int(cell_value))

                if len(clean_phone) != self.size_phone - 1:
                    state = CellsConditionState.NOT_OK

                    report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] ne correspond pas à un numéro de téléphone ({self.size_phone} chiffres)"

            except Exception:
                state = CellsConditionState.NOT_OK

                report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] ne correspond pas à un numéro de téléphone ({self.size_phone} chiffres 0XXXXXXXXX)"

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report

    def get_parent_condition_str(self) -> str:
        """_summary_

        Returns:
            str: _description_
        """

        return f". Car la cellule {self.cell.cell_address} est remplie."


class ConditionHasToBeChecked(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: CheckBoxToCheck

    def __init__(self, cell: CheckBoxToCheck,
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_FILLED,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: CheckBoxToCheck = cell

    def check(self) -> CellsConditionReport:
        cell_value: Optional[bool] = self.cell.get_value()

        if not cell_value:
            results: bool = False

        else:
            results = True

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        if state == CellsConditionState.NOT_OK:
            report_str = f"La checkbox {self.cell.alias_name if  self.cell.alias_name else self.cell.cell_address} [{self.cell.sheet_name}] doit être cochée"

        else:
            report_str = f"La checkbox {self.cell.alias_name if  self.cell.alias_name else self.cell.cell_address} [{self.cell.sheet_name}] est cochée"

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report

    def get_parent_condition_str(self) -> str:
        """_summary_

        Returns:
            str: _description_
        """

        return f". Car la checkbox {self.cell.cell_address} [{self.cell.sheet_name}] est cochée."


class ConditionOneByChecked(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    possible_cells: List[CheckBoxToCheck]

    cells_to_check: List[CheckBoxToCheck]

    def __init__(self, cells_to_check: List[CheckBoxToCheck],
                 possible_cells: List[CheckBoxToCheck],
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_FILLED,
                         is_parent_condition=is_parent_condition,
                         cells_list=cells_to_check)

        self.cells_to_check: List[CheckBoxToCheck] = cells_to_check

        self.possible_cells: List[CheckBoxToCheck] = possible_cells

    def check(self) -> CellsConditionReport:

        # logger.info("Check:")

        # for cell in self.cells_to_check:

        #     logger.warning(cell.get_value())

        # logger.info("Possible:")

        # for cell in self.possible_cells:

        #     logger.warning(cell.get_value())

        if all(not cell.get_value() for cell in self.possible_cells):

            return CellsConditionReport(condition=copy.deepcopy(self),
                                        state=CellsConditionState.OK,
                                        report_str="")

        if all(not cell.get_value() for cell in self.cells_to_check):

            checkbox_name: str = " ou ".join(
                f'{cell.alias_name} [{cell.sheet_name}]'
                for cell in self.cells_to_check if cell.alias_name in [
                    cell.alias_name for cell in self.possible_cells
                    if cell.get_value()
                ])

            return CellsConditionReport(
                condition=copy.deepcopy(self),
                state=CellsConditionState.NOT_OK,
                report_str=f"Une des checkbox {checkbox_name} doit être cochés"
            )

        if not all(cell.alias_name in [
                cell.alias_name
                for cell in self.possible_cells if cell.get_value()
        ] for cell in self.cells_to_check if cell.get_value()):

            checkbox_name: str = " ou ".join(
                f'{cell.alias_name} [{cell.sheet_name}]'
                for cell in self.cells_to_check if cell.alias_name in [
                    cell.alias_name for cell in self.possible_cells
                    if cell.get_value()
                ])

            return CellsConditionReport(
                condition=copy.deepcopy(self),
                state=CellsConditionState.NOT_OK,
                report_str=
                f"Seules les checkbox {checkbox_name} doit être cochés")

        return CellsConditionReport(condition=copy.deepcopy(self),
                                    state=CellsConditionState.OK,
                                    report_str="")


class ConditionHasToBeSigned(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    def __init__(self, cell: BoxToCheck, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_FILLED,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: BoxToCheck = cell

    def check(self) -> CellsConditionReport:

        if not excel_handler.cell_contains_signature(
                sheet_name=self.cell.sheet_name,
                cell_address=self.cell.cell_address):
            results: bool = False

        else:
            results = True

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        if state == CellsConditionState.NOT_OK:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] doit contenir une signature"

        else:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] contient une signature"

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionHasToBeValues(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    value: List[str]

    def __init__(self, cell: BoxToCheck, value: List[str],
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_VALUE,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: BoxToCheck = cell

        self.value: List[str] = value

    def check(self) -> CellsConditionReport:
        cell_value: Optional[str] = self.cell.get_value()

        if not cell_value or cell_value not in self.value:
            results: bool = False

        else:
            results = True

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        if state == CellsConditionState.NOT_OK:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] doit être égale à {self.value}"

        else:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] est égale à {self.value}"

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report

    def get_parent_condition_str(self) -> str:
        """_summary_

        Returns:
            str: _description_
        """

        return f". Car la cellule {self.cell.cell_address}  [{self.cell.sheet_name}] vaut {self.value}."


class ConditionIsNCFromCellText(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    def __init__(self, cell: BoxToCheck, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_VALUE,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: BoxToCheck = cell

    def check(self) -> CellsConditionReport:
        cell_value: Optional[str] = self.cell.get_value()

        try:
            cell_value_str = str(cell_value)

        except Exception as e:

            logger.info(
                "Error converting cell value to str for NC from cell text : %s",
                e)

            return CellsConditionReport(
                condition=self,
                state=CellsConditionState.NOT_OK,
                report_str=
                f"Les non conformités de la cellule {self.cell.cell_address} sont mal écrites"
            )

        ids_nc: List[str] = extract_ids(input_text=cell_value_str)

        all_sheets: List[str] = excel_handler.get_all_sheets()

        results: bool = True

        report_str: str = ""

        for id_nc in ids_nc:
            nc_sheet = f"NC ind {id_nc}"

            if nc_sheet not in all_sheets:
                logger.info("%s is missing in sheet name", nc_sheet)

                results = False

                report_str += f"La fiche NC {id_nc} n'a pas été créée\n"

                break

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        if state == CellsConditionState.OK:
            report_str = f"Toutes les fiches NC {ids_nc} indiquée à la cellule {self.cell.cell_address} ont été créée"

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionIsNCFromCellNumber(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    def __init__(self, cell: BoxToCheck, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_VALUE,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: BoxToCheck = cell

    def check(self) -> CellsConditionReport:
        cell_value: Optional[str] = self.cell.get_value()

        try:
            cell_value_int = int(cell_value)

        except Exception as e:

            logger.info(
                "Error converting cell value to str for NC from cell text : %s",
                e)

            return CellsConditionReport(
                condition=self,
                state=CellsConditionState.NOT_OK,
                report_str=
                f"Les nombre de NC mineure de la cellule {self.cell.cell_address} n'est pas un nombre"
            )

        #nb_nc_min_sheet: int = count_nc_min_sheet()

        nb_nc_min: int = count_nc_min()

        # if cell_value_int != nb_nc_min_sheet:
        #     results: bool = False

        #     report_str: str = f"Le nombre de NC mineure définie à la cellule {self.cell.cell_address} ({cell_value_int}) ne correspond pas au nombre de fiche NC mineur créée ({nb_nc_min_sheet})"

        if cell_value_int != nb_nc_min:
            results = False

            report_str = f"Le nombre de NC mineure définie à la cellule {self.cell.cell_address} ({cell_value_int}) ne correspond pas au nombre de NC mineure définie dans la rapport d'audit ({nb_nc_min})"

        else:
            results = True

            report_str = f"Le nombre de NC mineure définie à la cellule {self.cell.cell_address} ({cell_value_int}) correspond au nombre de NC mineure définie dans la rapport d'audit ({nb_nc_min})"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionIsNcMajFromCellNumber(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    def __init__(self, cell: BoxToCheck, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_VALUE,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: BoxToCheck = cell

    def check(self) -> CellsConditionReport:
        cell_value: Optional[str] = self.cell.get_value()

        try:
            cell_value_int = int(cell_value)

        except Exception as e:

            logger.info(
                "Error converting cell value to str for NC from cell text : %s",
                e)

            return CellsConditionReport(
                condition=self,
                state=CellsConditionState.NOT_OK,
                report_str=
                f"Les nombre de NC de la cellule {self.cell.cell_address} n'est pas un nombre"
            )

        nb_nc_maj: int = count_nc_maj()

        if cell_value_int == nb_nc_maj:
            results: bool = True

            report_str: str = f"Le nombre de NC majeure définie à la cellule {self.cell.cell_address} ({cell_value_int}) correspond au nombre de NC majeure dans le Rapport d'audit ({nb_nc_maj})"

        else:
            results = False

            report_str = f"Le nombre de NC majeure définie à la cellule {self.cell.cell_address} ({cell_value_int}) ne correspond pas au nombre de NC majeure dans le Rapport d'audit ({nb_nc_maj})"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionHasNc(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    def __init__(self, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.HAS_NC,
                         is_parent_condition=is_parent_condition,
                         cells_list=[])

    def check(self) -> CellsConditionReport:

        nb_nc_min: int = count_nc_min()

        nb_nc_maj: int = count_nc_maj()

        if nb_nc_min > 0 or nb_nc_maj > 0:
            results: bool = True

        else:
            results = False

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str="")

        return cells_report

    def get_parent_condition_str(self) -> str:
        """_summary_

        Returns:
            str: _description_
        """

        return f". Car il y a des non conformités."


class ConditionNcAllJChoosed(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    def __init__(self, cell: BoxToCheck, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CHECK_NC_ALL_J_CHOOSED,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell = cell

    @time_execution
    def check(self) -> CellsConditionReport:

        ref: str = get_ref()

        nb_inc_for_ref: int = REF_TO_NB_INDIC[ref]

        nb_not_none: int = count_not_none_in_nc_j()

        if nb_inc_for_ref > nb_not_none:
            results: bool = False

            report_str: str = "Toutes les conformités des indicateurs doivent être définies [Rapport d'audit :colonne J]"

        else:
            results = True

            report_str = "Toutes les conformités des indicateurs sont définies [Rapport d'audit :colonne J]"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=copy.deepcopy(self),
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionCheckAllSheetDropDown(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    sheet_name: str

    no_na_cells: Dict[str, List[str]]

    def __init__(self, sheet_name: str, is_parent_condition: bool,
                 no_na_cells: Dict[str, List[str]]) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.CHECK_ALL_SHEET_DROP_DOWN,
            is_parent_condition=is_parent_condition,
            cells_list=[
                BoxToCheck(
                    sheet_name=SheetName.SHEET_5.value,
                    cell_address="L5",
                )
            ])

        self.sheet_name = sheet_name

        self.no_na_cells: Dict[str, List[str]] = no_na_cells

    @time_execution
    def check(self) -> List[CellsConditionReport]:

        report_str: str = ""

        cells_reports: List[CellsConditionReport] = []

        current_j_value: Optional[str] = None

        for column in ["L", "M", "N", "O", "P", "Q", "R", "S"]:
            for row in range(5, 186):

                if not excel_handler.is_merged(
                        sheet_name=SheetName.SHEET_5.value,
                        cell_adress=f"J{row}"):
                    current_j_value = excel_handler.read_cell_value(
                        sheet_name=SheetName.SHEET_5.value,
                        cell_address=f"J{row}")

                cell_adress = f"{column}{row}"

                if current_j_value in [
                        "Non-conformité mineure", "Non-conformité majeure",
                        "Conformité", "None"
                ] or current_j_value is None:

                    if not excel_handler.is_line_hidden(
                            sheet_name=self.sheet_name, cell_adress=cell_adress
                    ) and not excel_handler.is_column_hidden(
                            sheet_name=self.sheet_name,
                            cell_adress=cell_adress):

                        if excel_handler.is_drop_down(
                                sheet_name=self.sheet_name,
                                cell_adress=cell_adress):
                            cell = BoxToCheck(
                                sheet_name=SheetName.SHEET_5.value,
                                cell_address=cell_adress,
                            )

                            if not cell.get_value():

                                if cell_adress in self.no_na_cells.keys():

                                    for adress_to_check in self.no_na_cells[
                                            cell_adress]:
                                        box_to_check = CheckBoxToCheck(
                                            sheet_name=SheetName.SHEET_2.value,
                                            checkbox_name=adress_to_check,
                                            cell_address=adress_to_check,
                                            checkbox_params=checkbox_params)

                                        if box_to_check.get_value():
                                            report_str = f"La valeur choisie pour la cellule {cell_adress} [{self.sheet_name}] doit être 'Oui' ou 'Non' car la checkbox {adress_to_check} [{SheetName.SHEET_2.value}] est cochée"

                                            break

                                        report_str = f"Une valeur doit être choisie pour la cellule {cell_adress} [{self.sheet_name}]"

                                else:
                                    report_str = f"Une valeur doit être choisie pour la cellule {cell_adress} [{self.sheet_name}]"

                                self.cells_list = [cell]

                                cells_reports.append(
                                    CellsConditionReport(
                                        condition=copy.deepcopy(self),
                                        state=CellsConditionState.NOT_OK,
                                        report_str=report_str))

                            elif cell_adress in self.no_na_cells.keys():
                                for adress_to_check in self.no_na_cells[
                                        cell_adress]:
                                    box_to_check = CheckBoxToCheck(
                                        sheet_name=SheetName.SHEET_2.value,
                                        checkbox_name=adress_to_check,
                                        cell_address=adress_to_check,
                                        checkbox_params=checkbox_params)

                                    if cell.get_value() not in ["Oui", "Non"]:
                                        report_str = f"La valeur choisie pour la cellule {cell_adress} [{self.sheet_name}] doit être 'Oui' ou 'Non'"

                                        self.cells_list = [cell]

                                        cells_reports.append(
                                            CellsConditionReport(
                                                condition=copy.deepcopy(self),
                                                state=CellsConditionState.
                                                NOT_OK,
                                                report_str=report_str))

                                        break

                            if excel_handler.read_cell_value(
                                    sheet_name=SheetName.SHEET_5.value,
                                    cell_address=cell.cell_address
                            ) == "Non" and current_j_value == "Conformité":

                                self.cells_list = [cell]

                                cells_reports.append(
                                    CellsConditionReport(
                                        condition=copy.deepcopy(self),
                                        state=CellsConditionState.NOT_OK,
                                        report_str=
                                        f"La cellule J{row} ne peut pas être conforme car la cellule {cell_adress} [{self.sheet_name}] est Non"
                                    ))

        return cells_reports


class ConditionCheckAllSheetReference(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    sheet_name: str

    def __init__(self, sheet_name: str, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.CHECK_ALL_SHEET_REFERENCE,
            is_parent_condition=is_parent_condition,
            cells_list=[
                BoxToCheck(
                    sheet_name=SheetName.SHEET_5.value,
                    cell_address="L5",
                )
            ])

        self.sheet_name = sheet_name

    @time_execution
    def check(self) -> List[CellsConditionReport]:

        cells_reports: List[CellsConditionReport] = []

        report_str: str = ""

        references_cells: List[str] = get_references_cells()

        for cell in references_cells:
            if not excel_handler.read_cell_value(
                    sheet_name=SheetName.SHEET_5.value, cell_address=cell):
                report_str = f"La cellule {cell} [{SheetName.SHEET_5.value}] ne peut pas être vide"

                self.cells_list = [
                    BoxToCheck(
                        sheet_name=SheetName.SHEET_5.value,
                        cell_address=cell,
                    )
                ]

                cells_reports.append(
                    CellsConditionReport(condition=copy.deepcopy(self),
                                         state=CellsConditionState.NOT_OK,
                                         report_str=report_str))

        return cells_reports


START_LINE_REPORT_AUDIT = 5

NB_LINE_REPORT_AUDIT = 188


def get_references_cells() -> List[str]:
    """_summary_

    Returns:
        List[str]: _description_
    """

    references_cells: List[str] = []

    is_l_column_hidden: bool = excel_handler.is_column_hidden(
        sheet_name=SheetName.SHEET_5.value, cell_adress="L4")

    is_n_column_hidden: bool = excel_handler.is_column_hidden(
        sheet_name=SheetName.SHEET_5.value, cell_adress="N4")

    is_p_column_hidden: bool = excel_handler.is_column_hidden(
        sheet_name=SheetName.SHEET_5.value, cell_adress="P4")

    is_r_column_hidden: bool = excel_handler.is_column_hidden(
        sheet_name=SheetName.SHEET_5.value, cell_adress="R4")

    current_j_value: Optional[str] = None

    current_b_value: Optional[str] = None

    current_c_value: Optional[str] = None

    current_d_value: Optional[str] = None

    current_e_value: Optional[str] = None

    for row in range(START_LINE_REPORT_AUDIT, NB_LINE_REPORT_AUDIT):

        if not excel_handler.is_merged(sheet_name=SheetName.SHEET_5.value,
                                       cell_adress=f"J{row}"):

            current_j_value = excel_handler.read_cell_value(
                sheet_name=SheetName.SHEET_5.value, cell_address=f"J{row}")

            current_b_value = excel_handler.read_cell_value(
                sheet_name=SheetName.SHEET_5.value, cell_address=f"B{row}")

            current_c_value = excel_handler.read_cell_value(
                sheet_name=SheetName.SHEET_5.value, cell_address=f"C{row}")

            current_d_value = excel_handler.read_cell_value(
                sheet_name=SheetName.SHEET_5.value, cell_address=f"D{row}")

            current_e_value = excel_handler.read_cell_value(
                sheet_name=SheetName.SHEET_5.value, cell_address=f"E{row}")

        if current_j_value in [
                "Non-conformité mineure", "Non-conformité majeure",
                "Conformité", "None"
        ] or current_j_value is None:
            cell = f"K{row}"

            cell_value: Optional[str] = excel_handler.read_cell_value(
                sheet_name=SheetName.SHEET_5.value, cell_address=cell)

            if cell_value and "Références" in cell_value and not excel_handler.is_line_hidden(
                    sheet_name=SheetName.SHEET_5.value, cell_adress=cell):
                if not is_l_column_hidden and current_b_value and current_b_value.lower(
                ) == "x":
                    references_cells.append(f"L{row}")

                if not is_n_column_hidden and current_c_value and current_c_value.lower(
                ) == "x":
                    references_cells.append(f"N{row}")

                if not is_p_column_hidden and current_d_value and current_d_value.lower(
                ) == "x":
                    references_cells.append(f"P{row}")

                if not is_r_column_hidden and current_e_value and current_e_value.lower(
                ) == "x":
                    references_cells.append(f"R{row}")

            if cell_value and "Description" in cell_value and not excel_handler.is_line_hidden(
                    sheet_name=SheetName.SHEET_5.value, cell_adress=cell):
                if not is_l_column_hidden and current_b_value and current_b_value.lower(
                ) == "x":
                    references_cells.append(f"L{row}")

                if not is_n_column_hidden and current_c_value and current_c_value.lower(
                ) == "x":
                    references_cells.append(f"N{row}")

                if not is_p_column_hidden and current_d_value and current_d_value.lower(
                ) == "x":
                    references_cells.append(f"P{row}")

                if not is_r_column_hidden and current_e_value and current_e_value.lower(
                ) == "x":
                    references_cells.append(f"R{row}")

    return references_cells


def count_nc_min_sheet() -> int:
    """_summary_

    Returns:
        int: _description_
    """

    all_sheets: List[str] = excel_handler.get_all_sheets()

    nb_nc_min = len([sheet for sheet in all_sheets if "NC ind" in sheet])

    return nb_nc_min


def count_nc_min() -> int:
    """_summary_

    Returns:
        int: _description_
    """

    counter_nc_min: int = 0

    for row in range(START_LINE_REPORT_AUDIT, NB_LINE_REPORT_AUDIT):
        cell_adress = f"J{row}"

        cell = BoxToCheck(
            sheet_name=SheetName.SHEET_5.value,
            cell_address=cell_adress,
        )

        try:

            cell_value: Optional[str] = cell.get_value()

            if cell_value and isinstance(
                    cell_value,
                    str) and cell_value.strip() == "Non-conformité mineure":
                counter_nc_min += 1

        except Exception:
            pass

    return counter_nc_min


def count_nc_maj() -> int:
    """_summary_

    Returns:
        int: _description_
    """

    counter_nc_maj: int = 0

    for row in range(START_LINE_REPORT_AUDIT, NB_LINE_REPORT_AUDIT):
        cell_adress = f"J{row}"

        cell = BoxToCheck(
            sheet_name=SheetName.SHEET_5.value,
            cell_address=cell_adress,
        )

        try:

            cell_value: Optional[str] = cell.get_value()

            if cell_value and isinstance(
                    cell_value,
                    str) and cell_value.strip() == "Non-conformité majeure":
                counter_nc_maj += 1

        except Exception:
            pass

    nb_nc_min = count_nc_min()

    counter_nc_maj += int(nb_nc_min / 5)

    return counter_nc_maj


def get_ref() -> str:
    """_summary_

    Returns:
        int: _description_
    """

    cell = BoxToCheck(
        sheet_name=SheetName.SHEET_5.value,
        cell_address="B3",
    )

    cell_value: Optional[str] = cell.get_value()

    if not cell_value:
        return "V1"

    return cell_value.strip().split(":")[0].strip()


def count_not_none_in_nc_j() -> int:
    """_summary_

    Returns:
        int: _description_
    """

    counter_nc_not_none: int = 0

    for row in range(START_LINE_REPORT_AUDIT, NB_LINE_REPORT_AUDIT):
        cell_adress = f"J{row}"

        cell = BoxToCheck(
            sheet_name=SheetName.SHEET_5.value,
            cell_address=cell_adress,
        )

        try:

            cell_value: Optional[str] = cell.get_value()

            if cell_value:
                counter_nc_not_none += 1

        except Exception as _:
            pass

    return counter_nc_not_none


def extract_ids(input_text: Optional[str]) -> List[str]:
    """_summary_

    Args:
        input_text (_type_): _description_

    Returns:
        _type_: _description_
    """

    if not input_text:
        return []

    french_numbers = {
        "zéro": 0,
        "zero": 0,
        "un": 1,
        "1er": 1,
        "premier": 1,
        "première": 1,
        "deux": 2,
        "second": 2,
        "seconde": 2,
        "trois": 3,
        "quatre": 4,
        "cinq": 5,
        "six": 6,
        "sept": 7,
        "huit": 8,
        "neuf": 9,
        "dix": 10,
        "onze": 11,
        "douze": 12,
        "treize": 13,
        "quatorze": 14,
        "quinze": 15,
        "seize": 16,
        "dix-sept": 17,
        "dix sept": 17,
        "dixsept": 17,
        "dix-huit": 18,
        "dix huit": 18,
        "dixhuit": 18,
        "dix-neuf": 19,
        "dix neuf": 19,
        "dixneuf": 19,
        "vingt": 20,
        "vingt et un": 21,
        "vingt-et-un": 21,
        "vingtetun": 21,
        "vingt-deux": 22,
        "vingt deux": 22,
        "vingtdeux": 22,
        "vingt-trois": 23,
        "vingt trois": 23,
        "vingttrois": 23,
        "vingt-quatre": 24,
        "vingt quatre": 24,
        "vingtquatre": 24,
        "vingt-cinq": 25,
        "vingt cinq": 25,
        "vingtcinq": 25,
        "vingt-six": 26,
        "vingt six": 26,
        "vingtsix": 26,
        "vingt-sept": 27,
        "vingt sept": 27,
        "vingtsept": 27,
        "vingt-huit": 28,
        "vingt huit": 28,
        "vingthuit": 28,
        "vingt-neuf": 29,
        "vingt neuf": 29,
        "vingtneuf": 29,
        "trente": 30,
        "trente et un": 31,
        "trente-et-un": 31,
        "trenteetun": 31,
        "trente-deux": 32,
        "trente deux": 32,
        "trentedeux": 32
    }

    word_pattern = "|".join(french_numbers.keys())

    number_matches = re.findall(r'\b\d+\b', input_text)

    word_matches = re.findall(r'\b(?:{})\b'.format(word_pattern), input_text,
                              re.IGNORECASE)

    ids = [int(num) for num in number_matches if 1 <= int(num) <= 32]

    ids.extend(french_numbers[word.lower()] for word in word_matches
               if word.lower() in french_numbers)

    return sorted(ids)
