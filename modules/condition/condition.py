"""_summary_"""

import re

from datetime import datetime as dt

from typing import List, Optional

from venv import logger

from modules.cells.schemas import BoxToCheck, CheckBoxToCheck, DateToCheck

from modules.condition.schemas import Condition, ConditionType, CellsConditionReport, CellsConditionState

from modules.excel import excel_handler
from modules.sheet.schemas import SheetName

REF_TO_NB_INDIC = {"V1": 21}


class CellsConditions:
    """_summary_

    Returns:
        _type_: _description_
    """

    def __init__(self, conditions: List[Condition]) -> None:
        self.conditions: List[Condition] = conditions

    def check(self) -> Optional[CellsConditionReport]:
        """_summary_

        Returns:
            _type_: _description_
        """

        try:

            for condition in self.conditions:
                cells_condition_report: CellsConditionReport = condition.check(
                )

                if cells_condition_report.state == CellsConditionState.NOT_OK:
                    if not condition.is_parent_condition:
                        break

                    return None

        except Exception as e:
            logger.warning("Error checking condition %s : %s",
                           condition.condition_type.value, e)

            cells_condition_report = CellsConditionReport(
                condition=condition,
                state=CellsConditionState.NOT_OK,
                report_str="Error interne")

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

            report_str = f"Une des cellule {self.cell_date_stop.cell_address} [{self.cell_date_stop.sheet_name}], \
{self.cell_date_start.cell_address} [{self.cell_date_start.sheet_name}] ou  {self.cell_duration.cell_address} [{self.cell_duration.sheet_name}] \
ne sont pas remplies"

        else:

            try:
                duration_cell_value_int: int = int(duration_cell_value)

                date_start: dt = date_start_cell_value

                date_stop: dt = date_stop_cell_value

                results = (date_stop -
                           date_start).days >= duration_cell_value_int

                if results:
                    report_str = f"La durée indiqué à la cellule {self.cell_duration.cell_address} [{self.cell_duration.sheet_name}]\
     correspond aux dates de la cellule {self.cell_date_stop.cell_address} [{self.cell_date_stop.sheet_name}]\
     et de la cellule {self.cell_date_start.cell_address} [{self.cell_date_start.sheet_name}]"

                else:
                    report_str = f"La durée indiqué à la cellule {self.cell_duration.cell_address} [{self.cell_duration.sheet_name}]\
     ne correspond pas aux dates {self.cell_date_stop.cell_address} [{self.cell_date_stop.sheet_name}]\
     et de la cellule {self.cell_date_start.cell_address} [{self.cell_date_start.sheet_name}]"

            except Exception as e:

                logger.error(e)

                results = False

                report_str = f"La valuer de la cellule {self.cell_duration.cell_address} [{self.cell_duration.sheet_name}] n'est pas un nombre"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionOneCheckBoxAmongList(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cells: List[CheckBoxToCheck]

    def __init__(self, cells: List[CheckBoxToCheck],
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.ONE_BOX_CHECKED_AMONG_LIST,
            is_parent_condition=is_parent_condition,
            cells_list=cells)

        self.cells: List[CheckBoxToCheck] = cells

    def check(self) -> CellsConditionReport:
        cells_value: List[Optional[bool]] = [
            cell.get_value() for cell in self.cells
        ]

        checkbox_name: str = " ou ".join(
            f'{cell.cell_address} [{cell.sheet_name}]' for cell in self.cells)

        if not any(cells_value):
            results: bool = False

            report_str = f"Une des check box {checkbox_name} doit être cochée"

        elif sum(1 for value in cells_value if value) > 1:
            results = False

            report_str = f"Seule une des check box {checkbox_name} doit être cochée"

        else:
            results = True

            report_str = f"Une des check box {checkbox_name} à bien été cochée"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionAtLeastOneCheckBoxAmongList(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cells: List[CheckBoxToCheck]

    def __init__(self, cells: List[CheckBoxToCheck],
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.AT_LEAST_ONE_BOX_CHECKED_AMONG_LIST,
            is_parent_condition=is_parent_condition,
            cells_list=cells)

        self.cells: List[CheckBoxToCheck] = cells

    def check(self) -> CellsConditionReport:
        cells_value: List[Optional[bool]] = [
            cell.get_value() for cell in self.cells
        ]

        checkbox_name: str = " ou ".join(
            f'{cell.cell_address} [{cell.sheet_name}]' for cell in self.cells)

        if not any(cells_value):
            results: bool = False

            report_str = f"Une des checkbox doit être cochée : {checkbox_name}"

        else:
            results = True

            report_str = f"Une des check box à bien été cochée : {checkbox_name}"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


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

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_FILLED,
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
            report_str = f"La checkbox {self.cell.cell_address} [{self.cell.sheet_name}] doit être remplie"

        else:
            report_str = f"La checkbox {self.cell.cell_address} [{self.cell.sheet_name}] est remplie"

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionHasToBeValue(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    cell: BoxToCheck

    value: str

    def __init__(self, cell: BoxToCheck, value: str,
                 is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(condition_type=ConditionType.CELL_HAS_TO_BE_VALUE,
                         is_parent_condition=is_parent_condition,
                         cells_list=[cell])

        self.cell: BoxToCheck = cell

        self.value: str = value

    def check(self) -> CellsConditionReport:
        cell_value: Optional[str] = self.cell.get_value()

        logger.info("Value list : %s", cell_value)

        if not cell_value or cell_value != self.value:
            results: bool = False

        else:
            results = True

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        if state == CellsConditionState.NOT_OK:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] doit être égale à {self.value}"

        else:
            report_str = f"La cellule {self.cell.cell_address} [{self.cell.sheet_name}] est égale à {self.value}"

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


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

        logger.info("ID in cell text : %s", ids_nc)

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

        cells_report = CellsConditionReport(condition=self,
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

        nb_nc_min_sheet: int = count_nc_min_sheet()

        nb_nc_min: int = count_nc_min()

        if cell_value_int != nb_nc_min_sheet:
            results: bool = False

            report_str: str = f"Le nombre de NC mineure définie à la cellule {self.cell.cell_address} ({cell_value_int}) ne correspond pas au nombre de fiche NC créée ({nb_nc_min_sheet})"

        elif cell_value_int != nb_nc_min:
            results = False

            report_str = f"Le nombre de NC mineure définie à la cellule {self.cell.cell_address} ({cell_value_int}) ne correspond pas au nombre de NC définie dans la rapport d'audit ({nb_nc_min})"

        else:
            results = True

            report_str = f"Le nombre de NC mineure définie à la cellule {self.cell.cell_address} ({cell_value_int}) correspond au nombre de fiche NC créée ({nb_nc_min_sheet}) et au nombre de NC définie dans la rapport d'audit ({nb_nc_min})"

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=self,
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

        cells_report = CellsConditionReport(condition=self,
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

        nb_nc_min: int = count_nc_min_sheet()

        nb_nc_maj: int = count_nc_maj()

        if nb_nc_min > 0 or nb_nc_maj > 0:
            results: bool = True

        else:
            results = False

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str="")

        return cells_report


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

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionCheckAllSheetDropDown(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    sheet_name: str

    def __init__(self, sheet_name: str, is_parent_condition: bool) -> None:
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

    def check(self) -> CellsConditionReport:

        report_str: str = ""

        for column in ["L", "M", "N", "O", "P", "Q", "R", "S"]:
            for row in range(5, 186):
                cell_adress = f"{column}{row}"

                if excel_handler.is_drop_down(sheet_name=self.sheet_name,
                                              cell_adress=cell_adress):
                    cell = BoxToCheck(
                        sheet_name=SheetName.SHEET_5.value,
                        cell_address=cell_adress,
                    )

                    if not cell.get_value():
                        report_str += f"Une valeur doit être choisie pour la cellule {cell_adress}\n"

        if report_str == "":
            results: bool = True

        else:
            results = False

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


class ConditionCheckAllSheetDescription(Condition):
    """_summary_

    Args:
        Condition (_type_): _description_
    """

    sheet_name: str

    def __init__(self, sheet_name: str, is_parent_condition: bool) -> None:
        """Initialize the ConditionDateSup with start and stop dates."""

        super().__init__(
            condition_type=ConditionType.CHECK_ALL_SHEET_DESCRIPTION,
            is_parent_condition=is_parent_condition,
            cells_list=[
                BoxToCheck(
                    sheet_name=SheetName.SHEET_5.value,
                    cell_address="L5",
                )
            ])

        self.sheet_name = sheet_name

    def check(self) -> CellsConditionReport:

        report_str: str = ""

        description_cells: List[str] = get_description_cells()

        for cell in description_cells:
            if not excel_handler.read_cell_value(
                    sheet_name=SheetName.SHEET_5.value, cell_address=cell):
                report_str += f"La cellule description {cell} [{SheetName.SHEET_5.value}] ne peut pas être vide\n"

        if report_str == "":
            results: bool = True

        else:
            results = False

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


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

    def check(self) -> CellsConditionReport:

        report_str: str = ""

        description_cells: List[str] = get_references_cells()

        for cell in description_cells:
            if not excel_handler.read_cell_value(
                    sheet_name=SheetName.SHEET_5.value, cell_address=cell):
                report_str += f"La cellule références {cell} [{SheetName.SHEET_5.value}] ne peut pas être vide\n"

        if report_str == "":
            results: bool = True

        else:
            results = False

        state = CellsConditionState.OK if results else CellsConditionState.NOT_OK

        cells_report = CellsConditionReport(condition=self,
                                            state=state,
                                            report_str=report_str)

        return cells_report


START_LINE_REPORT_AUDIT = 5

NB_LINE_REPORT_AUDIT = 188


def get_nb_audit_type() -> int:
    """_summary_

    Returns:
        int: _description_
    """

    cells_to_check: List[str] = ["B4", "C4", "D4", "E4", "F4"]

    nb_audit_type: int = 0

    for cell in cells_to_check:
        if excel_handler.is_hidden(sheet_name=SheetName.SHEET_5.value,
                                   cell_adress=cell):
            nb_audit_type += 1

    return nb_audit_type


def get_description_cells() -> List[str]:
    """_summary_

    Returns:
        List[str]: _description_
    """

    description_cells: List[str] = []

    nb_audit_type = get_nb_audit_type()

    for row in range(START_LINE_REPORT_AUDIT, NB_LINE_REPORT_AUDIT):
        cell = f"K{row}"

        cell_value: Optional[str] = excel_handler.read_cell_value(
            sheet_name=SheetName.SHEET_5.value, cell_address=cell)

        if cell_value and "Description" in cell_value:
            if nb_audit_type >= 1:
                description_cells.append(f"L{row}")

            if nb_audit_type >= 2:
                description_cells.append(f"N{row}")

            if nb_audit_type >= 3:
                description_cells.append(f"P{row}")

            if nb_audit_type >= 3:
                description_cells.append(f"R{row}")

    return description_cells


def get_references_cells() -> List[str]:
    """_summary_

    Returns:
        List[str]: _description_
    """

    references_cells: List[str] = []

    nb_audit_type = get_nb_audit_type()

    for row in range(START_LINE_REPORT_AUDIT, NB_LINE_REPORT_AUDIT):
        cell = f"K{row}"

        cell_value: Optional[str] = excel_handler.read_cell_value(
            sheet_name=SheetName.SHEET_5.value, cell_address=cell)

        if cell_value and "Références" in cell_value:
            if nb_audit_type >= 1:
                references_cells.append(f"L{row}")

            if nb_audit_type >= 2:
                references_cells.append(f"N{row}")

            if nb_audit_type >= 3:
                references_cells.append(f"P{row}")

            if nb_audit_type >= 3:
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

        except Exception:
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

    # Invert the mapping to create a regex pattern
    word_pattern = "|".join(french_numbers.keys())

    # Find all occurrences of numbers (both numeric and word forms) in the text
    number_matches = re.findall(r'\b\d+\b', input_text)
    word_matches = re.findall(r'\b(?:{})\b'.format(word_pattern), input_text,
                              re.IGNORECASE)

    # Convert number strings to integers
    ids = [int(num) for num in number_matches if 1 <= int(num) <= 32]

    # Convert word matches to their corresponding integers
    ids.extend(french_numbers[word.lower()] for word in word_matches
               if word.lower() in french_numbers)

    return sorted(ids)
