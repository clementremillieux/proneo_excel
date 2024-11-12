"""_summary_"""

from typing import List, Optional

from modules.cells.schemas import BoxToCheck, CheckBoxToCheck

from modules.excel import excel_handler

from modules.condition.condition import CellsConditions, ConditionHasToBeFilled, ConditionOneByChecked, ConditionOneCheckBoxAmongList

from modules.condition.schemas import CellsConditionReport

from modules.excel.schemas import CheckboxParams

from modules.sheet.schemas import SheetName

checkbox_params = CheckboxParams(
    apple_script_path="modules/excel/apple_script/checkbox.scpt")


class Checker:
    """_summary_
    """

    def __init__(self, cells_conditions: List[CellsConditions]) -> None:
        self.nc_names: List[str] = []

        self.cells_conditions: List[CellsConditions] = cells_conditions

        self.nc_cells_conditions: List[CellsConditions] = []

    def check_cells_conditions(self) -> List[CellsConditionReport]:
        """_summary_"""

        if not excel_handler.excel_abs_path:
            return []

        self.handle_nc()

        cells_conditions_reports: List[CellsConditionReport] = []

        for cell_conditions in self.cells_conditions:
            cells_conditions_report: Optional[
                CellsConditionReport
                | List[CellsConditionReport]] = cell_conditions.check()

            if isinstance(cells_conditions_report, CellsConditionReport):
                if cells_conditions_report:
                    cells_conditions_reports.append(cells_conditions_report)

            if isinstance(cells_conditions_report, List):
                if cells_conditions_report:
                    cells_conditions_reports += cells_conditions_report

        for nc_cell_conditions in self.nc_cells_conditions:
            cells_conditions_report: Optional[
                CellsConditionReport
                | List[CellsConditionReport]] = nc_cell_conditions.check()

            if isinstance(cells_conditions_report, CellsConditionReport):
                if cells_conditions_report:
                    cells_conditions_reports.append(cells_conditions_report)

            if isinstance(cells_conditions_report, List):
                if cells_conditions_report:
                    cells_conditions_reports += cells_conditions_report

        return cells_conditions_reports

    def handle_nc(self) -> None:
        """_summary_"""

        self.nc_cells_conditions = []

        all_sheets: List[str] = excel_handler.get_all_sheets()

        nc_sheets: List[str] = [
            sheet for sheet in all_sheets if sheet.startswith("NC")
        ]

        self.create_nc_condition(nc_sheets=nc_sheets)

    def create_nc_condition(self, nc_sheets: List[str]) -> None:
        """_summary_"""

        for sheet_name in nc_sheets:

            self.nc_cells_conditions.append(
                self.create_orga_type_condition(sheet_name=sheet_name))

            self.nc_cells_conditions.append(
                self.create_date_company_condition(sheet_name=sheet_name))

            self.nc_cells_conditions.append(
                self.create_date_2_company_condition(sheet_name=sheet_name))

            self.nc_cells_conditions.append(
                self.create_name_company_condition(sheet_name=sheet_name))

            self.nc_cells_conditions.append(
                self.create_pass_condition(sheet_name=sheet_name))

            self.nc_cells_conditions.append(
                self.create_date_filled_condition(sheet_name=sheet_name))

            self.nc_cells_conditions.append(
                self.create_default_condition(sheet_name=sheet_name))

            self.nc_cells_conditions.append(
                self.create_clue_condition(sheet_name=sheet_name))

    def create_orga_type_condition(self, sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_

        Returns:
            CellsConditions: _description_
        """

        cell_af = CheckBoxToCheck(sheet_name=sheet_name,
                                  checkbox_name="Check Box 3",
                                  cell_address="H20",
                                  checkbox_params=checkbox_params,
                                  alias_name="D20 (Action de formation)")

        cell_bc = CheckBoxToCheck(sheet_name=sheet_name,
                                  checkbox_name="Check Box 3",
                                  cell_address="J20",
                                  checkbox_params=checkbox_params,
                                  alias_name="D20 (Bilan de compétences)")

        cell_vae = CheckBoxToCheck(sheet_name=sheet_name,
                                   checkbox_name="Check Box 3",
                                   cell_address="I20",
                                   checkbox_params=checkbox_params,
                                   alias_name="D20 (VAE)")

        cell_a = CheckBoxToCheck(sheet_name=sheet_name,
                                 checkbox_name="Check Box 3",
                                 cell_address="K20",
                                 checkbox_params=checkbox_params,
                                 alias_name="D20 (Apprentissage)")

        cell_to_check_1 = CheckBoxToCheck(
            sheet_name=SheetName.SHEET_2.value,
            checkbox_name="Check Box 41",
            cell_address="F22",
            checkbox_params=checkbox_params,
            alias_name="D20 (Action de formation)")

        cell_to_check_2 = CheckBoxToCheck(
            sheet_name=SheetName.SHEET_2.value,
            checkbox_name="Check Box 45",
            cell_address="G22",
            checkbox_params=checkbox_params,
            alias_name="D20 (Action de formation)")

        cell_to_check_3 = CheckBoxToCheck(
            sheet_name=SheetName.SHEET_2.value,
            checkbox_name="Check Box 41",
            cell_address="F23",
            checkbox_params=checkbox_params,
            alias_name="D20 (Bilan de compétences)")

        cell_to_check_4 = CheckBoxToCheck(
            sheet_name=SheetName.SHEET_2.value,
            checkbox_name="Check Box 45",
            cell_address="G23",
            checkbox_params=checkbox_params,
            alias_name="D20 (Bilan de compétences)")

        cell_to_check_5 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                          checkbox_name="Check Box 41",
                                          cell_address="F24",
                                          checkbox_params=checkbox_params,
                                          alias_name="D20 (VAE)")

        cell_to_check_6 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                          checkbox_name="Check Box 45",
                                          cell_address="G24",
                                          checkbox_params=checkbox_params,
                                          alias_name="D20 (VAE)")

        cell_to_check_7 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                          checkbox_name="Check Box 41",
                                          cell_address="F25",
                                          checkbox_params=checkbox_params,
                                          alias_name="D20 (Apprentissage)")

        cell_to_check_8 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                          checkbox_name="Check Box 45",
                                          cell_address="G25",
                                          checkbox_params=checkbox_params,
                                          alias_name="D20 (Apprentissage)")

        return CellsConditions(conditions=[
            ConditionOneByChecked(
                cells_to_check=[cell_af, cell_bc, cell_vae, cell_a],
                possible_cells=[
                    cell_to_check_1, cell_to_check_2, cell_to_check_3,
                    cell_to_check_4, cell_to_check_5, cell_to_check_6,
                    cell_to_check_7, cell_to_check_8
                ],
                is_parent_condition=False)
        ])

    def create_date_company_condition(self,
                                      sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_

        Returns:
            CellsConditions: _description_
        """

        return CellsConditions(conditions=[
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=sheet_name,
                cell_address="A28",
            ),
                                   is_parent_condition=True),
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=sheet_name,
                cell_address="C36",
            ),
                                   is_parent_condition=False),
        ])

    def create_date_2_company_condition(self,
                                        sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_

        Returns:
            CellsConditions: _description_
        """

        return CellsConditions(conditions=[
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=sheet_name,
                cell_address="A28",
            ),
                                   is_parent_condition=True),
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=sheet_name,
                cell_address="B37",
            ),
                                   is_parent_condition=False),
        ])

    def create_name_company_condition(self,
                                      sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_

        Returns:
            CellsConditions: _description_
        """

        return CellsConditions(conditions=[
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=sheet_name,
                cell_address="A28",
            ),
                                   is_parent_condition=True),
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=sheet_name,
                cell_address="E37",
            ),
                                   is_parent_condition=False),
        ])

    def create_date_filled_condition(self, sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_
        """

        return CellsConditions(conditions=[
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=SheetName.SHEET_2.value,
                cell_address="G60",
            ),
                                   is_parent_condition=True),
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=sheet_name,
                cell_address="B46",
            ),
                                   is_parent_condition=False),
        ])

    def create_pass_condition(self, sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_
        """

        return CellsConditions(conditions=[
            ConditionHasToBeFilled(cell=BoxToCheck(
                sheet_name=SheetName.SHEET_2.value,
                cell_address="G60",
            ),
                                   is_parent_condition=True),
            ConditionOneCheckBoxAmongList(cells=[
                CheckBoxToCheck(sheet_name=sheet_name,
                                checkbox_name="Check Box 41",
                                cell_address="H40",
                                checkbox_params=checkbox_params,
                                alias_name="E40 (Accepté)"),
                CheckBoxToCheck(sheet_name=sheet_name,
                                checkbox_name="Check Box 41",
                                cell_address="H41",
                                checkbox_params=checkbox_params,
                                alias_name="E41 (Refusé)")
            ],
                                          is_parent_condition=False)
        ])

    def create_default_condition(self, sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_

        Returns:
            CellsConditions: _description_
        """

        cell = BoxToCheck(sheet_name=sheet_name, cell_address="A21")

        default_condition = ConditionHasToBeFilled(
            cell=cell,
            is_parent_condition=False,
            sentence_to_remove=
            "Défaillance identifiée au regard de l'exigence :")

        return CellsConditions(conditions=[default_condition])

    def create_clue_condition(self, sheet_name: str) -> CellsConditions:
        """_summary_

        Args:
            sheet_name (str): _description_

        Returns:
            CellsConditions: _description_
        """

        cell = BoxToCheck(sheet_name=sheet_name, cell_address="A22")

        default_condition = ConditionHasToBeFilled(
            cell=cell,
            is_parent_condition=False,
            sentence_to_remove="Preuve, constat factuel :")

        return CellsConditions(conditions=[default_condition])
