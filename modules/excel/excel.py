"""_summary_"""

from datetime import datetime

import subprocess

from typing import List, Optional

import xlwings as xw

import openpyxl

from modules.excel.schemas import CheckboxParams

from config.logger_config import logger


class ExcelHandler:
    """_summary_"""

    def __init__(self) -> None:

        self.excel_abs_path: Optional[str] = None

    def load_excel(self, excel_abs_path: str) -> None:
        """_summary_
        """

        try:
            self.excel_abs_path: str = excel_abs_path

            self.app = xw.apps.active

            if self.app is None:
                self.app = xw.App(visible=True)

            for book in self.app.books:
                if book.fullname == excel_abs_path:
                    self.wb = book

                    break
            else:

                self.wb = self.app.books.open(excel_abs_path)

            self.app.screen_updating = True

            self.wb.activate()

            self.wb_openpyxl = openpyxl.load_workbook(excel_abs_path,
                                                      keep_vba=True)

            self.excel_abs_path = excel_abs_path

        except Exception as e:
            logger.error("Error openning excel file : %s", e)

    def read_cell_value(self, sheet_name: str, cell_address: str) -> str:
        """_summary_

        Args:
            sheet_name (str): _description_
            cells_address (str): _description_

        Returns:
            str: _description_
        """

        try:

            sheet = self.wb.sheets[sheet_name]

            cells_value: str = sheet.range(cell_address).value

            # logger.info("Excel read value => %s [%s] : %s", cell_address,
            #             sheet_name, cells_value)

            return cells_value

        except Exception as e:
            logger.error("Error reading cells : %s", e)

            return ""

    def read_cell_date_value(self, sheet_name: str,
                             cell_address: str) -> datetime:
        """_summary_

        Args:
            sheet_name (str): _description_
            cells_address (str): _description_

        Returns:
            str: _description_
        """

        try:

            sheet = self.wb.sheets[sheet_name]

            cells_value: datetime = sheet.range(cell_address).value

            # logger.info("Excel read value => %s [%s] : %s", cell_address,
            #             sheet_name, cells_value)

            return cells_value

        except Exception as e:
            logger.error("Error reading cell date : %s", e)

            return datetime.min

    def write_cell_value(self, sheet_name: str, cell_address: str,
                         value: str) -> None:
        """_summary_

        Args:
            sheet_name (str): _description_
            cells_address (str): _description_

        Returns:
            str: _description_
        """

        try:
            sheet = self.wb.sheets[sheet_name]

            current_value = sheet.range(cell_address).value

            if current_value != value:

                sheet.range(cell_address).value = value

                # logger.info("Excel write value => %s [%s] : %s", cell_address,
                #             sheet_name, value)

        except Exception as e:
            logger.error("Error writing cell : %s", e)

    def get_checkbox_state_old(self, checkbox_params: CheckboxParams,
                               checkbox_name: str) -> bool:
        """
        Vérifie l'état d'une case à cocher dans un fichier Excel.

        Args:
        checkbox_params (CheckboxParams): Paramètres de la case à cocher
        checkbox_name (str): Nom de la case à cocher

        Returns:
        bool: True si la case est cochée, False sinon
        """

        try:
            result = subprocess.run([
                'osascript', checkbox_params.apple_script_path, checkbox_name,
                self.excel_abs_path
            ],
                                    capture_output=True,
                                    text=True,
                                    check=True)

            # logger.info("Excel read checkbox value => %s", result)

            return " on" in result.stdout.strip()

        except subprocess.CalledProcessError as e:
            logger.error(
                "Erreur lors de l'exécution du script AppleScript pour la case à cocher %s : %s",
                checkbox_name, e)

            return False

    def get_checkbox_state(self, sheet_name: str, cell_address: str) -> bool:
        """
        Vérifie l'état d'une case à cocher dans un fichier Excel.

        Args:
        checkbox_params (CheckboxParams): Paramètres de la case à cocher
        checkbox_name (str): Nom de la case à cocher

        Returns:
        bool: True si la case est cochée, False sinon
        """

        try:

            sheet = self.wb.sheets[sheet_name]

            cells_value: bool = sheet.range(cell_address).value

            # logger.info("Excel read value => %s [%s] : %s", cell_address,
            #             sheet_name, cells_value)

            return cells_value

        except Exception as e:
            logger.error("Error reading cells : %s", e)

            return False

    def write_commentary(self, sheet_name: str, cell_address: str,
                         value: str) -> None:
        """_summary_

        Args:
            sheet_name (str): _description_
            cells (str): _description_
            value (_type_): _description_
        """

        sheet = self.wb.sheets[sheet_name]

        cell = sheet.range(cell_address)

        if cell.note:
            cell.note.text = value

        else:
            cell.note.add(value)

    def get_all_sheets(self) -> List[str]:
        """_summary_

        Returns:
            List[str]: _description_
        """

        try:
            sheet_names = [sheet.name for sheet in self.wb.sheets]

            return sheet_names

        except Exception as e:

            logger.error("Erreur while getting sheet name's : %s", e)

            return []

    def go_to_sheet_and_cell(self, sheet_name: str, cell_address: str) -> None:
        """
        Active la feuille spécifiée et sélectionne la cellule spécifiée.

        Args:
            sheet_name (str): Le nom de la feuille à activer.
            cell_address (str): L'adresse de la cellule à sélectionner (par exemple, "E34").
        """

        try:
            sheet = self.wb.sheets[sheet_name]

            sheet.activate()

            sheet.range(cell_address).select()

            # logger.info("Navigated to %s in sheet %s", cell_address,
            #             sheet_name)

        except Exception as e:
            logger.error(
                "Erreur lors de la navigation vers %s dans la feuille %s : %s",
                cell_address, sheet_name, e)

    def is_drop_down(self, sheet_name: str, cell_adress: str) -> bool:
        """_summary_

        Returns:
            bool: _description_
        """

        ws = self.wb_openpyxl[sheet_name]

        cell = ws[cell_adress]

        for dv in ws.data_validations.dataValidation:
            if cell.coordinate in dv.cells:

                return True

        return False

    def is_hidden(self, sheet_name: str, cell_adress: str) -> bool:
        """_summary_

        Returns:
            bool: _description_
        """

        ws = self.wb_openpyxl[sheet_name]

        cell = ws[cell_adress]

        column_letter = cell.column_letter

        if ws.column_dimensions[column_letter].hidden:
            return True

        return False
