"""_summary_"""

from datetime import datetime

from typing import List, Optional

import xlwings as xw

import openpyxl

from config.logger_config import logger

from modules.performances.time_counter import time_execution


class ExcelHandler:
    """_summary_"""

    def __init__(self) -> None:

        self.excel_abs_path: Optional[str] = None

    @time_execution
    def load_excel(self, excel_abs_path: str) -> None:
        """_summary_
        """

        try:

            self.app = xw.apps.active

            if self.app is None:
                self.app = xw.App(visible=True)

            for book in self.app.books:
                if book.fullname == excel_abs_path:
                    self.wb = book

                    break
            else:

                self.wb = self.app.books.open(excel_abs_path)

            self.wb.save()

            self.wb_openpyxl = openpyxl.load_workbook(excel_abs_path,
                                                      keep_vba=True)

            self.excel_abs_path = excel_abs_path

        except Exception as e:
            logger.error("Error openning excel file : %s", e)

    @time_execution
    def load_excel_xlwings(self, excel_abs_path: str) -> None:
        """_summary_
        """

        try:

            self.app = xw.apps.active

            if self.app is None:
                self.app = xw.App(visible=True)

            for book in self.app.books:
                if book.fullname == excel_abs_path:
                    self.wb = book

                    break
            else:

                self.wb = self.app.books.open(excel_abs_path)

            self.wb.save()

            self.excel_abs_path = excel_abs_path

        except Exception as e:
            logger.error("Error openning excel xlwings file : %s", e)

    @time_execution
    def load_openpyxl(self, excel_abs_path: str) -> None:
        """_summary_
        """

        try:
            self.app = xw.apps.active

            if self.app is None:
                self.app = xw.App(visible=True)

            for book in self.app.books:
                if book.fullname == excel_abs_path:
                    self.wb = book

                    break
            else:

                self.wb = self.app.books.open(excel_abs_path)

            self.wb_openpyxl = openpyxl.load_workbook(excel_abs_path,
                                                      keep_vba=True)

            self.excel_abs_path = excel_abs_path

        except Exception as e:
            logger.error("Error openning openpy excel file : %s", e)

    def read_cell_value(self, sheet_name: str, cell_address: str) -> str:
        """_summary_

        Args:
            sheet_name (str): _description_
            cells_address (str): _description_

        Returns:
            str: _description_
        """

        try:

            ws = self.wb_openpyxl[sheet_name]

            cells_value = ws[cell_address].value

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

            ws = self.wb_openpyxl[sheet_name]

            cells_value: datetime = ws[cell_address].value

            return cells_value

        except Exception as e:
            logger.error("Error reading cell date : %s", e)

            return datetime.min

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

            ws = self.wb_openpyxl[sheet_name]

            cells_value: bool = ws[cell_address].value

            return cells_value

        except Exception as e:
            logger.error("Error reading cells : %s", e)

            return False

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

        # if self.excel_abs_path:
        #     self.load_excel_xlwings(excel_abs_path=self.excel_abs_path)

        try:
            self.app = xw.apps.active

            if self.app is None:
                self.app = xw.App(visible=True)

            for book in self.app.books:
                if book.fullname == self.excel_abs_path:
                    self.wb = book

                    break
            else:

                self.wb = self.app.books.open(self.excel_abs_path)

            sheet = self.wb.sheets[sheet_name]

            sheet.activate()

            sheet.range(cell_address).select()

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

    def is_merged(self, sheet_name: str, cell_adress: str) -> bool:
        """_summary_

        Returns:
            bool: _description_
        """

        try:
            ws = self.wb_openpyxl[sheet_name]

            cell = ws[cell_adress]

            cell.column_letter

            return False

        except Exception as e:

            return True

    def is_column_hidden(self, sheet_name: str, cell_adress: str) -> bool:
        """_summary_

        Returns:
            bool: _description_
        """

        try:
            ws = self.wb_openpyxl[sheet_name]

            cell = ws[cell_adress]

            column_letter = cell.column_letter

            if ws.column_dimensions[column_letter].hidden:

                return True

            return False

        except Exception:

            return False

    def is_line_hidden(self, sheet_name: str, cell_adress: str) -> bool:
        """_summary_

        Returns:
            bool: _description_
        """

        try:
            ws = self.wb_openpyxl[sheet_name]

            cell = ws[cell_adress]

            row_number = cell.row

            if ws.row_dimensions[row_number].hidden:

                return True

            return False

        except Exception:

            return False

    def cell_contains_signature(self, sheet_name: str,
                                cell_address: str) -> bool:
        """
        Vérifie si la cellule spécifiée contient une signature (image ou objet).

        Args:
            sheet_name (str): Le nom de la feuille Excel.
            cell_address (str): L'adresse de la cellule (par exemple, 'A20').

        Returns:
            bool: True si la cellule contient une signature, False sinon.
        """
        try:
            sheet = self.wb.sheets[sheet_name]
            cell = sheet.range(cell_address)

            cell_left = cell.left
            cell_top = cell.top
            cell_right = cell_left + cell.width
            cell_bottom = cell_top + cell.height

            for shape in sheet.shapes:

                if shape.type == 'Picture':

                    shape_left = shape.left
                    shape_top = shape.top
                    shape_right = shape_left + shape.width
                    shape_bottom = shape_top + shape.height

                    if not (shape_right < cell_left or shape_left > cell_right
                            or shape_bottom < cell_top
                            or shape_top > cell_bottom):

                        return True

            return False

        except Exception as e:
            logger.error(
                "Erreur lors de la vérification de la signature dans la cellule : %s",
                e)
            return False
