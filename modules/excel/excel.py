"""_summary_"""
import os

from datetime import datetime

import time
from typing import List, Optional

import openpyxl

from openpyxl.worksheet.worksheet import Worksheet

import xlwings as xw

from config.logger_config import logger

from modules.performances.time_counter import time_execution


def add_xlwings_conf_sheet(file_path: str):
    """
    Open an Excel file using openpyxl, create a hidden sheet named 'xlwings.conf',
    and add the configuration settings.

    Args:
        file_path (str): Path to the Excel file where the hidden sheet should be added.
    """
    try:
        workbook = openpyxl.load_workbook(file_path, keep_vba=True)

        folder_path = os.path.dirname(file_path)

        xlwings_conf_sheet: Worksheet = workbook.create_sheet(
            title='xlwings.conf')

        xlwings_conf_sheet.sheet_state = 'hidden'

        folder_path = 'C:/Users/Remillieux/OneDrive - TowardsChange'

        config_data = [["ONEDRIVE_CONSUMER_MAC", folder_path],
                       ["ONEDRIVE_COMMERCIAL_MAC", folder_path],
                       ["SHAREPOINT_MAC", folder_path],
                       ["ONEDRIVE_CONSUMER_WIN", folder_path],
                       ["ONEDRIVE_COMMERCIAL_WIN", folder_path],
                       ["SHAREPOINT_WIN", folder_path]]

        for row_index, (key, value) in enumerate(config_data, start=1):
            xlwings_conf_sheet.cell(row=row_index, column=1, value=key)
            xlwings_conf_sheet.cell(row=row_index, column=2, value=value)

        logger.warning(file_path)

        workbook.save(file_path)

        logger.info(
            "xlwings.conf sheet has been added and configured successfully in '%s'.",
            file_path)

        workbook.close()

        time.sleep(5)

    except Exception as e:
        logger.info("An error occurred while modifying the Excel file: %s", e)


class ExcelHandler:
    """_summary_"""

    def __init__(self) -> None:

        self.excel_abs_path: Optional[str] = None

    @time_execution
    def load_excel(self, excel_abs_path: str) -> None:
        """_summary_
        """
        try:
            self.wb_openpyxl.close()

            logger.warning("WB file closed")


        except Exception as e:
            logger.warning(e)

        time.sleep(5)

        logger.warning(excel_abs_path)

        try:

            self.wb_openpyxl = openpyxl.load_workbook(excel_abs_path,
                                                      keep_vba=True, read_only=True)

            self.excel_abs_path = excel_abs_path

        except Exception as e:
            logger.error("Error openning excel file with openpyxl : %s",
                         e,
                         exc_info=True)

        try:

            self.app = xw.apps.active

            if self.app is None:
                self.app = xw.App(visible=True)

            for book in self.app.books:
                book
                if book.fullname == excel_abs_path:
                    self.wb = book

                    break
            else:

                self.wb = self.app.books.open(excel_abs_path)

            self.wb.save()

        except Exception as e:
            logger.error("Error openning excel file with xlwings : %s",
                         e,
                         exc_info=True)

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
            self.wb_openpyxl.close()

            logger.warning("WB file closed")


        except Exception as e:
            logger.warning(e)

        time.sleep(5)

        logger.warning(excel_abs_path)

        self.wb_openpyxl = openpyxl.load_workbook(excel_abs_path,
                                                  keep_vba=True)

        self.excel_abs_path = excel_abs_path

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
        """Get all sheet names using openpyxl.

        Returns:
            List[str]: A list of sheet names.
        """
        try:
            sheet_names = self.wb_openpyxl.sheetnames

            return sheet_names

        except Exception as e:
            logger.error("Error while getting sheet names: %s", e)

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
            # self.app = xw.apps.active

            # if self.app is None:
            #     self.app = xw.App(visible=True)

            # for book in self.app.books:
            #     if book.fullname == self.excel_abs_path:
            #         self.wb = book

            #         break
            # else:

            #     self.wb = self.app.books.open(self.excel_abs_path)

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

        try:

            ws = self.wb_openpyxl[sheet_name]

            cell = ws[cell_adress]

            for dv in ws.data_validations.dataValidation:
                if cell.coordinate in dv.cells:

                    return True
                
        except Exception as e:
            pass

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

        except Exception:

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
