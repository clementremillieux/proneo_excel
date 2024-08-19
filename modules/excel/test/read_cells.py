"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.params.schemas import AppParams

from config.logger_config import logger

SHEET_NAME = "Rapprt d'audit"

CELLS_ADDRESS = "F4"

excel_handler = ExcelHandler(excel_abs_path=AppParams().excel_abs_path)

cell_value = excel_handler.read_cell_value(sheet_name=SHEET_NAME,
                                           cell_address=CELLS_ADDRESS)

logger.info("%s value (%s) : %s", CELLS_ADDRESS, SHEET_NAME, cell_value)
