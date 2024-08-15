"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.params.schemas import AppParams

from config.logger_config import logger

SHEET_NAME = "OPAC"

CELLS_ADDRESS = "A18"

VALUE = "test_write"

excel_handler = ExcelHandler(excel_abs_path=AppParams().excel_abs_path)

excel_handler.write_cell_value(sheet_name=SHEET_NAME,
                               cell_address=CELLS_ADDRESS,
                               value=VALUE)

cell_value = excel_handler.read_cell_value(sheet_name=SHEET_NAME,
                                           cell_address=CELLS_ADDRESS)

logger.info("Value write on %s (%s) : %s", CELLS_ADDRESS, SHEET_NAME,
            cell_value)
