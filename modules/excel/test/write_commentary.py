"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.params.schemas import AppParams

from config.logger_config import logger

SHEET_NAME = "OPAC"

CELLS_ADDRESS = "G17"

VALUE = "test_commentary"

excel_handler = ExcelHandler(excel_abs_path=AppParams().excel_abs_path)

excel_handler.write_commentary(sheet_name=SHEET_NAME,
                               cell_address=CELLS_ADDRESS,
                               value=VALUE)

logger.info("Value write on %s (%s) ", CELLS_ADDRESS, SHEET_NAME)
