"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.excel.schemas import CheckboxParams
from modules.params.schemas import AppParams

from config.logger_config import logger

CELL_ADRESS = "A12"

SHEET_NAME = "OPAC"

checkbox_params = CheckboxParams(
    apple_script_path="modules/excel/apple_script/checkbox.scpt")

excel_handler = ExcelHandler(excel_abs_path=AppParams().excel_abs_path)

checkbox_state = excel_handler.get_checkbox_state(cell_address=CELL_ADRESS,
                                                  sheet_name=SHEET_NAME)

logger.info("%s state  : %s", CELL_ADRESS, checkbox_state)
