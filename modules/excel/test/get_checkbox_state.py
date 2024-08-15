"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.excel.schemas import CheckboxParams
from modules.params.schemas import AppParams

from config.logger_config import logger

CHECKBOX_NAME = "Check Box 59"

checkbox_params = CheckboxParams(
    apple_script_path="modules/excel/apple_script/checkbox.scpt")

excel_handler = ExcelHandler(excel_abs_path=AppParams().excel_abs_path)

checkbox_state = excel_handler.get_checkbox_state(
    checkbox_name=CHECKBOX_NAME, checkbox_params=checkbox_params)

logger.info("%s state  : %s", CHECKBOX_NAME, checkbox_state)
