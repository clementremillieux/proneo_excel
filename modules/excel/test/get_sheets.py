"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.params.schemas import AppParams

from config.logger_config import logger

excel_handler = ExcelHandler(excel_abs_path=AppParams().excel_abs_path)

sheets = excel_handler.get_all_sheets()

for sheet in sheets:
    logger.info("- %s", sheet)
