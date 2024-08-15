"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.params.schemas import AppParams

excel_handler = ExcelHandler(excel_abs_path=AppParams().excel_abs_path)
