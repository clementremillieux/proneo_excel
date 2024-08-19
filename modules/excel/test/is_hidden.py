"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.params.schemas import AppParams

excel_handler = ExcelHandler(excel_abs_path=AppParams().excel_abs_path)

is_drop_down: bool = excel_handler.is_hidden(sheet_name="Rapport d'audit",
                                             cell_adress="F4")

print(is_drop_down)
