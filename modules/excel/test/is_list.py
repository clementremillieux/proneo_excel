"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.params.schemas import AppParams

excel_handler = ExcelHandler()

excel_handler.load_excel(
    excel_abs_path="/Users/remillieux/Downloads/test_pornie_de.xlsm")

is_drop_down: bool = excel_handler.is_drop_down(
    sheet_name="Rapport d'audit",
    cell_address="L7",
    excel_abs_path="/Users/remillieux/Downloads/test_pornie_de.xlsm")

print(is_drop_down)
