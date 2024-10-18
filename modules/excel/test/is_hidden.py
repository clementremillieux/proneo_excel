"""_summary_"""

from modules.excel.excel import ExcelHandler

excel_handler = ExcelHandler()

excel_handler.load_excel(
    excel_abs_path="C:/Users/Remillieux/OneDrive - TowardsChange/tests_sync_excel67.xlsm")

cell = "N5"

is_hidden: bool = excel_handler.is_column_hidden(sheet_name="Rapport d'audit",
                                               cell_address=cell)

print(is_hidden)

value: str = excel_handler.read_cell_value(sheet_name="Rapport d'audit",
                                           cell_address=cell)

print(value)
