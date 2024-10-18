"""_summary_"""

from modules.excel.excel import ExcelHandler

from modules.params.schemas import AppParams

excel_handler = ExcelHandler()

excel_handler.load_excel(
    excel_abs_path="C:/Users/Remillieux/OneDrive - TowardsChange/tests_sync_excel67.xlsm")


is_drop_down: bool = excel_handler.is_drop_down(sheet_name="Rapport d'audit",
                                                cell_address="O7", excel_abs_path="C:/Users/Remillieux/OneDrive - TowardsChange/tests_sync_excel67.xlsm")

print(is_drop_down)
