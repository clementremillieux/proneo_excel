"""_summary_"""

from modules.excel.excel import ExcelHandler

excel_handler = ExcelHandler()

excel_handler.load_excel(
    excel_abs_path="/Users/remillieux/Desktop/testV32.xlsm")

cell = "E72"

is_signature: bool = excel_handler.cell_contains_signature(sheet_name="OPAC",
                                                           cell_address=cell)

print(is_signature)

cell = "E75"

is_signature: bool = excel_handler.cell_contains_signature(sheet_name="OPAC",
                                                           cell_address=cell)

print(is_signature)
