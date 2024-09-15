"""_summary_"""

from modules.excel.excel import ExcelHandler

excel_handler = ExcelHandler()

excel_handler.load_excel(
    excel_abs_path="/Users/remillieux/Desktop/testV32.xlsm")

cell = "L10"

is_hidden: bool = excel_handler.is_line_hidden(sheet_name="Rapport d'audit",
                                               cell_adress=cell)

print(is_hidden)

value: str = excel_handler.read_cell_value(sheet_name="Rapport d'audit",
                                           cell_address=cell)

print(value)

for i in range(90, 110):

    cell_adress = f"S{i}"

    b_cell_adress = f"G{i}"

    print(f"\n{cell_adress}")

    value = excel_handler.read_cell_value(sheet_name="Rapport d'audit",
                                          cell_address=cell_adress)

    b_value = excel_handler.read_value_2(sheet_name="Rapport d'audit",
                                         cell_address=b_cell_adress)

    is_hidden = excel_handler.is_line_hidden(sheet_name="Rapport d'audit",
                                             cell_adress=cell_adress)

    is_drop_down = excel_handler.is_drop_down(sheet_name="Rapport d'audit",
                                              cell_adress=cell_adress)

    print(f"Hiden : {is_hidden}")

    print(f"Value : {value}")

    print(f"Drop down : {is_drop_down}")

    print(f"B Value : {b_value}")
