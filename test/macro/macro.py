from appscript import app, mactypes


def run_excel_macro(
    excel_path, macro_name="GetCheckboxValue", checkbox_name="Check Box 1"
):
    excel = app("Microsoft Excel")
    try:
        print(f"Attempting to open Excel file: {excel_path}")
        workbook = excel.open(mactypes.File(excel_path))

        print(f"Running macro: {macro_name}")
        macro_call = f'{macro_name}("{checkbox_name}")'
        result = excel.run_VB_macro(macro_call)

        print(f"Raw result: {result}")

        if result is None:
            print("Macro returned None. Checking active cell value...")
            active_cell_value = excel.active_sheet.active_cell.value.get()
            print(f"Active cell value: {active_cell_value}")

        return f"Macro executed. Result: {result if result is not None else active_cell_value}"
    except Exception as e:
        print(f"Error details: {type(e).__name__}: {str(e)}")
        return f"Error: {str(e)}"
    finally:
        excel.quit()


# Usage
excel_file_path = "/Users/remillieux/Documents/Proneo/logiciel/data/test.xlsm"
macro_result = run_excel_macro(excel_file_path)
print(macro_result)
