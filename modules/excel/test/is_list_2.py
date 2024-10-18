import xlwings as xw

def get_dropdown_values_from_validation(sheet_name: str, cell_address: str):
    """
    Retrieves the drop-down (data validation) values for a cell in Excel.
    
    Args:
        sheet_name (str): The name of the sheet.
        cell_address (str): The address of the cell containing the drop-down list.

    Returns:
        list: A list of values from the drop-down list, or None if no drop-down is found.
    """
    try:
        # Open the workbook
        app = xw.App(visible=False)
        wb = xw.Book('C:/Users/Remillieux/OneDrive - TowardsChange/tests_sync_excel67.xlsm')  # Adjust the path to your Excel file
        sheet = wb.sheets[sheet_name]
        cell = sheet.range(cell_address)

        # Access the cell's data validation settings via the Excel API
        validation = cell.api.Validation

        # Check if the cell has a data validation drop-down list (Type 3 indicates a list)
        if validation.Type == 3:  # 3 indicates a list validation (drop-down)
            # Get the formula for the drop-down list (the source range of the values)
            formula1 = validation.Formula1
            print(f"Drop-down source formula: {formula1}")

            # Extract the range of the drop-down values from the formula
            if formula1.startswith('='):
                # Remove the equals sign and get the range
                dropdown_range = formula1[1:]  # Remove '=' at the start
                dropdown_values = sheet.range(dropdown_range).value  # Read values from that range
                
                # If it's a single column or row, flatten the values into a list
                if isinstance(dropdown_values[0], list):
                    dropdown_values = [item[0] for item in dropdown_values if item[0] is not None]  # Flatten columns
                else:
                    dropdown_values = [item for item in dropdown_values if item is not None]  # Rows
                    
                return dropdown_values
            else:
                print("No range found in the validation formula.")
                return None
        else:
            print(f"Cell {cell_address} does not contain a drop-down list.")
            return None

    except Exception as e:
        print(f"Error retrieving drop-down values: {e}")
        return None

    finally:
        wb.close()
        app.quit()

# Example Usage
sheet_name = "Rapport d'audit"
cell_address = 'N7'  # The cell containing the drop-down
dropdown_values = get_dropdown_values_from_validation(sheet_name, cell_address)
print(f"Drop-down values in {cell_address}: {dropdown_values}")
