import xlwings as xw

def add_comment_to_excel(sheet_name, cell_address, comment_text):
    try:
        # Connect to the active Excel application
        app = xw.apps.active
        
        # Get the active workbook
        wb = app.books.active
        
        # Get the specified sheet
        sheet = wb.sheets[sheet_name]
        
        # Get the specified cell
        cell = sheet.range(cell_address)
        
        # Add or update the comment
        cell.note.text = comment_text
        
        print(f"Comment added to cell {cell_address} in sheet '{sheet_name}'")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Example usage
SHEET_NAME = "OPAC"
cell_address = "G17"
comment_text = "Pipi"

add_comment_to_excel(SHEET_NAME, cell_address, comment_text)