import xlwings as xw

# Path to your .xlsm file
file_path = "/Users/remillieux/Documents/Proneo/logiciel/data/Plan et Rapport d'audit certification V32.xlsm"

# VBA code you want to add
vba_code = """
Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
"""

# Open the workbook
wb = xw.Book(file_path)

# Add a new module
module = wb.api.VBProject.VBComponents.Add(1)  # 1 corresponds to a standard module

# Insert the VBA code into the module
module.CodeModule.AddFromString(vba_code)

# Save and close the workbook
wb.save()
wb.close()
