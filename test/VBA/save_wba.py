from openpyxl import load_workbook

# Load the existing workbook
wb = load_workbook("/Users/remillieux/Documents/Proneo/logiciel/data/Plan et Rapport d'audit certification V32.xlsm", keep_vba=True)

# Your VBA code as a string
vba_code = '''Sub StoreSpecificCheckboxValue()
        Dim shp As Shape
        Dim ws As Worksheet
        Dim checkbox As Object
        
        ' Set the worksheet to the active sheet or specify the sheet name
        Set ws = ThisWorkbook.ActiveSheet
        
        ' Loop through all shapes in the worksheet
        For Each shp In ws.Shapes
            ' Check if the shape is a form control and has the specified name
            If shp.Type = msoFormControl Then
                If shp.FormControlType = xlCheckBox And shp.Name = "Check Box 59" Then
                    ' Set checkbox object
                    Set checkbox = shp.ControlFormat
                    ' Store the name and value of the checkbox to the sheet
                    ws.Cells(1, 1).Value = "Checkbox Name: " & shp.Name
                    ws.Cells(1, 2).Value = "Value: " & IIf(checkbox.Value = xlOn, "Checked", "Unchecked")
                    Exit For
                End If
            End If
        Next shp
    End Sub
'''

# Add the VBA code to the workbook
wb.vba_archive.writestr('StoreSpecificCheckboxValue', vba_code)

wb.save("/Users/remillieux/Documents/Proneo/logiciel/data/Plan et Rapport d'audit certification V32.xlsm")

