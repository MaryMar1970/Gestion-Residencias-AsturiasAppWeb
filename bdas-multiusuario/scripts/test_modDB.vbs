Dim fso, folder, file, excelPath, excel, wb, proj, comp
Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("H:\ResidenciaApp\bdas-multiusuario\")
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
        excelPath = file.Path
        Exit For
    End If
Next

WScript.Echo "Target Excel: " & excelPath
Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.WindowState = -4137
excel.DisplayAlerts = False
excel.EnableEvents = False

Set wb = excel.Workbooks.Open(excelPath, 0, False)
Set proj = wb.VBProject

' Remove existing modDatabase if present
On Error Resume Next
Set comp = proj.VBComponents.Item("modDatabase")
If Not comp Is Nothing Then proj.VBComponents.Remove comp
On Error GoTo 0

WScript.Echo "Attempting to import modDatabase.bas..."
Set comp = proj.VBComponents.Import("H:\ResidenciaApp\bdas-multiusuario\scripts\modDatabase.bas")
WScript.Echo "SUCCESS: Imported " & comp.Name

wb.Save
wb.Close False
excel.Quit
