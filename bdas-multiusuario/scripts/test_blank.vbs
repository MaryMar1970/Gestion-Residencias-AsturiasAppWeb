Dim excel, wb, proj, comp
Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.WindowState = -4137
excel.DisplayAlerts = False

Set wb = excel.Workbooks.Add()
Set proj = wb.VBProject

On Error Resume Next
Set comp = proj.VBComponents.Import("H:\ResidenciaApp\bdas-multiusuario\scripts\modDatabase.bas")
If Err.Number <> 0 Then
    WScript.Echo "ERROR importing modDatabase: " & Err.Description & " (Code: " & Hex(Err.Number) & ")"
Else
    WScript.Echo "SUCCESS: Imported " & comp.Name
End If

wb.Close False
excel.Quit
