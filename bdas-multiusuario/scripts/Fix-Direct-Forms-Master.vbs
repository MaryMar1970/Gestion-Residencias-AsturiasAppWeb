Dim xl, wb, proj, comp, codeMod, fso, bdasDir, folder, file, excelPath

Set fso = CreateObject("Scripting.FileSystemObject")
bdasDir = "H:\ResidenciaApp\bdas-multiusuario\"

Set folder = fso.GetFolder(bdasDir)
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
        excelPath = file.Path
        Exit For
    End If
Next

WScript.Echo "Abriendo libro maestro: " & excelPath

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

Set wb = xl.Workbooks.Open(excelPath, 0, False)
Set proj = wb.VBProject

' 1. Fijar frmLogin
Set comp = proj.VBComponents.Item("frmLogin")
Set codeMod = comp.CodeModule
codeMod.ReplaceLine 8, "    Me.Caption = ""BDAS - Inicio de Sesi"" & Chr(243) & ""n"""
codeMod.ReplaceLine 14, "    ctl.Caption = ""BDAS - Residencias"""

' 2. Fijar frmSelectorResidencia
Set comp = proj.VBComponents.Item("frmSelectorResidencia")
Set codeMod = comp.CodeModule
codeMod.ReplaceLine 11, "    Me.Caption = ""BDAS - Selecci"" & Chr(243) & ""n de Residencia"""

wb.Save
wb.Close False
xl.Quit
WScript.Echo "REPARACION EN MASTER COMPLETADA."
