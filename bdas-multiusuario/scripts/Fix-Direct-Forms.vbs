Dim xl, wb, proj, comp, codeMod, targetPath

targetPath = "H:\ResidenciaBD\BDAS_Multiusuario.xlsm"

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

WScript.Echo "Abriendo: " & targetPath
Set wb = xl.Workbooks.Open(targetPath, 0, False)
Set proj = wb.VBProject

' 1. Fijar frmLogin
Set comp = proj.VBComponents.Item("frmLogin")
Set codeMod = comp.CodeModule
WScript.Echo "frmLogin Linea 8 antes: " & codeMod.Lines(8, 1)
codeMod.ReplaceLine 8, "    Me.Caption = ""BDAS - Inicio de Sesi"" & Chr(243) & ""n"""
codeMod.ReplaceLine 14, "    ctl.Caption = ""BDAS - Residencias"""
WScript.Echo "frmLogin Linea 8 despues: " & codeMod.Lines(8, 1)

' 2. Fijar frmSelectorResidencia
Set comp = proj.VBComponents.Item("frmSelectorResidencia")
Set codeMod = comp.CodeModule
WScript.Echo "frmSelectorResidencia Linea 11 antes: " & codeMod.Lines(11, 1)
codeMod.ReplaceLine 11, "    Me.Caption = ""BDAS - Selecci"" & Chr(243) & ""n de Residencia"""
WScript.Echo "frmSelectorResidencia Linea 11 despues: " & codeMod.Lines(11, 1)

wb.Save
wb.Close False
xl.Quit
WScript.Echo "REPARACION DIRECTA DE FORMS COMPLETADA."
