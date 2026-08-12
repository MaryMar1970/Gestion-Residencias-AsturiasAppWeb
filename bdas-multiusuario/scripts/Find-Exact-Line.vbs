Dim xl, wb, proj, comp, codeMod, i, line, targetPath

targetPath = "H:\ResidenciaBD\BDAS_Multiusuario.xlsm"

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

Set wb = xl.Workbooks.Open(targetPath, 0, False)
Set proj = wb.VBProject
Set comp = proj.VBComponents.Item("frmLogin")
Set codeMod = comp.CodeModule

For i = 1 To codeMod.CountOfLines
    line = codeMod.Lines(i, 1)
    If InStr(line, "8212") > 0 Then
        WScript.Echo "Found 8212 on Line " & i & ": " & line
        codeMod.ReplaceLine i, "    Me.Caption = ""BDAS - Inicio de Sesi"" & Chr(243) & ""n"""
    End If
Next

wb.Save
wb.Close False
xl.Quit
WScript.Echo "DONE."
