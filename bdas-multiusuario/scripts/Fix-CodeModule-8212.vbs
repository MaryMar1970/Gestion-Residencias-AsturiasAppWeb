Dim xl, wb, proj, comp, codeMod, i, line, fixedCount, fso, bdasDir, folder, file, excelPath

Set fso = CreateObject("Scripting.FileSystemObject")
bdasDir = "H:\ResidenciaApp\bdas-multiusuario\"

Set folder = fso.GetFolder(bdasDir)
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
        excelPath = file.Path
        Exit For
    End If
Next

WScript.Echo "Abriendo libro: " & excelPath

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

Set wb = xl.Workbooks.Open(excelPath, 0, False)
Set proj = wb.VBProject

fixedCount = 0

For Each comp In proj.VBComponents
    Set codeMod = comp.CodeModule
    If codeMod.CountOfLines > 0 Then
        For i = 1 To codeMod.CountOfLines
            line = codeMod.Lines(i, 1)
            If InStr(line, "8212") > 0 Or InStr(line, "Chr(160)") > 0 Then
                WScript.Echo "Reparando en [" & comp.Name & "] Linea " & i & ": " & line
                line = Replace(line, "Chr(160) & Chr(8212) & Chr(160)", """ - """)
                line = Replace(line, "ChrW(8212)", """ - """)
                line = Replace(line, "Chr(8212)", """ - """)
                line = Replace(line, "Chr(160)", """ """)
                codeMod.ReplaceLine i, line
                fixedCount = fixedCount + 1
            End If
        Next
    End If
Next

WScript.Echo "Total lineas reparadas: " & fixedCount
wb.Save
wb.Close False
xl.Quit
WScript.Echo "REPARACION FINALIZADA."
