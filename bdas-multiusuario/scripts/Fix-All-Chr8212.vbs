Sub FixWorkbook(excelPath)
    Dim xl, wb, proj, comp, codeMod, i, line, countFixes
    
    WScript.Echo "========================================="
    WScript.Echo "Procesando: " & excelPath
    
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    
    Set wb = xl.Workbooks.Open(excelPath, 0, False)
    Set proj = wb.VBProject
    countFixes = 0
    
    For Each comp In proj.VBComponents
        Set codeMod = comp.CodeModule
        If codeMod.CountOfLines > 0 Then
            For i = 1 To codeMod.CountOfLines
                line = codeMod.Lines(i, 1)
                If InStr(line, "8212") > 0 Or InStr(line, "Chr(160)") > 0 Then
                    If comp.Name <> "ModuloLavanderia_Core" Then
                        WScript.Echo "  [" & comp.Name & ":" & i & "] " & line
                        line = Replace(line, "Chr(160) & Chr(8212) & Chr(160)", """ - """)
                        line = Replace(line, "ChrW(8212)", """ - """)
                        line = Replace(line, "Chr(8212)", """ - """)
                        codeMod.ReplaceLine i, line
                        countFixes = countFixes + 1
                    End If
                End If
            Next
        End If
    Next
    
    WScript.Echo "Total lineas corregidas: " & countFixes
    wb.Save
    wb.Close False
    xl.Quit
End Sub

Dim fso, bdasDir, folder, file, masterPath, deployPath
Set fso = CreateObject("Scripting.FileSystemObject")
bdasDir = "H:\ResidenciaApp\bdas-multiusuario\"

Set folder = fso.GetFolder(bdasDir)
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
        masterPath = file.Path
        Exit For
    End If
Next

deployPath = "H:\ResidenciaBD\BDAS_Multiusuario.xlsm"

FixWorkbook masterPath
FixWorkbook deployPath

WScript.Echo "========================================="
WScript.Echo "PROCESO FINALIZADO CON EXITO."
