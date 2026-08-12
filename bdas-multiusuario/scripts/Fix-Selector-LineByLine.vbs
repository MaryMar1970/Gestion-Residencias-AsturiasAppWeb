Sub FixSelectorInWorkbook(p)
    Dim xl, wb, proj, comp, codeMod, i, line, countFixes
    
    WScript.Echo "========================================="
    WScript.Echo "Inspeccionando y reparando: " & p
    
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    
    Set wb = xl.Workbooks.Open(p, 0, False)
    Set proj = wb.VBProject
    Set comp = proj.VBComponents.Item("frmSelectorResidencia")
    Set codeMod = comp.CodeModule
    
    countFixes = 0
    WScript.Echo "CountOfLines en " & comp.Name & ": " & codeMod.CountOfLines
    
    For i = codeMod.CountOfLines To 1 Step -1
        line = codeMod.Lines(i, 1)
        If InStr(line, "lstR") > 0 Then
            WScript.Echo "  Linea " & i & " ANTES: " & line
            If Trim(line) = "Dim lstR As Object" Or Trim(line) = "Set lstR = Me.Controls(""lstResidencias"")" Then
                codeMod.DeleteLines i, 1
                WScript.Echo "  Linea " & i & " ELIMINADA"
            Else
                line = Replace(line, "lstR", "lstResidencias")
                codeMod.ReplaceLine i, line
                WScript.Echo "  Linea " & i & " DESPUES: " & line
            End If
            countFixes = countFixes + 1
        End If
    Next
    
    WScript.Echo "Total reemplazos en " & comp.Name & ": " & countFixes
    
    ' Imprimir el codigo final resultatorio
    WScript.Echo "--- CODIGO FINAL EN " & comp.Name & " ---"
    For i = 1 To codeMod.CountOfLines
        WScript.Echo i & ": " & codeMod.Lines(i, 1)
    Next
    WScript.Echo "------------------------------------------"
    
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

FixSelectorInWorkbook masterPath
FixSelectorInWorkbook deployPath

WScript.Echo "REPARACION LINEA A LINEA FINALIZADA CON EXITO."
