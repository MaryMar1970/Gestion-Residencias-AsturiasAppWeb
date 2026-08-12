Sub RecompilePCode(p)
    Dim xl, wb, proj, comp, codeMod, lineCount
    
    WScript.Echo "Forzando recompilacion completa de P-Code en: " & p
    
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    
    Set wb = xl.Workbooks.Open(p, 0, False)
    Set proj = wb.VBProject
    
    For Each comp In proj.VBComponents
        Set codeMod = comp.CodeModule
        If codeMod.CountOfLines > 0 Then
            ' Modificar un comentario en la primera linea para marcar sucio el modulo
            lineCount = codeMod.CountOfLines
            codeMod.InsertLines 1, "' BDAS Recompiled " & Now()
            codeMod.DeleteLines 1, 1
        End If
    Next
    
    ' Guardar e imponer recompilacion al abrir
    wb.Save
    wb.Close False
    xl.Quit
    Set xl = Nothing
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

RecompilePCode masterPath
RecompilePCode deployPath

WScript.Echo "RECOMPILACION COMPLETA FINALIZADA CON EXITO."
