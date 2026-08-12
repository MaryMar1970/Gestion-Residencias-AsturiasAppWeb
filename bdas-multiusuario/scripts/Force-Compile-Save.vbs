Dim xl, wb, fso, bdasDir, folder, file, masterPath, deployPath

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

Sub ReSaveFile(p)
    WScript.Echo "Re-guardando: " & p
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    
    Set wb = xl.Workbooks.Open(p, 0, False)
    wb.Save
    wb.Close False
    xl.Quit
    Set xl = Nothing
End Sub

ReSaveFile masterPath
ReSaveFile deployPath

WScript.Echo "RE-GUARDADO FORZADO COMPLETADO."
