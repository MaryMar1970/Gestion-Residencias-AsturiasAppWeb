Sub UpdateFormsInWorkbook(excelPath)
    Dim xl, wb, proj, fso, scriptsDir
    Set fso = CreateObject("Scripting.FileSystemObject")
    scriptsDir = "H:\ResidenciaApp\bdas-multiusuario\scripts\"
    
    WScript.Echo "Actualizando formularios en: " & excelPath
    
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    
    Set wb = xl.Workbooks.Open(excelPath, 0, False)
    Set proj = wb.VBProject
    
    ' Actualizar frmLogin
    Dim compLogin, codeLogin, fLogin
    Set compLogin = proj.VBComponents.Item("frmLogin")
    If compLogin.CodeModule.CountOfLines > 0 Then compLogin.CodeModule.DeleteLines 1, compLogin.CodeModule.CountOfLines
    Set fLogin = fso.OpenTextFile(scriptsDir & "frmLogin.frm", 1)
    codeLogin = fLogin.ReadAll()
    fLogin.Close()
    compLogin.CodeModule.AddFromString codeLogin
    
    ' Actualizar frmSelectorResidencia
    Dim compSel, codeSel, fSel
    Set compSel = proj.VBComponents.Item("frmSelectorResidencia")
    If compSel.CodeModule.CountOfLines > 0 Then compSel.CodeModule.DeleteLines 1, compSel.CodeModule.CountOfLines
    Set fSel = fso.OpenTextFile(scriptsDir & "frmSelectorResidencia.frm", 1)
    codeSel = fSel.ReadAll()
    fSel.Close()
    compSel.CodeModule.AddFromString codeSel
    
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

UpdateFormsInWorkbook masterPath
UpdateFormsInWorkbook deployPath

WScript.Echo "LAYOUTS DE FORMULARIOS ACTUALIZADOS CON EXITO."
