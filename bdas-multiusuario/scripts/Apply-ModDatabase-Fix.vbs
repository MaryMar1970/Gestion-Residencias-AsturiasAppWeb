Dim fso, bdasDir, folder, file, deployPath
Dim masterPath
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

Dim modPath: modPath = "H:\ResidenciaApp\bdas-multiusuario\scripts\modDatabase.bas"
Dim f, code, line
Set f = fso.OpenTextFile(modPath, 1)
code = ""
Do Until f.AtEndOfStream
    line = f.ReadLine()
    If Left(Trim(line), 9) <> "Attribute" Then
        code = code & line & vbCrLf
    End If
Loop
f.Close()

Dim xl
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

Sub ApplyToPath(p)
    WScript.Echo "Actualizando modDatabase en: " & p
    Dim wb, comp
    Set wb = xl.Workbooks.Open(p, 0, False)
    Set comp = wb.VBProject.VBComponents.Item("modDatabase")
    If comp.CodeModule.CountOfLines > 0 Then
        comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
    End If
    comp.CodeModule.InsertLines 1, code
    wb.Save
    wb.Close False
End Sub

ApplyToPath masterPath
ApplyToPath deployPath

xl.Quit
Set xl = Nothing

WScript.Echo "MODDATABASE ACTUALIZADO Y GUARDADO EN AMBOS LIBROS CON EXITO."
