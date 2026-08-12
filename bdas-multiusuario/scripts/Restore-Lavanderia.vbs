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

Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False

Set wb = xl.Workbooks.Open(excelPath, 0, False)
Set comp = wb.VBProject.VBComponents.Item("ModuloLavanderia_Core")
Set codeMod = comp.CodeModule

codeMod.ReplaceLine 28, "    t = Replace(t, Chr(160), """")"
wb.Save
wb.Close False
xl.Quit
WScript.Echo "LAVANDERIA RESTAURADA."
