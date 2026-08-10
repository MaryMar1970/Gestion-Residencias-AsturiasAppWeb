' Exporta todos los modulos VBA de BDAS_GIJÓN-SOTO-OVIEDO_v16.8.4.xlsm a vba-modules
Dim fso, shell, excel, wb, comp, exportDir, bdasDir, xlsmPath

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

bdasDir = "H:\ResidenciaApp\bdas-multiusuario"
exportDir = bdasDir & "\vba-modules"

If Not fso.FolderExists(exportDir) Then
    fso.CreateFolder(exportDir)
End If

Dim folder, file
Set folder = fso.GetFolder(bdasDir)
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
        xlsmPath = file.Path
        Exit For
    End If
Next

WScript.Echo "Exportando componentes VBA de: " & xlsmPath
WScript.Echo "Directorio destino: " & exportDir

Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False

Set wb = excel.Workbooks.Open(xlsmPath, 0, True)

Dim ext, compType, count
count = 0

For Each comp In wb.VBProject.VBComponents
    compType = comp.Type
    ext = ""
    Select Case compType
        Case 1 ' vbext_ct_StdModule
            ext = ".bas"
        Case 2 ' vbext_ct_ClassModule
            ext = ".cls"
        Case 3 ' vbext_ct_MSForm
            ext = ".frm"
    End Select
    
    If ext <> "" Then
        comp.Export exportDir & "\" & comp.Name & ext
        count = count + 1
    End If
Next

wb.Close False
excel.Quit

WScript.Echo "EXPORTACION COMPLETADA: " & count & " componentes guardados en vba-modules."
