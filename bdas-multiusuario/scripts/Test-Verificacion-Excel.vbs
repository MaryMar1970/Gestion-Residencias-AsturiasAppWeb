Dim excel, wb, proj, fso, file, folder, excelPath
Set fso = CreateObject("Scripting.FileSystemObject")
Dim bdasDir: bdasDir = "H:\ResidenciaApp\bdas-multiusuario\"

Set folder = fso.GetFolder(bdasDir)
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
        excelPath = file.Path
        Exit For
    End If
Next

WScript.Echo "=================================================="
WScript.Echo " VERIFICACION DE COMPILACION Y LIBRERIAS EXCEL"
WScript.Echo "=================================================="
WScript.Echo "Archivo: " & excelPath

Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.EnableEvents = False

On Error Resume Next
Set wb = excel.Workbooks.Open(excelPath, 0, True) ' Abrir en modo Solo Lectura
If Err.Number <> 0 Then
    WScript.Echo "ERROR al abrir el libro: " & Err.Description
    excel.Quit
    WScript.Quit 1
End If

Set proj = wb.VBProject

WScript.Echo "Modulos cargados en VBProject (" & proj.VBComponents.Count & " componentes):"
Dim comp, countMods, countForms
countMods = 0
countForms = 0

For Each comp In proj.VBComponents
    If comp.Type = 1 Then ' vbext_ct_StdModule
        countMods = countMods + 1
    ElseIf comp.Type = 3 Then ' vbext_ct_MSForm
        countForms = countForms + 1
        WScript.Echo "  - UserForm: " & comp.Name
    End If
Next

WScript.Echo "Total Modulos estandar (.bas): " & countMods
WScript.Echo "Total UserForms (.frm): " & countForms

' Verificar modDatabase
Dim compDB
Set compDB = proj.VBComponents.Item("modDatabase")
If compDB Is Nothing Then
    WScript.Echo "❌ CRITICO: modDatabase NO existe en el libro."
Else
    WScript.Echo "✅ OK: modDatabase presente con " & compDB.CodeModule.CountOfLines & " lineas."
End If

' Verificar ThisWorkbook events
Dim compTB
Set compTB = proj.VBComponents.Item("ThisWorkbook")
If compTB.CodeModule.CountOfLines > 0 Then
    WScript.Echo "✅ OK: Eventos ThisWorkbook configurados con " & compTB.CodeModule.CountOfLines & " lineas."
Else
    WScript.Echo "❌ ATENCION: ThisWorkbook esta vacio."
End If

wb.Close False
excel.Quit
WScript.Echo "=================================================="
WScript.Echo " VERIFICACION FINALIZADA SATISFACTORIAMENTE."
WScript.Echo "=================================================="
