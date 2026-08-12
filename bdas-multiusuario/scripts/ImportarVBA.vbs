Dim excel, wb, proj, fso, f, eventsCode

' Habilitar AccessVBOM en el registro
Dim sh
Set sh = CreateObject("WScript.Shell")
On Error Resume Next
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM", 1, "REG_DWORD"
sh.RegWrite "HKCU\Software\Microsoft\Office\16.0\Excel\Security\VBAWarnings", 1, "REG_DWORD"
On Error GoTo 0

Set fso = CreateObject("Scripting.FileSystemObject")
Dim scriptsDir, bdasDir, excelPath, file, folder
scriptsDir = "H:\ResidenciaApp\bdas-multiusuario\scripts\"
bdasDir    = "H:\ResidenciaApp\bdas-multiusuario\"

Set folder = fso.GetFolder(bdasDir)
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
        excelPath = file.Path
        Exit For
    End If
Next

If Len(excelPath) = 0 Then
    WScript.Echo "No se encontro el archivo .xlsm en " & bdasDir
    WScript.Quit 1
End If

WScript.Echo "Iniciando Microsoft Excel..."
Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.WindowState = -4137 ' xlMaximized
excel.DisplayAlerts = False
excel.EnableEvents = False
excel.AskToUpdateLinks = False

WScript.Echo "Abriendo libro maestro: " & excelPath
Set wb = excel.Workbooks.Open(excelPath, 0, False)
Set proj = wb.VBProject

Sub UpdateModuleCode(compName, filePath)
    Dim comp, f, code, line
    On Error Resume Next
    Set comp = proj.VBComponents.Item(compName)
    On Error GoTo 0
    
    If comp Is Nothing Then
        WScript.Echo "Importando nuevo componente: " & compName
        proj.VBComponents.Import filePath
    Else
        WScript.Echo "Actualizando codigo de componente: " & compName
        If fso.FileExists(filePath) Then
            Set f = fso.OpenTextFile(filePath, 1)
            code = ""
            Do Until f.AtEndOfStream
                line = f.ReadLine()
                If Left(Trim(line), 9) <> "Attribute" Then
                    code = code & line & vbCrLf
                End If
            Loop
            f.Close()
            
            If comp.CodeModule.CountOfLines > 0 Then
                comp.CodeModule.DeleteLines 1, comp.CodeModule.CountOfLines
            End If
            comp.CodeModule.AddFromString code
        End If
    End If
End Sub

' ==============================================================
' PASO 1 & 2: Actualizar modulos base (.bas)
' ==============================================================
UpdateModuleCode "modDatabase", scriptsDir & "modDatabase.bas"
UpdateModuleCode "modMigracion", scriptsDir & "modMigracion.bas"

' ==============================================================
' PASO 3: Actualizar UserForms nativos
' ==============================================================
UpdateModuleCode "frmLogin", scriptsDir & "frmLogin.frm"
UpdateModuleCode "frmSelectorResidencia", scriptsDir & "frmSelectorResidencia.frm"

' ==============================================================
' PASO 4: Actualizar los 15 modulos adaptados
' ==============================================================
Dim vbaDir
vbaDir = bdasDir & "vba-modules\"

If fso.FolderExists(vbaDir) Then
    Dim modsToImport
    modsToImport = Array("FechasPeticion", "EvitarDuplicidadSolicitudesGyS", "AsignarNumFactura", _
                         "FacturacionMesGIJON", "FacturacionMesOVIEDO", "FacturacionMesSOTO", _
                         "MarcarSiPagadosEnResidencia", "ModuloCalendarioGijon", _
                         "ModuloCalendarioOviedo", "ModuloCalendarioSoto", _
                         "BusquedaDNIResidencias", "BusquedaOrdenNombreFactura", _
                         "ModuloLOG", "ModListaNegra", "ReevaluacionSolicitudes")

    Dim mName
    For Each mName In modsToImport
        If fso.FileExists(vbaDir & mName & ".bas") Then
            UpdateModuleCode mName, vbaDir & mName & ".bas"
        End If
    Next
End If

' ==============================================================
' PASO 5: Actualizar eventos en ThisWorkbook
' ==============================================================
WScript.Echo "Actualizando eventos en ThisWorkbook..."
Dim tb, codeMod
Set tb = proj.VBComponents.Item("ThisWorkbook")
Set codeMod = tb.CodeModule
If codeMod.CountOfLines > 0 Then
    codeMod.DeleteLines 1, codeMod.CountOfLines
End If

If fso.FileExists(scriptsDir & "ThisWorkbook_Events.bas") Then
    Set f = fso.OpenTextFile(scriptsDir & "ThisWorkbook_Events.bas", 1)
    eventsCode = f.ReadAll()
    f.Close()
    codeMod.AddFromString eventsCode
End If

' ==============================================================
' PASO 6: Guardar y cerrar
' ==============================================================
WScript.Echo "Guardando cambios en el libro Excel..."
wb.Save
WScript.Echo "IMPORTACION COMPLETADA CON EXITO."

wb.Close False
excel.Quit
