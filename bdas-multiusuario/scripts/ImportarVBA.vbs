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
bdasDir = "H:\ResidenciaApp\bdas-multiusuario\"

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

Sub RemoveComp(compName)
    On Error Resume Next
    Dim comp
    Set comp = proj.VBComponents.Item(compName)
    If Not comp Is Nothing Then
        proj.VBComponents.Remove comp
    End If
    On Error GoTo 0
End Sub

On Error Resume Next
WScript.Echo "Eliminando formularios y modulos antiguos..."
RemoveComp "frmLogin"
RemoveComp "frmSelectorResidencia"
RemoveComp "modDatabase"
RemoveComp "modMigracion"

WScript.Echo "Importando frmLogin.frm..."
proj.VBComponents.Import scriptsDir & "frmLogin.frm"

WScript.Echo "Importando frmSelectorResidencia.frm..."
proj.VBComponents.Import scriptsDir & "frmSelectorResidencia.frm"

WScript.Echo "Importando modDatabase.bas..."
proj.VBComponents.Import scriptsDir & "modDatabase.bas"

WScript.Echo "Importando modMigracion.bas..."
proj.VBComponents.Import scriptsDir & "modMigracion.bas"

' Importar módulos modificados de vba-modules
Dim vbaFolder, vbaFile, modName
Dim vbaDir
vbaDir = bdasDir & "vba-modules\"

If fso.FolderExists(vbaDir) Then
    Set vbaFolder = fso.GetFolder(vbaDir)
    Dim modsToImport
    modsToImport = Array("FechasPeticion", "EvitarDuplicidadSolicitudesGyS", "AsignarNumFactura", "FacturacionMesGIJON", "FacturacionMesOVIEDO", "FacturacionMesSOTO", "MarcarSiPagadosEnResidencia", "ModuloCalendarioGijon", "ModuloCalendarioOviedo", "ModuloCalendarioSoto", "BusquedaDNIResidencias", "BusquedaOrdenNombreFactura", "ModuloLOG", "ModListaNegra", "ReevaluacionSolicitudes")
    
    Dim mName
    For Each mName In modsToImport
        If fso.FileExists(vbaDir & mName & ".bas") Then
            WScript.Echo "Re-importando modulo adaptado: " & mName
            RemoveComp mName
            proj.VBComponents.Import vbaDir & mName & ".bas"
        End If
    Next
End If
On Error GoTo 0

WScript.Echo "Actualizando eventos en ThisWorkbook..."
Dim tb, codeMod
Set tb = proj.VBComponents.Item("ThisWorkbook")
Set codeMod = tb.CodeModule
If codeMod.CountOfLines > 0 Then
    codeMod.DeleteLines 1, codeMod.CountOfLines
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(scriptsDir & "ThisWorkbook_Events.bas") Then
    Set f = fso.OpenTextFile(scriptsDir & "ThisWorkbook_Events.bas", 1)
    eventsCode = f.ReadAll()
    f.Close()
    codeMod.AddFromString eventsCode
End If

WScript.Echo "Guardando cambios en el libro Excel..."
wb.Save
WScript.Echo "IMPORTACION COMPLETADA CON EXITO."

wb.Close False
excel.Quit
