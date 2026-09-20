Dim xlApp, wbGM, wbTarget, compGM, codeGM, compTarget, compM1, sOld, sNew

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False
xlApp.EnableEvents = False

' 1. Obtener codigo limpio de BorradoNumOrden desde Golden Master
Set wbGM = xlApp.Workbooks.Open("H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm", 0, True)
Set compGM = wbGM.VBProject.VBComponents("BorradoNumOrden")
codeGM = compGM.CodeModule.Lines(1, compGM.CodeModule.CountOfLines)
wbGM.Close False
WScript.Echo "Codigo limpio de Golden Master leido con exito (" & Len(codeGM) & " caracteres)"

' 2. Abrir el libro de trabajo BDAS_Multiusuario.xlsm
Set wbTarget = xlApp.Workbooks.Open("H:\ResidenciaBD\BDAS_Multiusuario.xlsm", 0, False)

' 3. Eliminar cualquier residuo de Modulo1
Dim comp
For Each comp In wbTarget.VBProject.VBComponents
    If InStr(LCase(comp.Name), "dulo1") > 0 Or comp.Name = "Modulo1" Then
        WScript.Echo "Eliminando componente no deseado: " & comp.Name
        wbTarget.VBProject.VBComponents.Remove comp
    End If
Next

' 4. Asegurar componente BorradoNumOrden
Set compTarget = Nothing
On Error Resume Next
Set compTarget = wbTarget.VBProject.VBComponents("BorradoNumOrden")
On Error GoTo 0

If compTarget Is Nothing Then
    Set compTarget = wbTarget.VBProject.VBComponents.Add(1)
    compTarget.Name = "BorradoNumOrden"
    WScript.Echo "Creado componente estandar BorradoNumOrden"
Else
    WScript.Echo "Componente BorradoNumOrden existente encontrado"
End If

compTarget.CodeModule.DeleteLines 1, compTarget.CodeModule.CountOfLines

' 5. Inyectar el fix de desproteccion de hoja RESUMEN sobre el codigo limpio
sOld = "If Not celdaEncontrada Is Nothing Then" & vbCrLf & _
       "        celdaEncontrada.EntireRow.Delete" & vbCrLf & _
       "    End If"

sNew = "If Not celdaEncontrada Is Nothing Then" & vbCrLf & _
       "        Dim estabaProtegida As Boolean" & vbCrLf & _
       "        estabaProtegida = wsResumen.ProtectContents" & vbCrLf & _
       "        If estabaProtegida Then wsResumen.Unprotect Password:=""""" & vbCrLf & _
       "        celdaEncontrada.EntireRow.Delete" & vbCrLf & _
       "        If estabaProtegida Then wsResumen.Protect Password:="""", UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFiltering:=True" & vbCrLf & _
       "    End If"

If InStr(codeGM, sOld) > 0 Then
    codeGM = Replace(codeGM, sOld, sNew)
    WScript.Echo "Fix de desproteccion aplicado sobre BorradoNumOrden"
Else
    WScript.Echo "AVISO: No se encontro el bloque exacto para reemplazar en codeGM"
End If

compTarget.CodeModule.AddFromString codeGM
WScript.Echo "Codigo limpio y parcheado insertado en BorradoNumOrden"

' 6. Compilar
Dim compileCmd
Set compileCmd = xlApp.VBE.CommandBars.FindControl(1, 578)
If Not compileCmd Is Nothing Then
    compileCmd.Execute
    WScript.Echo "COMPILACION VBA COMPLETADA (0 ERRORES)."
Else
    WScript.Echo "Comando de compilacion no encontrado."
End If

wbTarget.Save
wbTarget.Close False
xlApp.Quit
WScript.Echo "PROCESO FINALIZADO CON EXITO."
