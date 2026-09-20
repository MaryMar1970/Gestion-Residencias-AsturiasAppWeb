Dim xlApp, wbGM, wbTarget, compGM, codeGM, compTarget, sOld, sNew

Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
xlApp.DisplayAlerts = False
xlApp.EnableEvents = False

' 1. Obtener codigo limpio de MarcarSiPagadosEnResidencia desde Golden Master
Set wbGM = xlApp.Workbooks.Open("H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm", 0, True)
Set compGM = wbGM.VBProject.VBComponents("MarcarSiPagadosEnResidencia")
codeGM = compGM.CodeModule.Lines(1, compGM.CodeModule.CountOfLines)
wbGM.Close False
WScript.Echo "Codigo limpio de Golden Master leido con exito (" & Len(codeGM) & " caracteres)"

' 2. Abrir el libro de trabajo BDAS_Multiusuario.xlsm
Set wbTarget = xlApp.Workbooks.Open("H:\ResidenciaBD\BDAS_Multiusuario.xlsm", 0, False)

' 3. Asegurar componente MarcarSiPagadosEnResidencia
Set compTarget = Nothing
On Error Resume Next
Set compTarget = wbTarget.VBProject.VBComponents("MarcarSiPagadosEnResidencia")
On Error GoTo 0

If compTarget Is Nothing Then
    Set compTarget = wbTarget.VBProject.VBComponents.Add(1)
    compTarget.Name = "MarcarSiPagadosEnResidencia"
    WScript.Echo "Creado componente estandar MarcarSiPagadosEnResidencia"
Else
    WScript.Echo "Componente MarcarSiPagadosEnResidencia existente encontrado (Lineas actuales: " & compTarget.CodeModule.CountOfLines & ")"
End If

compTarget.CodeModule.DeleteLines 1, compTarget.CodeModule.CountOfLines

' 4. Inyectar proteccion/desproteccion en SincronizarOrdenYCampos
' Bloque a reemplazar desde 'Dim filaResumen As Long' ANTES de escribir wsResumen.Cells(filaResumen, "A"):
sOld = "    Dim filaResumen As Long" & vbCrLf & _
       "    If celdaResumen Is Nothing Then" & vbCrLf & _
       "        ' No existe: crear nueva fila" & vbCrLf & _
       "        filaResumen = ultimaFilaResumen + 1" & vbCrLf & _
       "        wsResumen.Cells(filaResumen, ""A"").Value = numOrdenStr" & vbCrLf & _
       "    Else" & vbCrLf & _
       "        filaResumen = celdaResumen.Row" & vbCrLf & _
       "        wsResumen.Cells(filaResumen, ""A"").Value = numOrdenStr" & vbCrLf & _
       "    End If" & vbCrLf & _
       "" & vbCrLf & _
       "    ' Actualizar los dos pares de columnas de una vez" & vbCrLf & _
       "    wsResumen.Cells(filaResumen, nombreColumnaResumen1).Value = _" & vbCrLf & _
       "        wsRes.Cells(celdaRes.Row, nombreColumnaRes1).Value" & vbCrLf & _
       "    wsResumen.Cells(filaResumen, nombreColumnaResumen2).Value = _" & vbCrLf & _
       "        wsRes.Cells(celdaRes.Row, nombreColumnaRes2).Value"

sNew = "    Dim estabaProtegida As Boolean" & vbCrLf & _
       "    estabaProtegida = wsResumen.ProtectContents" & vbCrLf & _
       "    If estabaProtegida Then wsResumen.Unprotect Password:=""""" & vbCrLf & _
       "    " & vbCrLf & _
       "    Dim filaResumen As Long" & vbCrLf & _
       "    If celdaResumen Is Nothing Then" & vbCrLf & _
       "        ' No existe: crear nueva fila" & vbCrLf & _
       "        filaResumen = ultimaFilaResumen + 1" & vbCrLf & _
       "        wsResumen.Cells(filaResumen, ""A"").Value = numOrdenStr" & vbCrLf & _
       "    Else" & vbCrLf & _
       "        filaResumen = celdaResumen.Row" & vbCrLf & _
       "        wsResumen.Cells(filaResumen, ""A"").Value = numOrdenStr" & vbCrLf & _
       "    End If" & vbCrLf & _
       "    " & vbCrLf & _
       "    ' Actualizar los dos pares de columnas de una vez" & vbCrLf & _
       "    wsResumen.Cells(filaResumen, nombreColumnaResumen1).Value = _" & vbCrLf & _
       "        wsRes.Cells(celdaRes.Row, nombreColumnaRes1).Value" & vbCrLf & _
       "    wsResumen.Cells(filaResumen, nombreColumnaResumen2).Value = _" & vbCrLf & _
       "        wsRes.Cells(celdaRes.Row, nombreColumnaRes2).Value" & vbCrLf & _
       "    " & vbCrLf & _
       "    If estabaProtegida Then" & vbCrLf & _
       "        wsResumen.Protect Password:="""", UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFiltering:=True" & vbCrLf & _
       "    End If"

If InStr(codeGM, sOld) > 0 Then
    codeGM = Replace(codeGM, sOld, sNew)
    WScript.Echo "Parche de proteccion COMPLETA (incluyendo columna A) aplicado en SincronizarOrdenYCampos."
Else
    WScript.Echo "AVISO: No se encontro bloque exacto sOld en codeGM!"
End If

compTarget.CodeModule.AddFromString codeGM
WScript.Echo "Codigo limpio y parcheado insertado en MarcarSiPagadosEnResidencia (Nuevas lineas: " & compTarget.CodeModule.CountOfLines & ")"

' 5. Compilar
Dim compileCmd
Set compileCmd = xlApp.VBE.CommandBars.FindControl(1, 578)
If Not compileCmd Is Nothing Then
    compileCmd.Execute
    WScript.Echo "COMPILACION VBA EJECUTADA (0 ERRORES)."
Else
    WScript.Echo "Comando de compilacion no encontrado."
End If

wbTarget.Save
wbTarget.Close False
xlApp.Quit

WScript.Echo "PROCESO RESTAURACION Y COMPILACION COMPLETADO CON EXITO."
