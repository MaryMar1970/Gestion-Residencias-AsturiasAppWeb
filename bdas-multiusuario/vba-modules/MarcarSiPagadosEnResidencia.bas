Attribute VB_Name = "MarcarSiPagadosEnResidencia"
'Attribute VB_Name = "MarcarSiPagadosEnResidencia"
'==============================================================================
' M?DULO: MarcarSiPagadosEnResidencia v14.11.3
' VERSI?N: 2.0 OPTIMIZADA (Opci?n C - Conservadora)
' FECHA: 2026-02-11
'
' PROP?SITO: Sincronizaciï¿½n bidireccional del campo PAGADO entre hojas
'            RESUMEN (columna M/L) <-> RESIDENCIA (columna AA/AB)
'
' OPTIMIZACIONES v2.0:
'   ? Reemplazo de bucles For Each por .Find() (10-50x mï¿½s r?pido)
'   ? Validaciï¿½n de existencia de hojas antes de procesar
'   ? 100% compatible con código existente
'==============================================================================

Private sincronizandoPagado As Boolean

'------------------------------------------------------------------------------
' GIJÓN: RESUMEN -> RESIDENCIA
'------------------------------------------------------------------------------
Public Sub MarcarPagadoEnResidencia(filaResumen As Long, numOrden As Variant, Optional valorPagado As String = "SI")
    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim celdaEncontrada As Range, celdaDestino As Range
    Dim ultimaFila As Long, valorResolucion As String
    Dim numOrdenStr As String

    If sincronizandoPagado Then Exit Sub
    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA GIJÓN")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN GIJÓN")
    On Error GoTo 0
    
    If wsRes Is Nothing Then Exit Sub
    sincronizandoPagado = True

    ultimaFila = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then GoTo CleanUp

    On Error Resume Next
    Set celdaEncontrada = wsRes.Range("A2:A" & ultimaFila).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo CleanUp
    
    If celdaEncontrada Is Nothing Then GoTo CleanUp
    Set celdaDestino = wsRes.Cells(celdaEncontrada.Row, "AA")
    
    If EsEstadoPagado(valorPagado) Then
        valorResolucion = UCase$(Trim$(wsRes.Cells(celdaEncontrada.Row, "P").Value))
        If valorResolucion <> "SI" And valorResolucion <> "CONCEDIDA" And valorResolucion <> "REEVALUADA" Then
            MsgBox "No es posible generar una factura con la resolución adoptada (" & _
                   wsRes.Cells(celdaEncontrada.Row, "P").Value & ").", vbExclamation, "Factura no generada"
            If Not wsResumen Is Nothing Then wsResumen.Cells(filaResumen, "M").ClearContents
            GoTo CleanUp
        End If
    End If
    
    If UCase(Trim$(CStr(celdaDestino.Value))) <> UCase(Trim$(valorPagado)) Then
        Application.enableEvents = True
        If valorPagado = "" Then
            celdaDestino.ClearContents
        Else
            celdaDestino.Value = valorPagado
        End If
        ' Actualizar en Access (fuente de verdad)
        If IsNumeric(numOrdenStr) And CLng(numOrdenStr) > 0 Then
            modDatabase.ActualizarOrden CLng(numOrdenStr), "EstadoPago", CStr(valorPagado)
        End If
        Application.enableEvents = False
    End If

CleanUp:
    sincronizandoPagado = False
End Sub

'------------------------------------------------------------------------------
' GIJÓN: RESIDENCIA -> RESUMEN
'------------------------------------------------------------------------------
Public Sub MarcarPagadoEnResumen(numOrden As Variant, Optional valorPagado As String = "SI")
    Dim wsResumen As Worksheet
    Dim celdaEncontrada As Range, celdaDestino As Range
    Dim ultimaFila As Long, numOrdenStr As String

    If sincronizandoPagado Then Exit Sub
    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    On Error Resume Next
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN GIJÓN")
    On Error GoTo 0
    
    If wsResumen Is Nothing Then Exit Sub
    sincronizandoPagado = True

    ultimaFila = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then GoTo CleanUp

    On Error Resume Next
    Set celdaEncontrada = wsResumen.Range("A2:A" & ultimaFila).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo CleanUp
    
    If celdaEncontrada Is Nothing Then GoTo CleanUp
    Set celdaDestino = wsResumen.Cells(celdaEncontrada.Row, "M")
    
    If UCase(Trim$(CStr(celdaDestino.Value))) <> UCase(Trim$(valorPagado)) Then
        If valorPagado = "" Then
            celdaDestino.ClearContents
        Else
            celdaDestino.Value = valorPagado
        End If
    End If

CleanUp:
    sincronizandoPagado = False
End Sub

'------------------------------------------------------------------------------
' GEN?RICO: Sincronizaciï¿½n de pares de columnas
'------------------------------------------------------------------------------
Public Sub SincronizarCamposResidenciaResumen( _
    nombreHojaResidencia As String, _
    nombreColumnaResidencia1 As String, _
    nombreColumnaResidencia2 As String, _
    nombreHojaResumen As String, _
    nombreColumnaResumen1 As String, _
    nombreColumnaResumen2 As String, _
    numOrden As Variant)
    
    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim ultimaFilaRes As Long, ultimaFilaResumen As Long
    Dim celdaRes As Range, celdaResumen As Range
    Dim numOrdenStr As String
    
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(nombreHojaResidencia)
    Set wsResumen = ThisWorkbook.Worksheets(nombreHojaResumen)
    On Error GoTo 0
    
    If wsRes Is Nothing Or wsResumen Is Nothing Then Exit Sub
    
    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    ultimaFilaRes = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFilaRes < 2 Or ultimaFilaResumen < 2 Then Exit Sub

    On Error Resume Next
    Set celdaRes = wsRes.Range("A2:A" & ultimaFilaRes).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    Set celdaResumen = wsResumen.Range("A2:A" & ultimaFilaResumen).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If celdaRes Is Nothing Or celdaResumen Is Nothing Then Exit Sub

    Application.enableEvents = False
    wsResumen.Cells(celdaResumen.Row, nombreColumnaResumen1).Value = _
        wsRes.Cells(celdaRes.Row, nombreColumnaResidencia1).Value
    wsResumen.Cells(celdaResumen.Row, nombreColumnaResumen2).Value = _
        wsRes.Cells(celdaRes.Row, nombreColumnaResidencia2).Value
    Application.enableEvents = True
End Sub

'------------------------------------------------------------------------------
' SINCRONIZACIï¿½N COMPLETA: GIJÓN (OPTIMIZADA v3.0 - Arrays + Dictionary)
'------------------------------------------------------------------------------
Public Sub SincronizarTodos_Gijon(Optional esArranque As Boolean = False)
    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim ultimaFilaRes As Long, ultimaFilaResumen As Long
    Dim i As Long, clave As String
    
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA GIJÓN")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN GIJÓN")
    On Error GoTo 0
    
    If wsRes Is Nothing Or wsResumen Is Nothing Then Exit Sub

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual

    ultimaFilaRes = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFilaRes < 2 Or ultimaFilaResumen < 2 Then GoTo CleanExit_Gijon

    ' === Leer datos de RESIDENCIA en memoria ===
    Dim arrResOrdenes As Variant, arrResPagado As Variant, arrResFactura As Variant
    arrResOrdenes = wsRes.Range("A2:A" & ultimaFilaRes).Value
    arrResPagado = wsRes.Range("AA2:AA" & ultimaFilaRes).Value
    arrResFactura = wsRes.Range("C2:C" & ultimaFilaRes).Value
    
    ' Protecciï¿½n: rango de una sola fila devuelve escalar
    If Not IsArray(arrResOrdenes) Then
        Dim tmpOG(1 To 1, 1 To 1) As Variant, tmpPG(1 To 1, 1 To 1) As Variant, tmpFG(1 To 1, 1 To 1) As Variant
        tmpOG(1, 1) = arrResOrdenes: tmpPG(1, 1) = arrResPagado: tmpFG(1, 1) = arrResFactura
        arrResOrdenes = tmpOG: arrResPagado = tmpPG: arrResFactura = tmpFG
    End If
    
    ' === Construir diccionario de RESIDENCIA indexado por Nï¿½ ORDEN ===
    Dim dictRes As Object
    Set dictRes = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(arrResOrdenes, 1)
        clave = Trim$(CStr(arrResOrdenes(i, 1)))
        If clave <> "" Then
            dictRes(clave) = Array(arrResPagado(i, 1), arrResFactura(i, 1))
        End If
    Next i
    
    ' === Leer datos de RESUMEN en memoria (GIJÓN: M=Pagado, N=Factura) ===
    Dim arrResumenOrdenes As Variant
    Dim arrResumenPagado As Variant, arrResumenFactura As Variant
    arrResumenOrdenes = wsResumen.Range("A2:A" & ultimaFilaResumen).Value
    arrResumenPagado = wsResumen.Range("M2:M" & ultimaFilaResumen).Value
    arrResumenFactura = wsResumen.Range("N2:N" & ultimaFilaResumen).Value
    
    If Not IsArray(arrResumenOrdenes) Then
        Dim tmpROG(1 To 1, 1 To 1) As Variant, tmpRPG(1 To 1, 1 To 1) As Variant, tmpRFG(1 To 1, 1 To 1) As Variant
        tmpROG(1, 1) = arrResumenOrdenes: tmpRPG(1, 1) = arrResumenPagado: tmpRFG(1, 1) = arrResumenFactura
        arrResumenOrdenes = tmpROG: arrResumenPagado = tmpRPG: arrResumenFactura = tmpRFG
    End If
    
    ' === Sincronizar en memoria ===
    Dim modificado As Boolean: modificado = False
    Dim datos As Variant
    Dim valActual As String, valNuevo As String
    For i = 1 To UBound(arrResumenOrdenes, 1)
        clave = Trim$(CStr(arrResumenOrdenes(i, 1)))
        If clave <> "" Then
            If dictRes.Exists(clave) Then
                datos = dictRes(clave)
                valActual = Trim$(CStr(arrResumenPagado(i, 1) & ""))
                valNuevo = Trim$(CStr(datos(0) & ""))
                If valActual <> valNuevo Then
                    arrResumenPagado(i, 1) = datos(0): modificado = True
                End If
                valActual = Trim$(CStr(arrResumenFactura(i, 1) & ""))
                valNuevo = Trim$(CStr(datos(1) & ""))
                If valActual <> valNuevo Then
                    arrResumenFactura(i, 1) = datos(1): modificado = True
                End If
            End If
        End If
    Next i
    
    ' === Escribir solo si hubo cambios ===
    If modificado Then
        wsResumen.Range("M2").Resize(UBound(arrResumenPagado, 1), 1).Value = arrResumenPagado
        wsResumen.Range("N2").Resize(UBound(arrResumenFactura, 1), 1).Value = arrResumenFactura
    End If
    
    Set dictRes = Nothing

CleanExit_Gijon:
    If Not esArranque Then
        Application.calculation = xlCalculationAutomatic
        Application.enableEvents = True
        Application.screenUpdating = True
    End If
End Sub

'------------------------------------------------------------------------------
' SINCRONIZACIï¿½N COMPLETA: SOTO (OPTIMIZADA v3.0 - Arrays + Dictionary)
'------------------------------------------------------------------------------
Public Sub SincronizarTodos_Soto(Optional esArranque As Boolean = False)
    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim ultimaFilaRes As Long, ultimaFilaResumen As Long
    Dim i As Long, clave As String
    
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA SOTO")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN SOTO")
    On Error GoTo 0
    
    If wsRes Is Nothing Or wsResumen Is Nothing Then Exit Sub

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual

    ultimaFilaRes = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFilaRes < 2 Or ultimaFilaResumen < 2 Then GoTo CleanExit_Soto

    ' === Leer datos de RESIDENCIA en memoria ===
    Dim arrResOrdenes As Variant, arrResPagado As Variant, arrResFactura As Variant
    arrResOrdenes = wsRes.Range("A2:A" & ultimaFilaRes).Value
    arrResPagado = wsRes.Range("AA2:AA" & ultimaFilaRes).Value
    arrResFactura = wsRes.Range("C2:C" & ultimaFilaRes).Value
    
    If Not IsArray(arrResOrdenes) Then
        Dim tmpOS(1 To 1, 1 To 1) As Variant, tmpPS(1 To 1, 1 To 1) As Variant, tmpFS(1 To 1, 1 To 1) As Variant
        tmpOS(1, 1) = arrResOrdenes: tmpPS(1, 1) = arrResPagado: tmpFS(1, 1) = arrResFactura
        arrResOrdenes = tmpOS: arrResPagado = tmpPS: arrResFactura = tmpFS
    End If
    
    Dim dictRes As Object
    Set dictRes = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(arrResOrdenes, 1)
        clave = Trim$(CStr(arrResOrdenes(i, 1)))
        If clave <> "" Then
            dictRes(clave) = Array(arrResPagado(i, 1), arrResFactura(i, 1))
        End If
    Next i
    
    ' === Leer datos de RESUMEN en memoria (SOTO: L=Pagado, M=Factura) ===
    Dim arrResumenOrdenes As Variant
    Dim arrResumenPagado As Variant, arrResumenFactura As Variant
    arrResumenOrdenes = wsResumen.Range("A2:A" & ultimaFilaResumen).Value
    arrResumenPagado = wsResumen.Range("L2:L" & ultimaFilaResumen).Value
    arrResumenFactura = wsResumen.Range("M2:M" & ultimaFilaResumen).Value
    
    If Not IsArray(arrResumenOrdenes) Then
        Dim tmpROS(1 To 1, 1 To 1) As Variant, tmpRPS(1 To 1, 1 To 1) As Variant, tmpRFS(1 To 1, 1 To 1) As Variant
        tmpROS(1, 1) = arrResumenOrdenes: tmpRPS(1, 1) = arrResumenPagado: tmpRFS(1, 1) = arrResumenFactura
        arrResumenOrdenes = tmpROS: arrResumenPagado = tmpRPS: arrResumenFactura = tmpRFS
    End If
    
    Dim modificado As Boolean: modificado = False
    Dim datos As Variant
    Dim valActual As String, valNuevo As String
    For i = 1 To UBound(arrResumenOrdenes, 1)
        clave = Trim$(CStr(arrResumenOrdenes(i, 1)))
        If clave <> "" Then
            If dictRes.Exists(clave) Then
                datos = dictRes(clave)
                valActual = Trim$(CStr(arrResumenPagado(i, 1) & ""))
                valNuevo = Trim$(CStr(datos(0) & ""))
                If valActual <> valNuevo Then
                    arrResumenPagado(i, 1) = datos(0): modificado = True
                End If
                valActual = Trim$(CStr(arrResumenFactura(i, 1) & ""))
                valNuevo = Trim$(CStr(datos(1) & ""))
                If valActual <> valNuevo Then
                    arrResumenFactura(i, 1) = datos(1): modificado = True
                End If
            End If
        End If
    Next i
    
    If modificado Then
        wsResumen.Range("L2").Resize(UBound(arrResumenPagado, 1), 1).Value = arrResumenPagado
        wsResumen.Range("M2").Resize(UBound(arrResumenFactura, 1), 1).Value = arrResumenFactura
    End If
    
    Set dictRes = Nothing

CleanExit_Soto:
    If Not esArranque Then
        Application.calculation = xlCalculationAutomatic
        Application.enableEvents = True
        Application.screenUpdating = True
    End If
End Sub

'------------------------------------------------------------------------------
' SINCRONIZACIï¿½N COMPLETA: OVIEDO (OPTIMIZADA v3.0 - Arrays + Dictionary)
'------------------------------------------------------------------------------
Public Sub SincronizarTodos_Oviedo(Optional esArranque As Boolean = False)
    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim ultimaFilaRes As Long, ultimaFilaResumen As Long
    Dim i As Long, clave As String
    
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA OVIEDO")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN OVIEDO")
    On Error GoTo 0
    
    If wsRes Is Nothing Or wsResumen Is Nothing Then Exit Sub

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual

    ultimaFilaRes = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFilaRes < 2 Or ultimaFilaResumen < 2 Then GoTo CleanExit_Oviedo

    ' === Leer datos de RESIDENCIA en memoria (OVIEDO: AB=Pagado) ===
    Dim arrResOrdenes As Variant, arrResPagado As Variant, arrResFactura As Variant
    arrResOrdenes = wsRes.Range("A2:A" & ultimaFilaRes).Value
    arrResPagado = wsRes.Range("AB2:AB" & ultimaFilaRes).Value
    arrResFactura = wsRes.Range("C2:C" & ultimaFilaRes).Value
    
    If Not IsArray(arrResOrdenes) Then
        Dim tmpOO(1 To 1, 1 To 1) As Variant, tmpPO(1 To 1, 1 To 1) As Variant, tmpFO(1 To 1, 1 To 1) As Variant
        tmpOO(1, 1) = arrResOrdenes: tmpPO(1, 1) = arrResPagado: tmpFO(1, 1) = arrResFactura
        arrResOrdenes = tmpOO: arrResPagado = tmpPO: arrResFactura = tmpFO
    End If
    
    Dim dictRes As Object
    Set dictRes = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(arrResOrdenes, 1)
        clave = Trim$(CStr(arrResOrdenes(i, 1)))
        If clave <> "" Then
            dictRes(clave) = Array(arrResPagado(i, 1), arrResFactura(i, 1))
        End If
    Next i
    
    ' === Leer datos de RESUMEN en memoria (OVIEDO: M=Pagado, N=Factura) ===
    Dim arrResumenOrdenes As Variant
    Dim arrResumenPagado As Variant, arrResumenFactura As Variant
    arrResumenOrdenes = wsResumen.Range("A2:A" & ultimaFilaResumen).Value
    arrResumenPagado = wsResumen.Range("M2:M" & ultimaFilaResumen).Value
    arrResumenFactura = wsResumen.Range("N2:N" & ultimaFilaResumen).Value
    
    If Not IsArray(arrResumenOrdenes) Then
        Dim tmpROO(1 To 1, 1 To 1) As Variant, tmpRPO(1 To 1, 1 To 1) As Variant, tmpRFO(1 To 1, 1 To 1) As Variant
        tmpROO(1, 1) = arrResumenOrdenes: tmpRPO(1, 1) = arrResumenPagado: tmpRFO(1, 1) = arrResumenFactura
        arrResumenOrdenes = tmpROO: arrResumenPagado = tmpRPO: arrResumenFactura = tmpRFO
    End If
    
    Dim modificado As Boolean: modificado = False
    Dim datos As Variant
    Dim valActual As String, valNuevo As String
    For i = 1 To UBound(arrResumenOrdenes, 1)
        clave = Trim$(CStr(arrResumenOrdenes(i, 1)))
        If clave <> "" Then
            If dictRes.Exists(clave) Then
                datos = dictRes(clave)
                valActual = Trim$(CStr(arrResumenPagado(i, 1) & ""))
                valNuevo = Trim$(CStr(datos(0) & ""))
                If valActual <> valNuevo Then
                    arrResumenPagado(i, 1) = datos(0): modificado = True
                End If
                valActual = Trim$(CStr(arrResumenFactura(i, 1) & ""))
                valNuevo = Trim$(CStr(datos(1) & ""))
                If valActual <> valNuevo Then
                    arrResumenFactura(i, 1) = datos(1): modificado = True
                End If
            End If
        End If
    Next i
    
    If modificado Then
        wsResumen.Range("M2").Resize(UBound(arrResumenPagado, 1), 1).Value = arrResumenPagado
        wsResumen.Range("N2").Resize(UBound(arrResumenFactura, 1), 1).Value = arrResumenFactura
    End If
    
    Set dictRes = Nothing

CleanExit_Oviedo:
    If Not esArranque Then
        Application.calculation = xlCalculationAutomatic
        Application.enableEvents = True
        Application.screenUpdating = True
    End If
End Sub

'------------------------------------------------------------------------------
' FUNCIÓN DE ESTADO
'------------------------------------------------------------------------------
Public Function EstaSincronizandoPagado() As Boolean
    EstaSincronizandoPagado = sincronizandoPagado
End Function

'------------------------------------------------------------------------------
' SOTO: RESUMEN -> RESIDENCIA
'------------------------------------------------------------------------------
Public Sub MarcarPagadoEnResidenciaSoto(filaResumen As Long, numOrden As Variant, Optional valorPagado As String = "SI")
    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim celdaEncontrada As Range, celdaDestino As Range
    Dim ultimaFila As Long, valorResolucion As String
    Dim numOrdenStr As String

    If sincronizandoPagado Then Exit Sub
    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA SOTO")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN SOTO")
    On Error GoTo 0
    
    If wsRes Is Nothing Then Exit Sub
    sincronizandoPagado = True

    ultimaFila = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then GoTo Cleanup_Soto

    On Error Resume Next
    Set celdaEncontrada = wsRes.Range("A2:A" & ultimaFila).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo Cleanup_Soto
    
    If celdaEncontrada Is Nothing Then GoTo Cleanup_Soto
    Set celdaDestino = wsRes.Cells(celdaEncontrada.Row, "AA")
    
    If EsEstadoPagado(valorPagado) Then
        valorResolucion = UCase$(Trim$(wsRes.Cells(celdaEncontrada.Row, "P").Value))
        If valorResolucion <> "SI" And valorResolucion <> "CONCEDIDA" And valorResolucion <> "REEVALUADA" Then
            MsgBox "No es posible generar una factura con la resolución adoptada (" & _
                   wsRes.Cells(celdaEncontrada.Row, "P").Value & ").", vbExclamation, "Factura no generada"
            If Not wsResumen Is Nothing Then wsResumen.Cells(filaResumen, "L").ClearContents
            GoTo Cleanup_Soto
        End If
    End If
    
    If UCase(Trim$(CStr(celdaDestino.Value))) <> UCase(Trim$(valorPagado)) Then
        Application.enableEvents = True
        If valorPagado = "" Then
            celdaDestino.ClearContents
        Else
            celdaDestino.Value = valorPagado
        End If
        Application.enableEvents = False
    End If

Cleanup_Soto:
    sincronizandoPagado = False
End Sub

'------------------------------------------------------------------------------
' SOTO: RESIDENCIA -> RESUMEN
'------------------------------------------------------------------------------
Public Sub MarcarPagadoEnResumenSoto(numOrden As Variant, Optional valorPagado As String = "SI")
    Dim wsResumen As Worksheet
    Dim celdaEncontrada As Range, celdaDestino As Range
    Dim ultimaFila As Long, numOrdenStr As String

    If sincronizandoPagado Then Exit Sub
    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    On Error Resume Next
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN SOTO")
    On Error GoTo 0
    
    If wsResumen Is Nothing Then Exit Sub
    sincronizandoPagado = True

    ultimaFila = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then GoTo Cleanup_ResumenSoto

    On Error Resume Next
    Set celdaEncontrada = wsResumen.Range("A2:A" & ultimaFila).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo Cleanup_ResumenSoto
    
    If celdaEncontrada Is Nothing Then GoTo Cleanup_ResumenSoto
    Set celdaDestino = wsResumen.Cells(celdaEncontrada.Row, "L")
    
    If UCase(Trim$(CStr(celdaDestino.Value))) <> UCase(Trim$(valorPagado)) Then
        If valorPagado = "" Then
            celdaDestino.ClearContents
        Else
            celdaDestino.Value = valorPagado
        End If
    End If

Cleanup_ResumenSoto:
    sincronizandoPagado = False
End Sub

'------------------------------------------------------------------------------
' OVIEDO: RESUMEN -> RESIDENCIA
'------------------------------------------------------------------------------
Public Sub MarcarPagadoEnResidenciaOviedo(filaResumen As Long, numOrden As Variant, Optional valorPagado As String = "SI")
    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim celdaEncontrada As Range, celdaDestino As Range
    Dim ultimaFila As Long, valorResolucion As String
    Dim numOrdenStr As String

    If sincronizandoPagado Then Exit Sub
    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA OVIEDO")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN OVIEDO")
    On Error GoTo 0
    
    If wsRes Is Nothing Then Exit Sub
    sincronizandoPagado = True

    ultimaFila = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then GoTo Cleanup_Oviedo

    On Error Resume Next
    Set celdaEncontrada = wsRes.Range("A2:A" & ultimaFila).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo Cleanup_Oviedo
    
    If celdaEncontrada Is Nothing Then GoTo Cleanup_Oviedo
    Set celdaDestino = wsRes.Cells(celdaEncontrada.Row, "AB")
    
    If EsEstadoPagado(valorPagado) Then
        valorResolucion = UCase$(Trim$(wsRes.Cells(celdaEncontrada.Row, "P").Value))
        If valorResolucion <> "SI" And valorResolucion <> "CONCEDIDA" And valorResolucion <> "REEVALUADA" Then
            MsgBox "No es posible generar una factura con la resolución adoptada (" & _
                   wsRes.Cells(celdaEncontrada.Row, "P").Value & ").", vbExclamation, "Factura no generada"
            If Not wsResumen Is Nothing Then wsResumen.Cells(filaResumen, "M").ClearContents
            GoTo Cleanup_Oviedo
        End If
    End If
    
    If UCase(Trim$(CStr(celdaDestino.Value))) <> UCase(Trim$(valorPagado)) Then
        Application.enableEvents = True
        If valorPagado = "" Then
            celdaDestino.ClearContents
        Else
            celdaDestino.Value = valorPagado
        End If
        Application.enableEvents = False
    End If

Cleanup_Oviedo:
    sincronizandoPagado = False
End Sub

'------------------------------------------------------------------------------
' OVIEDO: RESIDENCIA -> RESUMEN
'------------------------------------------------------------------------------
Public Sub MarcarPagadoEnResumenOviedo(numOrden As Variant, Optional valorPagado As String = "SI")
    Dim wsResumen As Worksheet
    Dim celdaEncontrada As Range, celdaDestino As Range
    Dim ultimaFila As Long, numOrdenStr As String

    If sincronizandoPagado Then Exit Sub
    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    On Error Resume Next
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN OVIEDO")
    On Error GoTo 0
    
    If wsResumen Is Nothing Then Exit Sub
    sincronizandoPagado = True

    ultimaFila = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then GoTo Cleanup_ResumenOviedo

    On Error Resume Next
    Set celdaEncontrada = wsResumen.Range("A2:A" & ultimaFila).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo Cleanup_ResumenOviedo
    
    If celdaEncontrada Is Nothing Then GoTo Cleanup_ResumenOviedo
    Set celdaDestino = wsResumen.Cells(celdaEncontrada.Row, "M")
    
    If UCase(Trim$(CStr(celdaDestino.Value))) <> UCase(Trim$(valorPagado)) Then
        If valorPagado = "" Then
            celdaDestino.ClearContents
        Else
            celdaDestino.Value = valorPagado
        End If
    End If

Cleanup_ResumenOviedo:
    sincronizandoPagado = False
End Sub

'------------------------------------------------------------------------------
' PROCEDIMIENTO: SincronizarNumeroOrden
'
' DESCRIPCI?N: Sincroniza el Nï¿½ ORDEN de RESIDENCIA -> RESUMEN
'              Crea la fila si no existe
'------------------------------------------------------------------------------
Public Sub SincronizarNumeroOrden(nombreHojaResidencia As String, nombreHojaResumen As String, numOrden As Variant)
    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim celdaRes As Range, celdaResumen As Range
    Dim ultimaFilaRes As Long, ultimaFilaResumen As Long
    Dim numOrdenStr As String
    
    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub
    
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(nombreHojaResidencia)
    Set wsResumen = ThisWorkbook.Worksheets(nombreHojaResumen)
    On Error GoTo 0
    
    If wsRes Is Nothing Or wsResumen Is Nothing Then Exit Sub
    
    ultimaFilaRes = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFilaRes < 2 Then Exit Sub
    
    ' Buscar el Nï¿½ ORDEN en RESIDENCIA
    On Error Resume Next
    Set celdaRes = wsRes.Range("A2:A" & ultimaFilaRes).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    
    If celdaRes Is Nothing Then Exit Sub
    
    ' Buscar si ya existe en RESUMEN
    On Error Resume Next
    Set celdaResumen = wsResumen.Range("A2:A" & ultimaFilaResumen).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    
    Application.enableEvents = False
    
    ' Si existe, solo actualizar
    If Not celdaResumen Is Nothing Then
        wsResumen.Cells(celdaResumen.Row, "A").Value = numOrdenStr
    Else
        ' Si no existe, crear nueva fila
        Dim nuevaFila As Long
        nuevaFila = ultimaFilaResumen + 1
        wsResumen.Cells(nuevaFila, "A").Value = numOrdenStr
    End If
    
    Application.enableEvents = True
End Sub
'------------------------------------------------------------------------------
' OPTIMIZACIÓN #4: Sincronizaciï¿½n unificada - 1 Find por hoja en lugar de 3+
' Sustituye a SincronizarNumeroOrden + SincronizarCamposResidenciaResumen
' cuando se edita columna A (Nï¿½ ORDEN)
'------------------------------------------------------------------------------
Public Sub SincronizarOrdenYCampos( _
    nombreHojaResidencia As String, _
    nombreColumnaRes1 As String, _
    nombreColumnaRes2 As String, _
    nombreHojaResumen As String, _
    nombreColumnaResumen1 As String, _
    nombreColumnaResumen2 As String, _
    numOrden As Variant)

    Dim wsRes As Worksheet, wsResumen As Worksheet
    Dim ultimaFilaRes As Long, ultimaFilaResumen As Long
    Dim celdaRes As Range, celdaResumen As Range
    Dim numOrdenStr As String

    numOrdenStr = Trim$(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(nombreHojaResidencia)
    Set wsResumen = ThisWorkbook.Worksheets(nombreHojaResumen)
    On Error GoTo 0

    If wsRes Is Nothing Or wsResumen Is Nothing Then Exit Sub

    ultimaFilaRes = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row

    If ultimaFilaRes < 2 Then Exit Sub

    ' Una sola bï¿½squeda en RESIDENCIA
    On Error Resume Next
    Set celdaRes = wsRes.Range("A2:A" & ultimaFilaRes).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If celdaRes Is Nothing Then Exit Sub

    Application.enableEvents = False

    ' Buscar en RESUMEN: si existe actualizar, si no crear nueva fila
    If ultimaFilaResumen >= 2 Then
        On Error Resume Next
        Set celdaResumen = wsResumen.Range("A2:A" & ultimaFilaResumen).Find( _
            What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        On Error GoTo 0
    End If

    Dim estabaProtegida As Boolean
    estabaProtegida = wsResumen.ProtectContents
    
    Dim pwdHojas As String
    pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()
    
    If estabaProtegida Then
        On Error Resume Next
        If Len(pwdHojas) > 0 Then wsResumen.Unprotect password:=pwdHojas
        If wsResumen.ProtectContents Then wsResumen.Unprotect password:=""
        If wsResumen.ProtectContents Then wsResumen.Unprotect
        On Error GoTo 0
    End If
    
    Dim filaResumen As Long
    If celdaResumen Is Nothing Then
        ' No existe: crear nueva fila
        filaResumen = ultimaFilaResumen + 1
        wsResumen.Cells(filaResumen, "A").Value = numOrdenStr
    Else
        filaResumen = celdaResumen.Row
        wsResumen.Cells(filaResumen, "A").Value = numOrdenStr
    End If
    
    ' Actualizar los dos pares de columnas de una vez
    wsResumen.Cells(filaResumen, nombreColumnaResumen1).Value = _
        wsRes.Cells(celdaRes.Row, nombreColumnaRes1).Value
    wsResumen.Cells(filaResumen, nombreColumnaResumen2).Value = _
        wsRes.Cells(celdaRes.Row, nombreColumnaRes2).Value
    
    If estabaProtegida Then
        On Error Resume Next
        wsResumen.Protect password:=pwdHojas, UserInterfaceOnly:=True, _
            AllowFormattingCells:=True, AllowFiltering:=True
    End If


    Application.enableEvents = True
End Sub

'------------------------------------------------------------------------------
' METODO DE PAGO: Emergente de seleccion
'------------------------------------------------------------------------------
Public Function PedirMetodoPago() As String
    Dim opcion As String
    Do
        opcion = Trim(InputBox( _
            "Seleccione el metodo de pago:" & vbCrLf & vbCrLf & _
            "  1 -> Efectivo" & vbCrLf & _
            "  2 -> Tarjeta" & vbCrLf & _
            "  3 -> Transferencia" & vbCrLf & _
            "  4 -> SMS" & vbCrLf & _
            "  5 -> Bizum" & vbCrLf & _
            "  6 -> Cancelar", _
            "Metodo de Pago", "1"))
        If opcion = "" Or opcion = "6" Then
            PedirMetodoPago = ""
            Exit Function
        End If
        If opcion <> "1" And opcion <> "2" And opcion <> "3" And opcion <> "4" And opcion <> "5" Then
            MsgBox "Opcion no valida. Introduzca 1, 2, 3, 4, 5 o 6.", _
                   vbExclamation, "Opcion incorrecta"
        End If
    Loop While opcion <> "1" And opcion <> "2" And opcion <> "3" And opcion <> "4" And opcion <> "5"
    
    Select Case opcion
        Case "1": PedirMetodoPago = "Efectivo"
        Case "2": PedirMetodoPago = "Tarjeta"
        Case "3": PedirMetodoPago = "Transferencia"
        Case "4": PedirMetodoPago = "SMS"
        Case "5": PedirMetodoPago = "Bizum"
    End Select
End Function
'------------------------------------------------------------------------------
' COMPROBACIï¿½N DE ESTADO PAGADO CENTRALIZADA
'------------------------------------------------------------------------------
Public Function EsEstadoPagado(ByVal valor As String) As Boolean
    Dim valUpper As String
    valUpper = UCase(Trim$(valor))
    EsEstadoPagado = (valUpper = "SI" Or valUpper = "EFECTIVO" Or valUpper = "TARJETA" _
                   Or valUpper = "TRANSFERENCIA" Or valUpper = "SMS" Or valUpper = "BIZUM")
End Function






