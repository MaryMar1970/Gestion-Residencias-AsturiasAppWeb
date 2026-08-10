Attribute VB_Name = "AsignarNumFactura"
Option Explicit

' ==============================================================================
' M�dulo: AsignarNumFactura (OPTIMIZADO v2.0) v14.11.3
' Sistema de asignaci�n inteligente por bloques + sincronizaci�n autom�tica
' ==============================================================================

' ===================================================================================
' Procedimiento principal: AsignarFacturaPorColumna
' OPTIMIZACI�N v2.0: A�adida sincronizaci�n autom�tica con RESUMEN
' ===================================================================================
Public Sub AsignarFacturaPorColumna(ws As Worksheet, _
                                    ByVal Target As Range, _
                                    ByVal columnaCheck As String, _
                                    Optional ByVal columnaFactura As String = "C")

    Dim celda As Range
    Dim nuevoNumero As Long
    Dim numeroOrden As Variant
    Dim respuesta As VbMsgBoxResult
    Dim nombreHojaFactura As String
    Dim hojaFactura As Worksheet
    Dim valorResolucion As String
    Dim mensajeAsignacion As String
    
    ' OPTIMIZACI�N v2.0: Variables para sincronizaci�n
    Dim nombreHojaResidencia As String
    Dim nombreHojaResumen As String
    Dim columnaPagadoResidencia As String
    Dim columnaPagadoResumen As String
    Dim columnaFacturaResumen As String

    ' Determinar las hojas correspondientes
    Select Case ws.Name
        Case "RESIDENCIA GIJ�N"
            nombreHojaFactura = "FACTURA GIJ�N"
            nombreHojaResidencia = "RESIDENCIA GIJ�N"
            nombreHojaResumen = "RESUMEN GIJ�N"
            columnaPagadoResidencia = "AA"
            columnaPagadoResumen = "M"
            columnaFacturaResumen = "N"
            
        Case "RESIDENCIA SOTO"
            nombreHojaFactura = "FACTURA SOTO"
            nombreHojaResidencia = "RESIDENCIA SOTO"
            nombreHojaResumen = "RESUMEN SOTO"
            columnaPagadoResidencia = "AA"
            columnaPagadoResumen = "L"
            columnaFacturaResumen = "M"
            
        Case "RESIDENCIA OVIEDO"
            nombreHojaFactura = "FACTURA OVIEDO"
            nombreHojaResidencia = "RESIDENCIA OVIEDO"
            nombreHojaResumen = "RESUMEN OVIEDO"
            columnaPagadoResidencia = "AB"
            columnaPagadoResumen = "M"
            columnaFacturaResumen = "N"
            
        Case Else
            MsgBox "Esta hoja no est� habilitada para el proceso de facturaci�n autom�tica.", vbExclamation
            Exit Sub
    End Select

    ' Intentar obtener la hoja de factura
    On Error Resume Next
    Set hojaFactura = ws.Parent.Worksheets(nombreHojaFactura)
    On Error GoTo 0

    If hojaFactura Is Nothing Then
        MsgBox "No existe la hoja de factura llamada '" & nombreHojaFactura & "'.", vbCritical
        Exit Sub
    End If

    ' Comprobar si el cambio afecta a la columna de chequeo
    If Not Intersect(Target, ws.Columns(columnaCheck)) Is Nothing Then
        Application.enableEvents = False

        For Each celda In Intersect(Target, ws.Columns(columnaCheck))
            ' Usamos la validacin centralizada de pagos para SMS, Bizum, Transferencia, Tarjeta, Efectivo y SI
            If MarcarSiPagadosEnResidencia.EsEstadoPagado(celda.Value) Then
                ' Validar columna RESOLUCION (P = columna 16)
                valorResolucion = UCase$(Trim(ws.Cells(celda.Row, 16).Value))
                
                If valorResolucion = "SI" Or valorResolucion = "CONCEDIDA" Or valorResolucion = "REEVALUADA" Then
                    ' === NUEVA LÓGICA: Obtener número atómico desde Access ===
                    Dim resCode As String
                    Select Case ws.Name
                        Case "RESIDENCIA GIJN", "RESIDENCIA GIJON": resCode = "GIJON"
                        Case "RESIDENCIA SOTO": resCode = "SOTO"
                        Case "RESIDENCIA OVIEDO": resCode = "OVIEDO"
                        Case Else: resCode = "GIJON"
                    End Select
                    
                    nuevoNumero = modDatabase.ObtenerSiguienteNumFactura(resCode, Year(Now))
                    mensajeAsignacion = "Asignación atómica de factura desde Access para " & resCode
                    
                    ' Capturar el N ORDEN
                    numeroOrden = ws.Cells(celda.Row, 1).Value
                    
                    ' Preguntar al usuario si desea asignar el nmero de factura
                    Dim preguntaAsignar As VbMsgBoxResult
                    preguntaAsignar = MsgBox("Se propone asignar el N de factura " & nuevoNumero & _
                                             " al N ORDEN " & numeroOrden & "." & vbCrLf & _
                                             mensajeAsignacion & vbCrLf & vbCrLf & _
                                             "Desea adjudicar este nmero de factura?", _
                                             vbYesNo + vbQuestion, "Confirmar Adjudicacin de Factura")
                                             
                    If preguntaAsignar = vbYes Then
                        ' Asignar el nmero de factura en Access (fuente de verdad) Y en la celda local
                        If IsNumeric(numeroOrden) And CLng(numeroOrden) > 0 Then
                            modDatabase.ActualizarOrden CLng(numeroOrden), "NumFactura", CStr(nuevoNumero)
                        End If
                        ws.Cells(celda.Row, columnaFactura).Value = nuevoNumero
                        
                        ' ========================================================================
                        ' OPTIMIZACIN v2.0: SINCRONIZAR AUTOMTICAMENTE CON RESUMEN
                        ' Sincroniza C (N FACTURA) y AA/AB (PAGADO) con el RESUMEN
                        ' ========================================================================
                        If Trim(CStr(numeroOrden)) <> "" Then
                            Call MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen( _
                                nombreHojaResidencia, _
                                columnaFactura, _
                                columnaPagadoResidencia, _
                                nombreHojaResumen, _
                                columnaFacturaResumen, _
                                columnaPagadoResumen, _
                                numeroOrden)
                        End If
                        ' ========================================================================
                        
                        ' Preguntar si desea imprimir la factura
                        respuesta = MsgBox("N� de factura " & nuevoNumero & " asignado correctamente." & vbCrLf & vbCrLf & _
                                           "�Desea imprimir la factura?", vbYesNo + vbQuestion, "Imprimir factura")
                        If respuesta = vbYes Then
                            hojaFactura.visible = xlSheetVisible
                            hojaFactura.Activate
                            hojaFactura.Range("D1").Value = numeroOrden
                        End If
                    End If
                Else
                    MsgBox "No es posible generar una factura con la resoluci�n adoptada (" & ws.Cells(celda.Row, 16).Value & ").", vbExclamation, "Factura no generada"
                    celda.Value = ""
                End If
            End If
        Next celda

        Application.enableEvents = True
    End If
End Sub

' ===================================================================================
' Funci�n: ObtenerNumeroFacturaInteligente
' Implementa el sistema de bloques con salto m�ximo de 3 n�meros
' ===================================================================================
Private Function ObtenerNumeroFacturaInteligente(ws As Worksheet, _
                                                 columnaFactura As String, _
                                                 ByRef mensajeInfo As String) As Long
    Dim numeros() As Long
    Dim bloques As Collection
    Dim bloqueSeleccionado As Collection
    Dim i As Long, ultimaFila As Long
    Dim valor As Variant
    Dim contador As Long
    
    ' === 1. ESCANEAR Y EXTRAER TODOS LOS N�MEROS ===
    ultimaFila = ws.Cells(ws.Rows.Count, columnaFactura).End(xlUp).Row
    ReDim numeros(1 To ultimaFila)
    contador = 0
    
    For i = 2 To ultimaFila ' Asumiendo fila 1 = encabezados
        valor = ws.Cells(i, columnaFactura).Value
        If IsNumeric(valor) Then
            If CLng(valor) > 0 Then
                contador = contador + 1
                numeros(contador) = CLng(valor)
            End If
        End If
    Next i
    
    ' Si no hay facturas, asignar 1
    If contador = 0 Then
        mensajeInfo = "Primera factura del sistema."
        ObtenerNumeroFacturaInteligente = 1
        Exit Function
    End If
    
    ' Redimensionar array al tama�o real
    ReDim Preserve numeros(1 To contador)
    
    ' === 2. ORDENAR N�MEROS ===
    Call OrdenarArray(numeros)
    
    ' === 3. AGRUPAR EN BLOQUES (salto >3 = nuevo bloque) ===
    Set bloques = AgruparEnBloques(numeros)
    
' === 4. SELECCIONAR BLOQUE CON N�MEROS M�S BAJOS ===
    Set bloqueSeleccionado = SeleccionarBloqueMasBajo(bloques)
    
    ' === 5. ASIGNAR N�MERO SEG�N HUECOS EN EL BLOQUE ===
    ObtenerNumeroFacturaInteligente = AsignarDesdeBloque(bloqueSeleccionado, mensajeInfo)
    
    ' === 6. INFORMAR HUECOS EN BLOQUES NO SELECCIONADOS ===
    Dim bloqueInfo As Collection
    Dim infoOtrosBloques As String
    Dim minB As Long, maxB As Long
    Dim huecoB As String
    Dim jj As Long
    Dim numB As Variant
    Dim encontradoB As Boolean
    infoOtrosBloques = ""
    
    For Each bloqueInfo In bloques
        ' Saltar el bloque seleccionado
        If bloqueInfo(1) = bloqueSeleccionado(1) And _
           bloqueInfo(bloqueInfo.Count) = bloqueSeleccionado(bloqueSeleccionado.Count) Then
            GoTo SiguienteBloque
        End If
        
        minB = bloqueInfo(1)
        maxB = bloqueInfo(bloqueInfo.Count)
        huecoB = ""
        
        For jj = minB To maxB
            encontradoB = False
            For Each numB In bloqueInfo
                If CLng(numB) = jj Then
                    encontradoB = True
                    Exit For
                End If
            Next numB
            If Not encontradoB Then
                If huecoB = "" Then
                    huecoB = CStr(jj)
                Else
                    huecoB = huecoB & ", " & CStr(jj)
                End If
            End If
        Next jj
        
        If huecoB <> "" Then
            If infoOtrosBloques = "" Then
                infoOtrosBloques = "Huecos en bloques anteriores:" & vbCrLf
            End If
            infoOtrosBloques = infoOtrosBloques & "  Bloque [" & minB & "-" & maxB & "]: falta " & huecoB & vbCrLf
        End If
        
SiguienteBloque:
    Next bloqueInfo
    
    If infoOtrosBloques <> "" Then
        mensajeInfo = mensajeInfo & vbCrLf & vbCrLf & "** " & infoOtrosBloques
    End If
End Function

' ===================================================================================
' Funci�n: AgruparEnBloques
' Agrupa n�meros en bloques separados por saltos >100
' ===================================================================================
Private Function AgruparEnBloques(numeros() As Long) As Collection
    Dim bloques As New Collection
    Dim bloqueActual As New Collection
    Dim i As Long
    
    ' A�adir primer n�mero al primer bloque
    bloqueActual.Add numeros(1)
    
    ' Recorrer resto de n�meros
    For i = 2 To UBound(numeros)
        ' Si el salto es >100, cerrar bloque actual y crear uno nuevo
        If numeros(i) - numeros(i - 1) > 100 Then
            bloques.Add bloqueActual
            Set bloqueActual = New Collection
        End If
        bloqueActual.Add numeros(i)
    Next i
    
    ' A�adir �ltimo bloque
    bloques.Add bloqueActual
    
    Set AgruparEnBloques = bloques
End Function

' ===================================================================================
' Funci�n: SeleccionarBloqueMasBajo
' Selecciona el bloque cuyo m�ximo sea el menor de todos
' ===================================================================================
Private Function SeleccionarBloqueMasBajo(bloques As Collection) As Collection
    Dim bloque As Collection
    Dim minimoMax As Long
    Dim maxActual As Long
    Dim bloqueGanador As Collection
    
    minimoMax = 2147483647 ' Valor m�ximo Long
    
    For Each bloque In bloques
        maxActual = bloque(bloque.Count) ' �ltimo elemento = m�ximo (ya ordenado)
        If maxActual < minimoMax Then
            minimoMax = maxActual
            Set bloqueGanador = bloque
        End If
    Next bloque
    
    Set SeleccionarBloqueMasBajo = bloqueGanador
End Function

' ===================================================================================
' Funci�n: AsignarDesdeBloque
' Busca huecos desde 1 o asigna siguiente consecutivo
' ===================================================================================
Private Function AsignarDesdeBloque(bloque As Collection, ByRef mensajeInfo As String) As Long
    Dim numeroMinimo As Long
    Dim numeroMaximo As Long
    Dim i As Long
    Dim huecos As String
    Dim contadorHuecos As Long
    Dim encontrado As Boolean
    
    numeroMinimo = bloque(1)
    numeroMaximo = bloque(bloque.Count)
    
    ' === BUSCAR HUECOS DESDE 1 HASTA EL M�XIMO DEL BLOQUE ===
    huecos = ""
    contadorHuecos = 0
    
    For i = numeroMinimo To numeroMaximo
        encontrado = False
        
        ' Verificar si el n�mero i existe en el bloque
        Dim num As Variant
        For Each num In bloque
            If CLng(num) = i Then
                encontrado = True
                Exit For
            End If
        Next num
        
        ' Si no existe, es un hueco
        If Not encontrado Then
            contadorHuecos = contadorHuecos + 1
            If huecos = "" Then
                huecos = CStr(i)
            Else
                huecos = huecos & ", " & CStr(i)
            End If
        End If
    Next i
    
    ' === DECIDIR QU� ASIGNAR ===
Dim rangoSerie As String
    If numeroMinimo = numeroMaximo Then
        rangoSerie = CStr(numeroMinimo)
    Else
        rangoSerie = CStr(numeroMinimo) & "-" & CStr(numeroMaximo)
    End If

    If contadorHuecos > 0 Then
        AsignarDesdeBloque = CLng(Split(huecos, ",")(0))
        mensajeInfo = "Existe un hueco en la serie (" & rangoSerie & ") y el primero disponible es (" & AsignarDesdeBloque & ")"
    Else
        AsignarDesdeBloque = numeroMaximo + 1
        mensajeInfo = "Se ha asignado el siguiente de la serie (" & rangoSerie & ")"
    End If
End Function

' ===================================================================================
' Procedimiento: OrdenarArray (Bubble Sort simple)
' ===================================================================================
Private Sub OrdenarArray(arr() As Long)
    Dim i As Long, j As Long
    Dim temp As Long
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

' ===================================================================================
' Procedimiento: ValidarFacturaManual
' Valida n�meros de factura escritos manualmente y detecta huecos
' ===================================================================================
Public Sub ValidarFacturaManual(ws As Worksheet, numeroEscrito As Long, columnaFactura As String)
    Dim numeros() As Long
    Dim bloques As Collection
    Dim bloque As Collection
    Dim i As Long, ultimaFila As Long
    Dim valor As Variant
    Dim contador As Long
    Dim minBloque As Long, maxBloque As Long
    Dim huecos As String, infoOtrosBloques As String
    Dim numHueco As Long
    Dim encontrado As Boolean
    Dim numB As Variant
    Dim esNuevoBloque As Boolean

    ' === 1. ESCANEAR N�MEROS EXISTENTES (excluyendo el reci�n escrito) ===
    ultimaFila = ws.Cells(ws.Rows.Count, columnaFactura).End(xlUp).Row
    ReDim numeros(1 To ultimaFila)
    contador = 0

    For i = 2 To ultimaFila
        valor = ws.Cells(i, columnaFactura).Value
        If IsNumeric(valor) Then
            If CLng(valor) > 0 And CLng(valor) <> numeroEscrito Then
                contador = contador + 1
                numeros(contador) = CLng(valor)
            End If
        End If
    Next i

    ' Si no hay otras facturas, no hay nada que validar
    If contador = 0 Then Exit Sub

    ReDim Preserve numeros(1 To contador)

    ' === 2. ORDENAR Y AGRUPAR EN BLOQUES ===
    Call OrdenarArray(numeros)
    Set bloques = AgruparEnBloques(numeros)

    ' === 3. COMPROBAR SI EL N�MERO ESCRITO INICIA UN NUEVO BLOQUE ===
    ' Un n�mero inicia nuevo bloque si el m�nimo de todos los bloques
    ' existentes es m�s de 100 veces mayor que el n�mero escrito
    Dim minimoGlobal As Long
    minimoGlobal = numeros(1) ' Ya est�n ordenados, el primero es el m�nimo global
    esNuevoBloque = (minimoGlobal > numeroEscrito * 100)

    ' === 4A. SI ES NUEVO BLOQUE: solo informar huecos en bloques existentes ===
    If esNuevoBloque Then
        infoOtrosBloques = ""
        For Each bloque In bloques
            minBloque = bloque(1)
            maxBloque = bloque(bloque.Count)
            huecos = ""
            For numHueco = minBloque To maxBloque
                encontrado = False
                For Each numB In bloque
                    If CLng(numB) = numHueco Then
                        encontrado = True
                        Exit For
                    End If
                Next numB
                If Not encontrado Then
                    If huecos = "" Then
                        huecos = CStr(numHueco)
                    Else
                        huecos = huecos & ", " & CStr(numHueco)
                    End If
                End If
            Next numHueco
            If huecos <> "" Then
                If infoOtrosBloques = "" Then
                    infoOtrosBloques = "**Huecos en bloques anteriores:" & vbCrLf
                End If
                infoOtrosBloques = infoOtrosBloques & "  Bloque [" & minBloque & "-" & maxBloque & "]: falta " & huecos & vbCrLf
            End If
        Next bloque

        If infoOtrosBloques <> "" Then
            MsgBox "N� de factura " & numeroEscrito & " inicia una nueva serie." & vbCrLf & vbCrLf & _
                   infoOtrosBloques, vbInformation, "Nueva serie de facturas"
        End If
        Exit Sub
    End If

    ' === 4B. SI NO ES NUEVO BLOQUE: comportamiento original con huecos ===
    Dim bloquePertenece As Collection
    Set bloquePertenece = Nothing

    For Each bloque In bloques
        minBloque = bloque(1)
        maxBloque = bloque(bloque.Count)
        If (numeroEscrito >= minBloque And numeroEscrito <= maxBloque) Or _
           (numeroEscrito > maxBloque And numeroEscrito - maxBloque <= 10) Then
            Set bloquePertenece = bloque
            Exit For
        End If
    Next bloque

    If bloquePertenece Is Nothing Then Exit Sub

    maxBloque = bloquePertenece(bloquePertenece.Count)

    If numeroEscrito > maxBloque Then
        huecos = ""
        For numHueco = maxBloque + 1 To numeroEscrito - 1
            If huecos = "" Then
                huecos = CStr(numHueco)
            Else
                huecos = huecos & ", " & CStr(numHueco)
            End If
        Next numHueco

        If huecos <> "" Then
            MsgBox "ADVERTENCIA: N�meros de factura faltantes" & vbCrLf & vbCrLf & _
                   "Ha introducido el n�mero: " & numeroEscrito & vbCrLf & _
                   "N� factura m�xima anterior del bloque: " & maxBloque & vbCrLf & vbCrLf & _
                   "N�meros faltantes: " & huecos & vbCrLf & vbCrLf & _
                   "Valore rellenar estos huecos para mantener la secuencia num�rica.", _
                   vbExclamation, "Huecos en Numeraci�n de Facturas"
        End If
    End If
End Sub
