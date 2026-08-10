Attribute VB_Name = "ValidacionFechasSoto"
'=================================================================================
' M�dulo: ValidacionFechasSoto
'
' Descripci�n:
'   Determina qu� apartamentos est�n disponibles en la hoja "RESIDENCIA SOTO"
'   para un rango de fechas y una separaci�n m�nima entre reservas.
'   Versi�n optimizada y robusta con manejo de fila �nica y bloqueos
'=================================================================================

Option Explicit

'---------------------------------------------------------------------------------
' VerificarDisponibilidadSoto
'   Versi�n corregida y optimizada
'---------------------------------------------------------------------------------
Public Sub VerificarDisponibilidadSoto(fechaEntrada As Date, fechaSalida As Date, filaActual As Long, Separacion As Long)
    On Error GoTo ManejoErrores

    '---------------------- Validaci�n de par�metros -----------------------
    If fechaSalida <= fechaEntrada Then
        MsgBox "Error: La fecha de salida debe ser posterior a la de entrada.", vbCritical
        Exit Sub
    End If

    If Separacion < 0 Then
        MsgBox "Error: El valor de separaci�n no puede ser negativo.", vbCritical
        Exit Sub
    End If

    '---------------------- Inicializaci�n de variables ----------------------
    Dim ws As Worksheet
    Dim wsBloq As Worksheet
    Set ws = ThisWorkbook.Sheets("RESIDENCIA SOTO")
    
    ' --- VERIFICAR DUPLICADOS EN LOS �LTIMOS 30 D�AS ---
    Call VerificarDuplicadoUltimos30Dias(ws, filaActual)
    
    ' Si la solicitud est� DESESTIMADA o vac�a tras la validaci�n, salimos tempranamente
    If UCase(Trim(ws.Cells(filaActual, "P").Value)) = "DESESTIMADA" Then Exit Sub
    If IsEmpty(ws.Cells(filaActual, "L").Value) Or IsEmpty(ws.Cells(filaActual, "M").Value) Then Exit Sub
    Set wsBloq = ThisWorkbook.Sheets("Apartamentos Bloqueados SOTO")
    
    ' Lista de apartamentos a comprobar
    Dim apartamentos() As String
    apartamentos = Split("Ap.2, Ap.4, Ap.5, Ap.6, Ap.7, Ap.8, Ap.9, Ap.10, Ap.12", ",")
    
    ' Colecci�n para apartamentos disponibles
    Dim apartamentosDisponibles As Collection
    Set apartamentosDisponibles = New Collection

    '---------------------- Cargar bloqueos en memoria -------------------------------
    Dim bloqueos As Collection
    Set bloqueos = New Collection
    Dim filaBloq As Long
    filaBloq = wsBloq.Cells(wsBloq.Rows.Count, "A").End(xlUp).Row
    
    If filaBloq >= 2 Then
        Dim arrBloq As Variant
        arrBloq = wsBloq.Range("A2:C" & filaBloq).Value
        
        If IsArray(arrBloq) Then
            Dim m As Long
            For m = 1 To UBound(arrBloq, 1)
                If arrBloq(m, 1) <> "" And IsDate(arrBloq(m, 2)) And IsDate(arrBloq(m, 3)) Then
                    ' Almacenar como array de Variant
                    bloqueos.Add Array(CStr(arrBloq(m, 1)), CDate(arrBloq(m, 2)), CDate(arrBloq(m, 3)))
                End If
            Next m
        End If
    End If

    '---------------------- Determinaci�n de la �ltima fila --------------------------
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
    ' Arrays para datos de reservas
    Dim arrEntradas As Variant, arrSalidas As Variant, arrApartamentos As Variant
    
    ' Manejo robusto para 1 fila
    If ultimaFila >= 2 Then
        If ultimaFila = 2 Then
            ' Caso de 1 fila: crear arrays manualmente
            ReDim arrEntradas(1 To 1, 1 To 1)
            arrEntradas(1, 1) = ws.Range("L2").Value
            
            ReDim arrSalidas(1 To 1, 1 To 1)
            arrSalidas(1, 1) = ws.Range("M2").Value
            
            ReDim arrApartamentos(1 To 1, 1 To 1)
            arrApartamentos(1, 1) = ws.Range("S2").Value
        Else
            ' Caso m�ltiples filas
            arrEntradas = ws.Range("L2:L" & ultimaFila).Value
            arrSalidas = ws.Range("M2:M" & ultimaFila).Value
            arrApartamentos = ws.Range("S2:S" & ultimaFila).Value
        End If
    End If

    '---------------------- B�squeda de apartamentos disponibles ---------------------
    Dim k As Long, i As Long, j As Long
    Dim apartamento As String
    Dim apartamentoLibre As Boolean
    Dim bloqueado As Boolean
    Dim apartamentosAdjudicadas As Variant
    Dim fechaEntradaExistente As Date, fechaSalidaExistente As Date
    Dim bloqItem As Variant  ' Variable para iterar bloqueos
    Dim bloqueoInicio As Date, bloqueoFin As Date

    For k = LBound(apartamentos) To UBound(apartamentos)
        apartamento = Trim(apartamentos(k))
        apartamentoLibre = True
        bloqueado = False

        ' 2. Comprobar reservas existentes (solo si hay datos)
        If ultimaFila >= 2 Then
            For i = 1 To UBound(arrEntradas, 1)
                ' Excluir la fila actual
                If (ultimaFila = 2 And filaActual = 2) Or (i + 1 = filaActual) Then GoTo SiguienteReserva
                
                ' Validar fechas
                If Not IsDate(arrEntradas(i, 1)) Or Not IsDate(arrSalidas(i, 1)) Then GoTo SiguienteReserva
                
                fechaEntradaExistente = arrEntradas(i, 1)
                fechaSalidaExistente = arrSalidas(i, 1)
                
                If arrApartamentos(i, 1) <> "" Then
                    apartamentosAdjudicadas = Split(Replace(arrApartamentos(i, 1), " ", ""), ",")
                    
                    For j = LBound(apartamentosAdjudicadas) To UBound(apartamentosAdjudicadas)
                        If Trim(apartamentosAdjudicadas(j)) = apartamento Then
                            ' Comprobar solape con separaci�n
                            If Not (fechaSalida <= DateAdd("d", -Separacion, fechaEntradaExistente) Or _
                                 fechaEntrada >= DateAdd("d", Separacion, fechaSalidaExistente)) Then
                                apartamentoLibre = False
                                Exit For
                            End If
                        End If
                    Next j
                    
                    If Not apartamentoLibre Then Exit For
                End If
SiguienteReserva:
            Next i
        End If

        ' Si sigue libre, a�adir a la colecci�n
        If apartamentoLibre Then
            apartamentosDisponibles.Add apartamento
        End If

SiguienteApartamento:
    Next k

    '---------------------- Resultados finales ------------------------
    If apartamentosDisponibles.Count > 0 Then
        Dim listaDisponibles As String
        listaDisponibles = ""
        
        For i = 1 To apartamentosDisponibles.Count
            listaDisponibles = listaDisponibles & apartamentosDisponibles(i) & ", "
        Next i
        
        listaDisponibles = Left(listaDisponibles, Len(listaDisponibles) - 2)
        
        ' Crear nombre definido
        On Error Resume Next
        ThisWorkbook.Names.Add Name:="ApartamentosDisponibles", _
            RefersToR1C1:="=""" & listaDisponibles & """", _
            visible:=False
        On Error GoTo 0
        
        ' Restaurar entorno Excel ANTES de mostrar el formulario
        ' para que el formulario responda correctamente a los clics del usuario
        Application.enableEvents = True
        Application.screenUpdating = True
        DoEvents
        
        ' Mostrar formulario
        SeleccionHabitacionSOTO.Show
    Else
        MsgBox "No hay apartamentos disponibles para el rango de fechas elegido con una separaci�n de " & Separacion & " d�a(s).", vbInformation
        Call EscribirResolucionDenegadaSoto(ws, filaActual, fechaEntrada)
    End If

    Exit Sub

ManejoErrores:
    MsgBox "Error en VerificarDisponibilidadSoto: " & Err.Description & vbCrLf & _
           "Origen: " & Err.Source & vbCrLf & _
           "N�mero de error: " & Err.Number, vbCritical
End Sub

'------------------------------------------------------------------------------------------
' HaySolapeBloqueoSoto
'   Funci�n optimizada para detecci�n de solapes con bloqueos
'------------------------------------------------------------------------------------------
Public Function HaySolapeBloqueoSoto(reservaInicio As Date, reservaFin As Date, bloqueoInicio As Date, bloqueoFin As Date) As Boolean
    ' Verifica si hay solape entre la reserva y el bloqueo
    ' Considera que el d�a de salida del bloqueo est� incluido
    Dim limiteFin As Date
    If bloqueoInicio = bloqueoFin Then
        limiteFin = bloqueoFin + 1
    Else
        limiteFin = bloqueoFin
    End If
    HaySolapeBloqueoSoto = Not (reservaFin <= bloqueoInicio Or reservaInicio >= limiteFin)
End Function

'------------------------------------------------------------------------------------------
' EscribirResolucionDenegadaSoto
'   Escribe DENEGADA o NO en col. P segun los dias de antelacion entre solicitud y entrada.
'   El umbral varia segun el dia de la semana de la fecha de solicitud (col. B):
'     Lunes=13, Martes=12, Mier=11, Jue=10, Vie=9, Sab=8, Dom=7
'   Si col. P ya contiene un valor definitivo (SI/CONCEDIDA/REEVALUADA/RENUNCIA), no sobreescribe.
'------------------------------------------------------------------------------------------
Public Sub EscribirResolucionDenegadaSoto(ws As Worksheet, fila As Long, fechaEntrada As Date)
    Dim valorActual As String
    valorActual = UCase(Trim(ws.Cells(fila, "P").Value))
' Guard clause: preguntar antes de sobreescribir resoluciones definitivas
    If valorActual = "SI" Or valorActual = "CONCEDIDA" Or _
       valorActual = "REEVALUADA" Or valorActual = "RENUNCIA" Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("La celda ya contiene '" & valorActual & "'." & vbCrLf & _
                      "�Desea sobreescribir la resoluci�n?", _
                      vbYesNo + vbQuestion, "Confirmar sobreescritura")
        If resp = vbNo Then Exit Sub
    End If
    
    If Not IsDate(ws.Cells(fila, "B").Value) Then Exit Sub
    
    Dim fechaSolicitud As Date
    fechaSolicitud = CDate(ws.Cells(fila, "B").Value)
    
    ' Umbral variable segun dia de la semana (L=13, M=12, X=11, J=10, V=9, S=8, D=7)
    Dim umbral As Long
    umbral = 14 - Weekday(fechaSolicitud, vbMonday)
    
    Dim diasAnt As Long
    diasAnt = DateDiff("d", fechaSolicitud, fechaEntrada)
    
        On Error GoTo SafeExit
    Application.enableEvents = False
    If diasAnt <= umbral Then
        ws.Cells(fila, "P").Value = "DENEGADA"
    Else
        ws.Cells(fila, "P").Value = "NO"
    End If
SafeExit:
    Application.enableEvents = True
End Sub
