Attribute VB_Name = "ModuloGestionHabitaciones"
'Attribute VB_Name = "ModuloGestionHabitaciones"
'=================================================================================
' M�DULO: ModuloGestionHabitaciones
' Prop�sito: Ocultar/Mostrar habitaciones en calendarios de forma interactiva
' Fecha: 2025-11-28
' Autor: Sistema de Gesti�n Residencias Asturias
'=================================================================================

Option Explicit

'=================================================================================
' FUNCIONES P�BLICAS: Puntos de entrada desde Ribbon (uno por residencia)
'=================================================================================

Public Sub GestionarHabitacionesGijon(control As IRibbonControl)
    Call GestionarHabitacionesPorResidencia("GIJ�N")
End Sub

Public Sub GestionarHabitacionesSoto(control As IRibbonControl)
    Call GestionarHabitacionesPorResidencia("SOTO")
End Sub

Public Sub GestionarHabitacionesOviedo(control As IRibbonControl)
    Call GestionarHabitacionesPorResidencia("OVIEDO")
End Sub

'=================================================================================
' FUNCI�N PRINCIPAL: Gestionar habitaciones por residencia
'=================================================================================
Private Sub GestionarHabitacionesPorResidencia(residencia As String)
    On Error GoTo ErrorHandler
    
    Dim wsCal As Worksheet
    Dim nombreHojaCalendario As String
    
    ' Determinar nombre de hoja de calendario
    nombreHojaCalendario = "Calendario " & residencia
    
    ' Verificar que existe la hoja
    On Error Resume Next
    Set wsCal = ThisWorkbook.Worksheets(nombreHojaCalendario)
    On Error GoTo ErrorHandler
    
    If wsCal Is Nothing Then
        MsgBox "No se encontr� la hoja '" & nombreHojaCalendario & "'.", _
               vbCritical, "Error"
        Exit Sub
    End If
    
    ' Activar la hoja de calendario si no est� activa
    If ActiveSheet.Name <> nombreHojaCalendario Then
        wsCal.Activate
    End If
    
    ' Mostrar formulario de gesti�n
    Call MostrarFormularioGestion(wsCal, residencia)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al ejecutar Gesti�n de Habitaciones:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
End Sub

'=================================================================================
' FUNCI�N PRIVADA: Mostrar formulario de gesti�n
'=================================================================================
Private Sub MostrarFormularioGestion(wsCal As Worksheet, residencia As String)
    On Error GoTo ErrorHandler
    
    Dim habitacionesConfig As Variant
    Dim listaHabitaciones As String
    Dim habitacionSeleccionada As String
    Dim filaHabitacion As Long
    Dim estaOculta As Boolean
    Dim accion As String
    Dim respuesta As VbMsgBoxResult
    
    ' Obtener configuraci�n de habitaciones seg�n residencia
    habitacionesConfig = ObtenerConfiguracionHabitaciones(residencia)
    
    If IsEmpty(habitacionesConfig) Then
        MsgBox "No se pudo cargar la configuraci�n de habitaciones.", vbCritical
        Exit Sub
    End If
    
    ' Construir lista de habitaciones con su estado (Visible/Oculta)
    listaHabitaciones = ConstruirListaHabitacionesConEstado(wsCal, habitacionesConfig, residencia)
    
    ' Solicitar selecci�n de habitaci�n
    habitacionSeleccionada = InputBox( _
        "GESTI�N DE ALOJAMIENTO - " & residencia & vbCrLf & vbCrLf & _
        listaHabitaciones & vbCrLf & vbCrLf & _
        "Introduce el alojamiento a ocultar/mostrar:" & vbCrLf & _
        "(Escribe exactamente como aparece en la lista)", _
        "Gesti�n de Alojamiento - " & residencia)
    
    ' Validar entrada
    If Trim(habitacionSeleccionada) = "" Then Exit Sub ' Usuario cancel�
    
    habitacionSeleccionada = Trim(habitacionSeleccionada)
    
    ' Obtener fila correspondiente
    filaHabitacion = ObtenerFilaHabitacion(habitacionSeleccionada, habitacionesConfig)
    
    If filaHabitacion = 0 Then
        MsgBox "Habitaci�n '" & habitacionSeleccionada & "' no v�lida." & vbCrLf & vbCrLf & _
               "Verifica que el nombre sea exacto (distingue may�sculas/min�sculas).", _
               vbExclamation, "Habitaci�n No V�lida"
        Exit Sub
    End If
    
    ' Verificar estado actual
    estaOculta = wsCal.Rows(filaHabitacion).Hidden
    
    ' Determinar acci�n a realizar
    If estaOculta Then
        accion = "MOSTRAR"
        respuesta = MsgBox("La habitaci�n '" & habitacionSeleccionada & "' est� actualmente OCULTA." & vbCrLf & vbCrLf & _
                          "�Deseas MOSTRARLA en el calendario?", _
                          vbQuestion + vbYesNo, "Mostrar Habitaci�n")
    Else
        accion = "OCULTAR"
        
        ' === NUEVO: Verificar reservas futuras antes de ocultar ===
        Dim reservasFuturas As Variant
        Dim mensajeOcultar As String
        Dim i As Long
        
        reservasFuturas = ObtenerReservasFuturasEnHabitacion(residencia, habitacionSeleccionada)
        
        ' Construir mensaje seg�n si hay reservas o no
        If IsEmpty(reservasFuturas) Then
            ' No hay reservas futuras
            mensajeOcultar = "La habitaci�n '" & habitacionSeleccionada & "' est� actualmente VISIBLE." & vbCrLf & vbCrLf & _
                           "No hay reservas futuras en esta habitaci�n." & vbCrLf & vbCrLf & _
                           "�Deseas OCULTARLA del calendario?"
        Else
            ' Hay reservas futuras - construir lista detallada
            mensajeOcultar = "La habitaci�n '" & habitacionSeleccionada & "' est� actualmente VISIBLE." & vbCrLf & vbCrLf & _
                           "ATENCI�N: Esta habitaci�n tiene adjudicadas " & UBound(reservasFuturas, 1) & " reserva(s):" & vbCrLf & vbCrLf
            
            For i = 1 To UBound(reservasFuturas, 1)
                Dim nombreCliente As String
                nombreCliente = Trim(CStr(reservasFuturas(i, 2)))
                
                ' Limitar nombre a 20 caracteres si es muy largo
                If Len(nombreCliente) > 20 Then
                    nombreCliente = Left(nombreCliente, 20) & "..."
                End If
                
                ' Si no hay nombre, mostrar solo "Sin nombre"
                If nombreCliente = "" Then nombreCliente = "Sin nombre"
                
           mensajeOcultar = mensajeOcultar & _
                "  � Reserva N� " & reservasFuturas(i, 1) & ": " & _
                Format(reservasFuturas(i, 3), "dd/mm/yyyy") & " a " & _
                Format(reservasFuturas(i, 4), "dd/mm/yyyy") & vbCrLf
            Next i
            
            mensajeOcultar = mensajeOcultar & vbCrLf & _
                        "Si la ocultas, las reservas NO se mostrar�n en el calendario." & vbCrLf & vbCrLf & _
                        "�Deseas continuar y OCULTARLA?"
        End If
        
        respuesta = MsgBox(mensajeOcultar, vbQuestion + vbYesNo, "Ocultar Habitaci�n")
    End If
    
    ' Ejecutar acci�n si el usuario confirma
    If respuesta = vbYes Then
        Application.screenUpdating = False
        
        Dim estadoAnterior As String
        Dim estadoNuevo As String
        
        If estaOculta Then
            ' Mostrar fila
            wsCal.Rows(filaHabitacion).Hidden = False
            
            estadoAnterior = "OCULTA"
            estadoNuevo = "VISIBLE"
            
            ' Registrar en LOG
            On Error Resume Next
            Call ModuloLOG.RegistrarAccionHabitacionLOG(residencia, habitacionSeleccionada, "MOSTRAR", estadoAnterior, estadoNuevo, ThisWorkbook.usuarioActual)
            On Error GoTo 0
            
            ' === NUEVO: Actualizar calendario para mostrar reservas ===
            Call ActualizarCalendarioDespuesDeGestion(residencia)
            
            MsgBox "Habitaci�n '" & habitacionSeleccionada & "' MOSTRADA correctamente." & vbCrLf & vbCrLf & _
                   "Ahora aparecer� en el calendario y en las selecciones.", _
                   vbInformation, "Habitaci�n Mostrada"
        Else
            ' Ocultar fila
            wsCal.Rows(filaHabitacion).Hidden = True
            
            estadoAnterior = "VISIBLE"
            estadoNuevo = "OCULTA"
            
            ' Registrar en LOG
            On Error Resume Next
            Call ModuloLOG.RegistrarAccionHabitacionLOG(residencia, habitacionSeleccionada, "OCULTAR", estadoAnterior, estadoNuevo, ThisWorkbook.usuarioActual)
            On Error GoTo 0
            
            ' === NUEVO: Actualizar calendario y verificar reservas ocultas ===
            Call ActualizarCalendarioDespuesDeGestion(residencia)
            
            MsgBox "Habitaci�n '" & habitacionSeleccionada & "' OCULTA correctamente." & vbCrLf & vbCrLf & _
                   "Ya no aparecer� en el calendario ni en las selecciones.", _
                   vbInformation, "Habitaci�n Oculta"
        End If
        
        Application.screenUpdating = True
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.screenUpdating = True
    MsgBox "Error en gesti�n de habitaciones:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
End Sub

'=================================================================================
' FUNCI�N: Obtener configuraci�n de habitaciones por residencia
' Retorna: Array con pares (NombreHabitaci�n, FilaCalendario)
'=================================================================================
Private Function ObtenerConfiguracionHabitaciones(residencia As String) As Variant
    On Error Resume Next
    
    Dim config As Variant
    
    Select Case UCase(residencia)
        Case "GIJ�N"
            config = Array( _
                "1", 2, _
                "2", 4, _
                "3", 6, _
                "4", 8, _
                "5", 10, _
                "6", 12, _
                "7", 14, _
                "Of.1", 16, _
                "Of.2", 18, _
                "Of.3", 20, _
                "Est.1", 22, _
                "Est.2", 24, _
                "Est.3", 26 _
            )
        
        Case "SOTO"
            ' TODO: Adaptar seg�n estructura de SOTO
            ' Ejemplo (ajusta seg�n tus filas reales):
            config = Array( _
                "Ap.2", 2, _
                "Ap.4", 4, _
                "Ap.5", 6, _
                "Ap.6", 8, _
                "Ap.7", 10, _
                "Ap.8", 12, _
                "Ap.9", 14, _
                "Ap.10", 16, _
                "Ap.12", 18 _
            )
        
        Case "OVIEDO"
            ' TODO: Adaptar seg�n estructura de OVIEDO
            ' Ejemplo (ajusta seg�n tus filas reales):
            config = Array( _
                "1", 2, _
                "2", 4, _
                "3", 6, _
                "4", 8, _
                "5", 10, _
                "6", 12, _
                "7", 14, _
                "8", 16, _
                "9", 18, _
                "10", 20, _
                "11", 22, _
                "12", 24, _
                "13", 26, _
                "14", 28, _
                "15", 30 _
            )

        
        Case Else
            config = Empty
    End Select
    
    ObtenerConfiguracionHabitaciones = config
End Function

'=================================================================================
' FUNCI�N: Construir lista de habitaciones con estado visible
'=================================================================================
Private Function ConstruirListaHabitacionesConEstado(wsCal As Worksheet, config As Variant, residencia As String) As String
    On Error Resume Next
    
    Dim i As Long
    Dim nombreHab As String
    Dim filaHab As Long
    Dim estado As String
    Dim lista As String
    
    lista = "HABITACIONES/APARTAMENTOS DISPONIBLES:" & vbCrLf & _
            String(50, "-") & vbCrLf
    
    For i = 0 To UBound(config) Step 2
        nombreHab = CStr(config(i))
        filaHab = CLng(config(i + 1))
        
        If wsCal.Rows(filaHab).Hidden Then
            estado = "[OCULTA]"
        Else
            estado = "[VISIBLE]"
        End If
        
        lista = lista & "� " & nombreHab & " " & estado & vbCrLf
    Next i
    
    ConstruirListaHabitacionesConEstado = lista
End Function

'=================================================================================
' FUNCI�N: Obtener n�mero de fila de una habitaci�n
'=================================================================================
Private Function ObtenerFilaHabitacion(nombreHab As String, config As Variant) As Long
    On Error Resume Next
    
    Dim i As Long
    ObtenerFilaHabitacion = 0
    
    For i = 0 To UBound(config) Step 2
        If StrComp(CStr(config(i)), nombreHab, vbTextCompare) = 0 Then
            ObtenerFilaHabitacion = CLng(config(i + 1))
            Exit Function
        End If
    Next i
End Function

'=================================================================================
' FUNCI�N: Obtener detalles de reservas futuras en una habitaci�n espec�fica
' Retorna: Array bidimensional con (NumOrden, Nombre, FechaEntrada, FechaSalida)
'          O Empty si no hay reservas
'=================================================================================
Private Function ObtenerReservasFuturasEnHabitacion(residencia As String, habitacion As String) As Variant
    On Error Resume Next
    
    Dim wsRes As Worksheet
    Dim nombreHojaReservas As String
    Dim lastRow As Long, i As Long
    Dim fechaSalida As Date, hoy As Date
    Dim habitaciones As String, arrHabs As Variant
    Dim hab As Variant
    Dim reservasTemp As Collection
    Dim reservaInfo As Variant
    Dim resultado() As Variant
    Dim contador As Long
    
    ' Determinar nombre de hoja de reservas seg�n residencia
    Select Case UCase(residencia)
        Case "GIJ�N"
            nombreHojaReservas = "RESIDENCIA GIJ�N"
        Case "SOTO"
            nombreHojaReservas = "RESIDENCIA SOTO"
        Case "OVIEDO"
            nombreHojaReservas = "RESIDENCIA OVIEDO"
        Case Else
            ObtenerReservasFuturasEnHabitacion = Empty
            Exit Function
    End Select
    
    ' Obtener hoja de reservas
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(nombreHojaReservas)
    On Error GoTo 0
    
    If wsRes Is Nothing Then
        ObtenerReservasFuturasEnHabitacion = Empty
        Exit Function
    End If
    
    hoy = Date
    Set reservasTemp = New Collection
    
    ' Obtener �ltima fila con datos
    lastRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        ObtenerReservasFuturasEnHabitacion = Empty
        Exit Function
    End If
    
    ' Determinar columna de habitaciones seg�n residencia
    Dim colHabitaciones As Long
    Select Case UCase(residencia)
        Case "GIJ�N"
            colHabitaciones = 19  ' Columna S
        Case "SOTO"
            colHabitaciones = 19  ' Columna S (ajusta si es diferente)
        Case "OVIEDO"
            colHabitaciones = 20  ' Columna T
    End Select
    
    ' Recorrer todas las reservas
    For i = 2 To lastRow
        ' Verificar que tenga fecha de salida y habitaciones asignadas
        If IsDate(wsRes.Cells(i, "M").Value) And Not IsEmpty(wsRes.Cells(i, colHabitaciones).Value) Then
            fechaSalida = wsRes.Cells(i, "M").Value
            
            ' Solo procesar si la fecha de salida es HOY o FUTURA
            If fechaSalida >= hoy Then
                habitaciones = Trim(CStr(wsRes.Cells(i, colHabitaciones).Value))
                
                ' Verificar si la habitaci�n buscada est� en esta reserva
                Dim habitacionEncontrada As Boolean
                habitacionEncontrada = False
                
                If InStr(habitaciones, ",") > 0 Then
                    ' M�ltiples habitaciones separadas por comas
                    arrHabs = Split(habitaciones, ",")
                    For Each hab In arrHabs
                        If Trim(CStr(hab)) = habitacion Then
                            habitacionEncontrada = True
                            Exit For
                        End If
                    Next hab
                Else
                    ' Una sola habitaci�n
                    If habitaciones = habitacion Then
                        habitacionEncontrada = True
                    End If
                End If
                
                ' Si encontramos la habitaci�n, a�adir a la colecci�n
                If habitacionEncontrada Then
                    ReDim reservaInfo(0 To 3)
                    reservaInfo(0) = wsRes.Cells(i, 1).Value  ' N� Orden (columna A)
                    reservaInfo(1) = wsRes.Cells(i, 11).Value  ' Nombre (columna K)
                    reservaInfo(2) = wsRes.Cells(i, 12).Value ' Fecha Entrada (columna L)
                    reservaInfo(3) = wsRes.Cells(i, 13).Value ' Fecha Salida (columna M)
                    
                    reservasTemp.Add reservaInfo
                End If
            End If
        End If
    Next i
    
    ' Convertir Collection a Array
    If reservasTemp.Count = 0 Then
        ObtenerReservasFuturasEnHabitacion = Empty
        Exit Function
    End If
    
    ReDim resultado(1 To reservasTemp.Count, 1 To 4)
    For contador = 1 To reservasTemp.Count
        reservaInfo = reservasTemp(contador)
        resultado(contador, 1) = reservaInfo(0) ' N� Orden
        resultado(contador, 2) = reservaInfo(1) ' Nombre
        resultado(contador, 3) = reservaInfo(2) ' Fecha Entrada
        resultado(contador, 4) = reservaInfo(3) ' Fecha Salida
    Next contador
    
    ObtenerReservasFuturasEnHabitacion = resultado
End Function

Private Sub ActualizarCalendarioDespuesDeGestion(residencia As String)
    On Error Resume Next
    
    Select Case UCase(residencia)
        Case "GIJ�N", "GIJON"
            Call ModuloCalendarioGijon.ActualizarCalendarioFuturoGijon
        Case "SOTO"
            Call ModuloCalendarioSoto.ActualizarCalendarioFuturoSoto
        Case "OVIEDO"
            Call ModuloCalendarioOviedo.ActualizarCalendarioFuturoOviedo
    End Select
    
    On Error GoTo 0
End Sub
