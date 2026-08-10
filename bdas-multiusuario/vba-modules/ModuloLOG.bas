Attribute VB_Name = "ModuloLOG"
'Attribute VB_Name = "ModuloLOG"
Option Explicit

' ============================================================================
' ModuloLOG: Gesti�n centralizada de registro y mantenimiento de hojas LOG
' ============================================================================
' CORRECCI�N: Obtener contrase�a de ModuloConfigSegura (consistencia con frmLogin)
' ACTUALIZACI�N: Desdoblar columna Celda en N� ORDEN y COLUMNA
' ============================================================================

Private nombresHojasLOG As Variant

' Inicializa array de nombres de hojas LOG si est� vac�o
Public Sub InicializarNombresHojasLOG()
    If IsEmpty(nombresHojasLOG) Then
        nombresHojasLOG = Array("LOG_GIJ�N", "LOG_SOTO", "LOG_OVIEDO")
    End If
End Sub

' Establece protecci�n en las hojas LOG con UserInterfaceOnly = True
Public Sub EstablecerProteccionLOGs()
    Dim j As Long
    Dim wsLOG As Worksheet
    Dim logPassword As String

    ' ? CORRECCI�N: Obtener contrase�a de forma segura
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    
    InicializarNombresHojasLOG

    For j = LBound(nombresHojasLOG) To UBound(nombresHojasLOG)
        On Error Resume Next
        Set wsLOG = ThisWorkbook.Worksheets(nombresHojasLOG(j))
        On Error GoTo 0

        If Not wsLOG Is Nothing Then
            On Error Resume Next
            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
            On Error GoTo 0
            Set wsLOG = Nothing
        End If
    Next j
End Sub

' Obtiene el nombre de la hoja RESIDENCIA correspondiente seg�n la hoja LOG
Private Function ObtenerNombreHojaResidencia(ByVal nombreHojaLog As String) As String
    Select Case nombreHojaLog
        Case "LOG_GIJ�N"
            ObtenerNombreHojaResidencia = "RESIDENCIA GIJ�N"
        Case "LOG_SOTO"
            ObtenerNombreHojaResidencia = "RESIDENCIA SOTO"
        Case "LOG_OVIEDO"
            ObtenerNombreHojaResidencia = "RESIDENCIA OVIEDO"
        Case Else
            ObtenerNombreHojaResidencia = ""
    End Select
End Function

' RegistrarCambioLOG: registra cambios en la hoja de LOG indicada
' ACTUALIZADO: Ahora registra N� ORDEN (columna C) y COLUMNA (columna D) por separado
Public Sub RegistrarCambioLOG(ByVal Target As Range, ByVal usuario As String, ByVal valorAnterior As Variant, ByVal nombreHojaLog As String)
    Dim wsLOG As Worksheet
    Dim wsResidencia As Worksheet
    Dim c As Range
    Dim filaLog As Long
    Dim errNum As Long
    Dim logPassword As String
    Dim nombreHojaResidencia As String
    Dim numeroOrden As Variant
    Dim nombreColumna As Variant
    Dim filaResidencia As Long
    
    ' ? CORRECCI�N: Obtener contrase�a de forma segura
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()

    On Error Resume Next
    Set wsLOG = ThisWorkbook.Worksheets(nombreHojaLog)
    errNum = Err.Number
    On Error GoTo 0

    If wsLOG Is Nothing Then
        MsgBox "Atenci�n: La hoja de registro '" & nombreHojaLog & "' no fue encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener hoja RESIDENCIA correspondiente
    nombreHojaResidencia = ObtenerNombreHojaResidencia(nombreHojaLog)
    If nombreHojaResidencia = "" Then
        MsgBox "Atenci�n: No se pudo determinar la hoja RESIDENCIA para '" & nombreHojaLog & "'.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsResidencia = ThisWorkbook.Worksheets(nombreHojaResidencia)
    On Error GoTo 0
    
    If wsResidencia Is Nothing Then
        MsgBox "Atenci�n: La hoja '" & nombreHojaResidencia & "' no fue encontrada.", vbExclamation
        Exit Sub
    End If

    ' Intento directo: funciona si UserInterfaceOnly=True
    On Error Resume Next
    Dim resCodeLog As String
    Select Case nombreHojaLog
        Case "LOG_GIJÓN", "LOG_GIJON": resCodeLog = "GIJON"
        Case "LOG_SOTO": resCodeLog = "SOTO"
        Case "LOG_OVIEDO": resCodeLog = "OVIEDO"
        Case Else: resCodeLog = "GIJON"
    End Select

    For Each c In Target.Cells
        filaResidencia = c.Row
        numeroOrden = wsResidencia.Cells(filaResidencia, 1).Value
        nombreColumna = wsResidencia.Cells(1, c.Column).Value
        filaLog = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
        wsLOG.Range(wsLOG.Cells(filaLog, 1), wsLOG.Cells(filaLog, 6)).Value = _
            Array(usuario, Now, numeroOrden, nombreColumna, valorAnterior, c.Value)
        
        ' Registrar audit trail centralizado en Access DB
        modDatabase.InsertarLog resCodeLog, usuario, CStr(numeroOrden), CStr(nombreColumna), CStr(valorAnterior), CStr(c.Value)
    Next c
    errNum = Err.Number
    On Error GoTo 0

    If errNum <> 0 Then
        ' Fallback seguro: desproteger, escribir y reproteger
        On Error GoTo ErrHandlerFallback
        wsLOG.Unprotect password:=logPassword
        For Each c In Target.Cells
            filaResidencia = c.Row
            numeroOrden = wsResidencia.Cells(filaResidencia, 1).Value
            nombreColumna = wsResidencia.Cells(1, c.Column).Value
            filaLog = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            wsLOG.Range(wsLOG.Cells(filaLog, 1), wsLOG.Cells(filaLog, 6)).Value = _
                Array(usuario, Now, numeroOrden, nombreColumna, valorAnterior, c.Value)
        Next c
        wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If

    Exit Sub

ErrHandlerFallback:
    MsgBox "No se pudo registrar el cambio en '" & nombreHojaLog & "': " & Err.Description, vbExclamation
    On Error GoTo 0
End Sub

' PrepararHojaLOG: limpieza de registros >45 d�as y asegurar protecci�n final
Public Sub PrepararHojaLOG()
    Dim wsLOG As Worksheet
    Dim filaUltima As Long, i As Long
    Dim fechaRegistro As Variant
    Dim j As Long
    Dim logPassword As String
    
    ' ? CORRECCI�N: Obtener contrase�a de forma segura
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()

    InicializarNombresHojasLOG

    For j = LBound(nombresHojasLOG) To UBound(nombresHojasLOG)
        On Error Resume Next
        Set wsLOG = ThisWorkbook.Worksheets(nombresHojasLOG(j))
        On Error GoTo 0

        If Not wsLOG Is Nothing Then
            ' Desproteger temporalmente para limpieza si est� protegida
            On Error Resume Next
            wsLOG.Unprotect password:=logPassword
            On Error GoTo 0

            ' Eliminar registros con fecha anterior a -45 d�as
            filaUltima = wsLOG.Cells(wsLOG.Rows.Count, "B").End(xlUp).Row
            For i = filaUltima To 2 Step -1
                fechaRegistro = wsLOG.Cells(i, 2).Value
                If IsDate(fechaRegistro) Then
                    If fechaRegistro < DateAdd("d", -45, Date) Then
                        wsLOG.Rows(i).Delete
                    End If
                End If
            Next i

            ' Asegurar protecci�n con UserInterfaceOnly al finalizar
            On Error Resume Next
            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
            On Error GoTo 0

            Set wsLOG = Nothing
        End If
    Next j
End Sub

'=================================================================================
' FUNCION: Registrar accion de ocultar/mostrar habitacion en LOG
' Parametros:
'   - residencia: "GIJON", "SOTO" o "OVIEDO"
'   - habitacion: Nombre de la habitacion (ej: "5", "Ap.  7", "Of.1")
'   - accion: "OCULTAR" o "MOSTRAR"
'   - estadoAnterior: "VISIBLE" o "OCULTA"
'   - estadoNuevo: "VISIBLE" o "OCULTA"
'   - usuario: Nombre del usuario (opcional, por defecto "Sistema")
'=================================================================================
Public Sub RegistrarAccionHabitacionLOG(ByVal residencia As String, _
                                        ByVal habitacion As String, _
                                        ByVal accion As String, _
                                        ByVal estadoAnterior As String, _
                                        ByVal estadoNuevo As String, _
                                        Optional ByVal usuario As String = "Sistema")
    On Error GoTo ErrorHandler
    
    Dim wsLOG As Worksheet
    Dim nombreHojaLog As String
    Dim filaLog As Long
    Dim logPassword As String
    Dim tipoHabitacion As String
    
    ' Determinar nombre de hoja LOG segun residencia
Select Case UCase(residencia)
    Case "GIJON", "GIJ�N"
        nombreHojaLog = "LOG_GIJ�N"
        tipoHabitacion = "HABITACION"
    Case "SOTO"
        nombreHojaLog = "LOG_SOTO"
        tipoHabitacion = "APARTAMENTO"
    Case "OVIEDO"
        nombreHojaLog = "LOG_OVIEDO"
        tipoHabitacion = "HABITACION"
    Case Else
        Exit Sub
End Select
    
    ' Obtener contrase�a de forma segura
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    
    ' Obtener hoja de LOG
    On Error Resume Next
    Set wsLOG = ThisWorkbook.Worksheets(nombreHojaLog)
    On Error GoTo ErrorHandler
    
    If wsLOG Is Nothing Then
        ' Si la hoja no existe, salir silenciosamente (no bloquear la operacion principal)
        Exit Sub
    End If
    
    ' Intentar escribir directamente (funciona si UserInterfaceOnly=True)
    On Error Resume Next
    filaLog = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
    wsLOG.Cells(filaLog, 1).Value = usuario
    wsLOG.Cells(filaLog, 2).Value = Now
    wsLOG.Cells(filaLog, 3).Value = tipoHabitacion
    wsLOG.Cells(filaLog, 4).Value = habitacion
    wsLOG.Cells(filaLog, 5).Value = estadoAnterior
    wsLOG.Cells(filaLog, 6).Value = estadoNuevo
    
    ' Si fallo, usar metodo fallback (desproteger/proteger)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ErrorHandler
        
        wsLOG.Unprotect password:=logPassword
        
        filaLog = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
        wsLOG.Cells(filaLog, 1).Value = usuario
        wsLOG.Cells(filaLog, 2).Value = Now
        wsLOG.Cells(filaLog, 3).Value = tipoHabitacion
        wsLOG.Cells(filaLog, 4).Value = habitacion
        wsLOG.Cells(filaLog, 5).Value = estadoAnterior
        wsLOG.Cells(filaLog, 6).Value = estadoNuevo
        
        wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    
    On Error GoTo 0
    Exit Sub

ErrorHandler:
    ' Error al registrar: no bloquear la operacion principal
    ' Simplemente salir silenciosamente
    On Error Resume Next
    If Not wsLOG Is Nothing Then
        wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    On Error GoTo 0
End Sub

