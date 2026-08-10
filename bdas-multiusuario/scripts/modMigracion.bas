Attribute VB_Name = "modMigracion"
Option Explicit

' ==============================================================================
' Módulo: modMigracion.bas
' Propósito: Migrar datos históricos del Excel BDAS a la base de datos Access.
'            Se ejecuta UNA SOLA VEZ, durante la Fase 3 de la implementación.
'
' Procedimientos:
'   - MigrarTodasLasResidencias()  → Ejecuta la migración completa
'   - MigrarHojaResidencia(...)    → Migra una hoja de residencia a Access
'   - MigrarListaNegra()           → Migra la hoja LISTA NEGRA
'   - MigrarLog()                  → Migra la hoja LOG
'   - VerificarMigracion()         → Compara conteos hoja vs Access
' ==============================================================================

''' Procedimiento principal: migra TODAS las hojas de datos a Access.
''' Ejecutar desde el editor VBA con F5 o desde un botón temporal.
Public Sub MigrarTodasLasResidencias()
    Dim inicio As Double
    inicio = Timer
    
    If MsgBox("ATENCIÓN: Este proceso migrará todos los datos existentes " & _
              "de las hojas RESIDENCIA GIJÓN, SOTO y OVIEDO, " & _
              "LISTA NEGRA y LOG a la base de datos Access." & vbCrLf & vbCrLf & _
              "¿Has realizado una copia de seguridad del archivo Excel?" & vbCrLf & _
              "¿Deseas continuar?", _
              vbQuestion + vbYesNo, "Migración de Datos a Access") = vbNo Then
        Exit Sub
    End If
    
    ' Verificar conexión con Access
    If Not modDatabase.ComprobarConexion() Then
        MsgBox "No se puede conectar con la base de datos Access." & vbCrLf & _
               "Verifica que el archivo H:\ResidenciaBD\Residencia_BE.accdb existe " & _
               "y que tienes acceso a la unidad H:\.", _
               vbCritical, "Error de Conexión"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Migrando datos a Access..."
    
    Dim totalMigrados As Long
    totalMigrados = 0
    
    ' --- Migrar las 3 hojas de residencia ---
    On Error Resume Next
    
    Dim n As Long
    
    Application.StatusBar = "Migrando RESIDENCIA GIJÓN..."
    n = MigrarHojaResidencia("RESIDENCIA GIJÓN", "GIJON")
    If n >= 0 Then totalMigrados = totalMigrados + n
    
    Application.StatusBar = "Migrando RESIDENCIA SOTO..."
    n = MigrarHojaResidencia("RESIDENCIA SOTO", "SOTO")
    If n >= 0 Then totalMigrados = totalMigrados + n
    
    Application.StatusBar = "Migrando RESIDENCIA OVIEDO..."
    n = MigrarHojaResidencia("RESIDENCIA OVIEDO", "OVIEDO")
    If n >= 0 Then totalMigrados = totalMigrados + n
    
    ' --- Migrar Lista Negra ---
    Application.StatusBar = "Migrando LISTA NEGRA..."
    Dim nLN As Long
    nLN = MigrarListaNegra()
    
    ' --- Migrar LOGs ---
    Application.StatusBar = "Migrando LOGs..."
    Dim nLog As Long
    nLog = MigrarLogHoja("LOG_GIJÓN") + MigrarLogHoja("LOG_SOTO") + MigrarLogHoja("LOG_OVIEDO")
    
    On Error GoTo 0
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Dim duracion As String
    duracion = Format(Timer - inicio, "0.0")
    
    MsgBox "Migración completada en " & duracion & " segundos." & vbCrLf & vbCrLf & _
           "Registros migrados:" & vbCrLf & _
           "  • Órdenes/Solicitudes: " & totalMigrados & vbCrLf & _
           "  • Lista Negra: " & nLN & vbCrLf & _
           "  • Log: " & nLog & vbCrLf & vbCrLf & _
           "Ejecuta 'VerificarMigracion' para comprobar la integridad.", _
           vbInformation, "Migración Completada"
End Sub

''' Migra una hoja de residencia completa a la tabla Ordenes de Access.
''' Retorna el número de registros migrados, o -1 si hay error.
Public Function MigrarHojaResidencia( _
    ByVal nombreHoja As String, _
    ByVal codigoResidencia As String _
) As Long
    Dim ws As Worksheet
    Dim fila As Long
    Dim ultimaFila As Long
    Dim migrados As Long
    Dim errores As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja '" & nombreHoja & "'.", vbExclamation, "Hoja no encontrada"
        MigrarHojaResidencia = -1
        Exit Function
    End If
    
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    migrados = 0
    errores = 0
    
    ' Determinar columnas según residencia (Oviedo desplaza +1 por CAMA SUPLE. en Col 19)
    Dim colHabitaciones As Long, colTelefono As Long, colDireccion As Long
    Dim colCP As Long, colPoblacion As Long, colProvincia As Long, colEstadoPago As Long
    
    If UCase(codigoResidencia) = "OVIEDO" Then
        colHabitaciones = 20  ' Col T (NÚM HAB.)
        colTelefono = 23      ' Col W (TELEFONO)
        colDireccion = 24     ' Col X (DIRECCIÓN)
        colCP = 25            ' Col Y (CP)
        colPoblacion = 26     ' Col Z (POBLACIÓN)
        colProvincia = 27     ' Col AA (PROVINCIA)
        colEstadoPago = 28    ' Col AB (PAGADO)
    Else
        colHabitaciones = 19  ' Col S (NÚM HAB.)
        colTelefono = 22      ' Col V (TELEFONO)
        colDireccion = 23     ' Col W (DIRECCIÓN)
        colCP = 24            ' Col X (CP)
        colPoblacion = 25     ' Col Y (POBLACIÓN)
        colProvincia = 26     ' Col Z (PROVINCIA)
        colEstadoPago = 27    ' Col AA (PAGADO)
    End If
    
    For fila = 2 To ultimaFila
        ' Solo migrar filas que tengan un Nº Orden válido
        If IsNumeric(ws.Cells(fila, 1).Value) And ws.Cells(fila, 1).Value > 0 Then
            
            Dim sql As String
            Dim numOrden As Long
            numOrden = CLng(ws.Cells(fila, 1).Value)
            
            ' Construir INSERT con los valores de cada columna
            sql = "INSERT INTO Ordenes (" & _
                  "NumOrden, FechaPeticion, NumFactura, Residencia, " & _
                  "DNI, Nombre, " & _
                  "NumHabIndividuales, NumHabDobles, " & _
                  "FechaEntrada, FechaSalida, " & _
                  "Resolucion, HabitacionesAsignadas, " & _
                  "Telefono, Direccion, CodigoPostal, Poblacion, Provincia, " & _
                  "EstadoPago, FechaCreacion, UsuarioCreacion" & _
                  ") VALUES ("
            
            ' NumOrden (A)
            sql = sql & numOrden & ", "
            
            ' FechaPeticion (B)
            If IsDate(ws.Cells(fila, 2).Value) Then
                sql = sql & modDatabase.FormatearFechaSQL(CDate(ws.Cells(fila, 2).Value)) & ", "
            Else
                sql = sql & "NULL, "
            End If
            
            ' NumFactura (C) — puede estar vacío
            If IsNumeric(ws.Cells(fila, 3).Value) And ws.Cells(fila, 3).Value > 0 Then
                sql = sql & CLng(ws.Cells(fila, 3).Value) & ", "
            Else
                sql = sql & "NULL, "
            End If
            
            ' Residencia
            sql = sql & "'" & modDatabase.EscaparSQL(codigoResidencia) & "', "
            
            ' DNI (I = Col 9)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, 9).Value, ""))) & "', "
            
            ' Nombre (K = Col 11)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, 11).Value, ""))) & "', "
            
            ' NumHabIndividuales (Q = Col 17), NumHabDobles (R = Col 18)
            sql = sql & CLng(Nz(ws.Cells(fila, 17).Value, 0)) & ", "
            sql = sql & CLng(Nz(ws.Cells(fila, 18).Value, 0)) & ", "
            
            ' FechaEntrada (L = Col 12), FechaSalida (M = Col 13)
            If IsDate(ws.Cells(fila, 12).Value) Then
                sql = sql & modDatabase.FormatearFechaSQL(CDate(ws.Cells(fila, 12).Value)) & ", "
            Else
                sql = sql & "NULL, "
            End If
            
            If IsDate(ws.Cells(fila, 13).Value) Then
                sql = sql & modDatabase.FormatearFechaSQL(CDate(ws.Cells(fila, 13).Value)) & ", "
            Else
                sql = sql & "NULL, "
            End If
            
            ' Resolución (P = Col 16)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, 16).Value, ""))) & "', "
            
            ' HabitacionesAsignadas
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colHabitaciones).Value, ""))) & "', "
            
            ' Telefono, Direccion, CodigoPostal, Poblacion, Provincia
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colTelefono).Value, ""))) & "', "
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colDireccion).Value, ""))) & "', "
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colCP).Value, ""))) & "', "
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colPoblacion).Value, ""))) & "', "
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colProvincia).Value, ""))) & "', "
            
            ' EstadoPago
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colEstadoPago).Value, ""))) & "', "
            
            ' FechaCreacion, UsuarioCreacion
            sql = sql & modDatabase.FormatearFechaHoraSQL(Now) & ", "
            sql = sql & "'MIGRACION'"
            
            sql = sql & ")"
            
            On Error Resume Next
            Dim ok As Boolean
            ok = modDatabase.ExecuteNonQuery(sql)
            If ok Then
                migrados = migrados + 1
            Else
                errores = errores + 1
                Debug.Print "Error fila " & fila & " (NumOrden=" & numOrden & "): " & Err.Description
            End If
            On Error GoTo 0
            
            ' Actualizar barra de estado cada 50 registros
            If migrados Mod 50 = 0 Then
                Application.StatusBar = "Migrando " & nombreHoja & ": " & migrados & " registros..."
                DoEvents
            End If
        End If
    Next fila
    
    If errores > 0 Then
        Debug.Print "[AVISO] " & nombreHoja & ": " & errores & " errores durante la migración."
    End If
    
    MigrarHojaResidencia = migrados
End Function

''' Migra la hoja LISTA NEGRA a la tabla ListaNegra de Access.
Public Function MigrarListaNegra() As Long
    Dim ws As Worksheet
    Dim fila As Long, ultimaFila As Long, migrados As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LISTA NEGRA")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MigrarListaNegra = 0
        Exit Function
    End If
    
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    migrados = 0
    
    For fila = 2 To ultimaFila
        If Len(CStr(ws.Cells(fila, 1).Value)) > 0 Then
            Dim dniLN As String, nombreLN As String, motivoLN As String
            dniLN = CStr(ws.Cells(fila, 1).Value)
            nombreLN = CStr(Nz(ws.Cells(fila, 2).Value, ""))
            motivoLN = CStr(Nz(ws.Cells(fila, 3).Value, ""))
            
            On Error Resume Next
            If modDatabase.AnadirAListaNegra(dniLN, nombreLN, motivoLN) Then
                migrados = migrados + 1
            End If
            On Error GoTo 0
        End If
    Next fila
    
    MigrarListaNegra = migrados
End Function

''' Migra una hoja de LOG a la tabla LogActividad de Access.
Public Function MigrarLogHoja(ByVal nombreHoja As String) As Long
    Dim ws As Worksheet
    Dim fila As Long, ultimaFila As Long, migrados As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MigrarLogHoja = 0
        Exit Function
    End If
    
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    migrados = 0
    
    For fila = 2 To ultimaFila
        If Not IsEmpty(ws.Cells(fila, 1).Value) Then
            Dim usuario As String, accion As String, detalle As String
            usuario = CStr(Nz(ws.Cells(fila, 2).Value, ""))
            accion = CStr(Nz(ws.Cells(fila, 3).Value, ""))
            detalle = CStr(Nz(ws.Cells(fila, 4).Value, ""))
            
            On Error Resume Next
            If modDatabase.InsertarLog(usuario, accion, detalle) Then
                migrados = migrados + 1
            End If
            On Error GoTo 0
        End If
    Next fila
    
    MigrarLogHoja = migrados
End Function

Public Function MigrarLog() As Long
    MigrarLog = MigrarLogHoja("LOG_GIJÓN") + MigrarLogHoja("LOG_SOTO") + MigrarLogHoja("LOG_OVIEDO")
End Function

''' Verifica la integridad de la migración comparando conteos.
Public Sub VerificarMigracion()
    Dim msg As String
    msg = "=== VERIFICACIÓN DE MIGRACIÓN ===" & vbCrLf & vbCrLf
    
    ' Contar filas en cada hoja
    Dim wsGij As Long, wsSot As Long, wsOvi As Long
    On Error Resume Next
    wsGij = ThisWorkbook.Sheets("BDAS GIJÓN").Cells(Rows.Count, 1).End(xlUp).Row - 1
    wsSot = ThisWorkbook.Sheets("BDAS SOTO").Cells(Rows.Count, 1).End(xlUp).Row - 1
    wsOvi = ThisWorkbook.Sheets("BDAS OVIEDO").Cells(Rows.Count, 1).End(xlUp).Row - 1
    On Error GoTo 0
    
    ' Contar registros en Access
    Dim dbGij As Variant, dbSot As Variant, dbOvi As Variant, dbTotal As Variant
    dbGij = modDatabase.ExecuteScalar("SELECT COUNT(*) FROM Ordenes WHERE Residencia = 'GIJON'")
    dbSot = modDatabase.ExecuteScalar("SELECT COUNT(*) FROM Ordenes WHERE Residencia = 'SOTO'")
    dbOvi = modDatabase.ExecuteScalar("SELECT COUNT(*) FROM Ordenes WHERE Residencia = 'OVIEDO'")
    dbTotal = modDatabase.ExecuteScalar("SELECT COUNT(*) FROM Ordenes")
    
    msg = msg & "RESIDENCIA GIJÓN:" & vbCrLf
    msg = msg & "  Excel: " & wsGij & " filas | Access: " & Nz(dbGij, "?") & " registros"
    msg = msg & IIf(CLng(Nz(dbGij, 0)) = wsGij, " ✅", " ⚠️ DISCREPANCIA") & vbCrLf & vbCrLf
    
    msg = msg & "RESIDENCIA SOTO:" & vbCrLf
    msg = msg & "  Excel: " & wsSot & " filas | Access: " & Nz(dbSot, "?") & " registros"
    msg = msg & IIf(CLng(Nz(dbSot, 0)) = wsSot, " ✅", " ⚠️ DISCREPANCIA") & vbCrLf & vbCrLf
    
    msg = msg & "RESIDENCIA OVIEDO:" & vbCrLf
    msg = msg & "  Excel: " & wsOvi & " filas | Access: " & Nz(dbOvi, "?") & " registros"
    msg = msg & IIf(CLng(Nz(dbOvi, 0)) = wsOvi, " ✅", " ⚠️ DISCREPANCIA") & vbCrLf & vbCrLf
    
    msg = msg & "TOTAL en Access: " & Nz(dbTotal, "?") & " registros"
    
    Dim dbLN As Variant, dbLog As Variant
    dbLN = modDatabase.ExecuteScalar("SELECT COUNT(*) FROM ListaNegra")
    dbLog = modDatabase.ExecuteScalar("SELECT COUNT(*) FROM LogActividad")
    msg = msg & vbCrLf & "Lista Negra: " & Nz(dbLN, "?") & " registros"
    msg = msg & vbCrLf & "Log Actividad: " & Nz(dbLog, "?") & " registros"
    
    MsgBox msg, vbInformation, "Verificación de Migración"
End Sub

''' Función auxiliar para manejar valores vacíos/Null.
Private Function Nz(ByVal valor As Variant, Optional ByVal defecto As Variant = "") As Variant
    If IsNull(valor) Or IsEmpty(valor) Then
        Nz = defecto
    Else
        Nz = valor
    End If
End Function
