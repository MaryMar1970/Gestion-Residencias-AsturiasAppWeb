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
    
    ' --- Migrar LOG ---
    Application.StatusBar = "Migrando LOG..."
    Dim nLog As Long
    nLog = MigrarLog()
    
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
'''
''' IMPORTANTE: Ajustar el mapeo de columnas según la estructura real de cada hoja.
''' Las columnas indicadas son las documentadas en el análisis del BDAS:
'''   A=NumOrden, C=NumFactura, K=Nombre, L=FechaEntrada, M=FechaSalida,
'''   P=Resolución, Q=NumHabInd, R=NumHabDob, S=HabitacionesAsignadas (Gij/Soto),
'''   T=HabitacionesAsignadas (Oviedo), AA=EstadoPago (Gij/Soto), AB=EstadoPago (Oviedo)
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
    
    ' Determinar columnas de habitaciones y estado de pago según residencia
    Dim colHabitaciones As Long
    Dim colEstadoPago As Long
    If UCase(codigoResidencia) = "OVIEDO" Then
        colHabitaciones = 20  ' Col T
        colEstadoPago = 28    ' Col AB
    Else
        colHabitaciones = 19  ' Col S
        colEstadoPago = 27    ' Col AA
    End If
    
    For fila = 2 To ultimaFila
        ' Solo migrar filas que tengan un Nº Orden válido
        If IsNumeric(ws.Cells(fila, 1).Value) And ws.Cells(fila, 1).Value > 0 Then
            
            Dim sql As String
            Dim numOrden As Long
            numOrden = CLng(ws.Cells(fila, 1).Value)
            
            ' Construir INSERT con los valores de cada columna
            sql = "INSERT INTO Ordenes (" & _
                  "NumOrden, NumFactura, Residencia, " & _
                  "DNI, Nombre, Apellidos, " & _
                  "NumHabIndividuales, NumHabDobles, " & _
                  "FechaEntrada, FechaSalida, " & _
                  "Resolucion, HabitacionesAsignadas, EstadoPago, " & _
                  "Observaciones, FechaCreacion, UsuarioCreacion" & _
                  ") VALUES ("
            
            ' NumOrden (A)
            sql = sql & numOrden & ", "
            
            ' NumFactura (C) — puede estar vacío
            If IsNumeric(ws.Cells(fila, 3).Value) And ws.Cells(fila, 3).Value > 0 Then
                sql = sql & CLng(ws.Cells(fila, 3).Value) & ", "
            Else
                sql = sql & "NULL, "
            End If
            
            ' Residencia
            sql = sql & "'" & modDatabase.EscaparSQL(codigoResidencia) & "', "
            
            ' DNI (E aprox — AJUSTAR según columna real)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, 5).Value, ""))) & "', "
            
            ' Nombre (K)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, 11).Value, ""))) & "', "
            
            ' Apellidos (F aprox — AJUSTAR según columna real)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, 6).Value, ""))) & "', "
            
            ' NumHabIndividuales (Q), NumHabDobles (R)
            sql = sql & CLng(Nz(ws.Cells(fila, 17).Value, 0)) & ", "
            sql = sql & CLng(Nz(ws.Cells(fila, 18).Value, 0)) & ", "
            
            ' FechaEntrada (L), FechaSalida (M)
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
            
            ' Resolución (P)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, 16).Value, ""))) & "', "
            
            ' HabitacionesAsignadas (S o T según residencia)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colHabitaciones).Value, ""))) & "', "
            
            ' EstadoPago (AA o AB según residencia)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, colEstadoPago).Value, ""))) & "', "
            
            ' Observaciones (columna variable — AJUSTAR)
            sql = sql & "'" & modDatabase.EscaparSQL(CStr(Nz(ws.Cells(fila, 25).Value, ""))) & "', "
            
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

''' Migra la hoja LOG a la tabla LogActividad de Access.
Public Function MigrarLog() As Long
    Dim ws As Worksheet
    Dim fila As Long, ultimaFila As Long, migrados As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LOG")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MigrarLog = 0
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
    
    MigrarLog = migrados
End Function

''' Verifica la integridad de la migración comparando conteos.
Public Sub VerificarMigracion()
    Dim msg As String
    msg = "=== VERIFICACIÓN DE MIGRACIÓN ===" & vbCrLf & vbCrLf
    
    ' Contar filas en cada hoja
    Dim wsGij As Long, wsSot As Long, wsOvi As Long
    On Error Resume Next
    wsGij = ThisWorkbook.Sheets("RESIDENCIA GIJÓN").Cells(Rows.Count, 1).End(xlUp).Row - 1
    wsSot = ThisWorkbook.Sheets("RESIDENCIA SOTO").Cells(Rows.Count, 1).End(xlUp).Row - 1
    wsOvi = ThisWorkbook.Sheets("RESIDENCIA OVIEDO").Cells(Rows.Count, 1).End(xlUp).Row - 1
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
