Attribute VB_Name = "FechasPeticion"
'=================================================================================
' M�dulo: ModuloFechas
'
' Descripci�n:
'   Valida y formatea fechas en las columnas B (fecha petici�n), L (fecha entrada) y M (fecha salida).
'   Aplica formatos correctos, verifica rangos permitidos y congruencia entre fechas,
'   y calcula la estancia en d�as (columna N). Muestra avisos en caso de errores.
'   Incluye manejo de errores para asegurar la robustez de la macro.
'=================================================================================

Option Explicit

'---------------------------------------------------------------------------------
' Sub: ValidacionYFormatoFechas
'
' Par�metros:
'   - Target: Rango de celdas modificadas por el usuario (Worksheet_Change)
'   - Hoja:   Hoja donde ocurre el cambio
'
' L�gica principal:
'   1. Solo opera si el cambio afecta a las columnas B, L o M.
'   2. Para cada celda editada:
'        a. Establece formato y verifica el valor de fecha/hora.
'        b. Aplica validaciones seg�n el tipo de fecha (petici�n, entrada, salida).
'        c. Calcula y actualiza la estancia en d�as en columna N.
'   3. Muestra mensajes claros ante errores o valores incoherentes.
'   4. Siempre reactiva los eventos al finalizar (aunque haya errores).
'---------------------------------------------------------------------------------
Public Sub ValidacionYFormatoFechas(ByVal Target As Range, ByVal hoja As Worksheet)
    On Error GoTo ErrHandler

    '----------------------- Declaraci�n de variables -----------------------------
    Dim celda As Range
    Dim fechaActual   As Date      ' Fecha/hora del sistema (referencia para validaciones)
    Dim fechaPeticion As Date      ' Fecha de la solicitud (columna B)
    Dim fechaEntrada  As Date      ' Fecha de entrada (columna L)
    Dim fechaSalida   As Date      ' Fecha de salida (columna M)
    Dim partesTiempo  As Variant   ' Partes de la fecha/hora para columna B
    Dim fechaCruda    As String    ' Texto original de la celda editada
    Dim fechaFinal    As Date      ' Fecha final tras formatear/convertir
    Dim fechaParte    As String    ' Parte de fecha (sin la hora)
    Dim horaParte     As String    ' Parte de hora
    Dim celdaDias     As Range     ' Celda de columna N (d�as de estancia)
    Dim celdaEntrada  As Range     ' Celda de entrada (L)
    Dim celdaSalida   As Range     ' Celda de salida (M)
    Dim rngCols       As Range     ' Rango de columnas relevantes (B, L, M)
    Dim limiteDias    As Long      ' L�mite m�ximo de estancia seg�n hoja
    Dim hojaProtegida As Boolean   ' Indica si la hoja est� protegida
    Dim password      As String    ' Contrase�a de desprotecci�n (si existe)

    '----------------- 1. Obtener fecha/hora actual del sistema -------------------
    fechaActual = Now

    '----------------- 2. Definir el rango de inter�s: columnas B, L y M ---------
    Set rngCols = Union(hoja.Columns("B"), hoja.Columns("L"), hoja.Columns("M"))

    ' Si el cambio no afecta a esas columnas, salir del procedimiento.
    If Intersect(Target, rngCols) Is Nothing Then Exit Sub

    ' Verificar si la hoja est� protegida
    hojaProtegida = hoja.ProtectContents
    If hojaProtegida Then
        ' Intentar desproteger (ajusta la contrase�a si es necesaria)
        On Error Resume Next
        hoja.Unprotect password:=""
        If Err.Number <> 0 Then
            MsgBox "No se puede modificar las celdas porque la hoja est� protegida.", vbExclamation, "Hoja Protegida"
            Exit Sub
        End If
        On Error GoTo ErrHandler
    End If

    ' Desactivar eventos para evitar bucles.
    Application.enableEvents = False

    '----------------- 3. Recorrer todas las celdas cambiadas relevantes ---------
    For Each celda In Intersect(Target, rngCols)
    If celda.Row < 2 Then GoTo Siguiente
        ' Solo procesar si la celda no est� vac�a.
        If Trim(celda.Value & "") <> "" Then

            ' ------------------ PARTE 1: Formato y validaci�n b�sica --------------
            fechaCruda = CStr(celda.Value)

            Select Case celda.Column
'------------------ Columna B: FECHA PETICI�N (fecha + hora) ----------
    Case 2
        ' Normalizar separadores y extraer partes de fecha/hora
        fechaCruda = Replace(Replace(fechaCruda, "-", "/"), ".", "/")
        partesTiempo = Split(fechaCruda, " ")
        fechaParte = partesTiempo(0)
    
        If UBound(partesTiempo) > 0 Then
            horaParte = partesTiempo(1)
        
        '--- NUEVA L�GICA PARA FORMATOS hhmm y hmm ---
        If InStr(horaParte, ":") = 0 And Len(horaParte) >= 3 Then
            ' Intentar parsear como hhmm o hmm
            Dim horaFormateada As String
            Select Case Len(horaParte)
                Case 3 ' Formato hmm
                    horaFormateada = Left(horaParte, 1) & ":" & Right(horaParte, 2)
                Case 4 ' Formato hhmm
                    horaFormateada = Left(horaParte, 2) & ":" & Right(horaParte, 2)
                Case Else
                    horaFormateada = horaParte
            End Select
            
            ' Validar que el formato sea correcto
            If IsDate("00:00") Then ' Peque�a comprobaci�n de formato de hora
                horaParte = horaFormateada
            End If
        End If
                Else
                    horaParte = "00:00"
                End If
                ' Si solo ponen d�a/mes, a�adir el a�o actual
                If Len(fechaParte) < 6 Then
                    fechaParte = fechaParte & "/" & Year(Date)
                End If
                ' Validar formato de fecha/hora
                If IsDate(fechaParte & " " & horaParte) Then
                    fechaFinal = CDate(fechaParte & " " & horaParte)
                    celda.Value = fechaFinal
                    ' Aplicar formato con manejo de errores
                    On Error Resume Next
                    celda.NumberFormat = "dd/mm/yyyy hh:mm"
                    If Err.Number <> 0 Then
                        ' Intentar formato alternativo si falla el primero
                        Err.Clear
                        celda.NumberFormat = "m/d/yyyy h:mm"
                    End If
                    On Error GoTo ErrHandler
                Else
                    MsgBox "Formato FECHA PETICION inv�lido. Use dd/mm hh:mm o dd-mm hh:mm", vbCritical, "Error"
                    celda.ClearContents
                    celda.Select
                    GoTo Siguiente
                End If

            '--------------- Columnas L y M: ENTRADA y SALIDA (solo fecha) --------
            Case 12, 13
                fechaCruda = Replace(fechaCruda, "-", "/")
                If Len(fechaCruda) < 6 Then fechaCruda = fechaCruda & "/" & Year(Date)
                If IsDate(fechaCruda) Then
                    celda.Value = CDate(fechaCruda)
                    ' Aplicar formato con manejo de errores
                    On Error Resume Next
                    celda.NumberFormat = "dd/mm/yyyy"
                    If Err.Number <> 0 Then
                        ' Intentar formato alternativo si falla el primero
                        Err.Clear
                        celda.NumberFormat = "m/d/yyyy"
                    End If
                    On Error GoTo ErrHandler
                Else
                    If celda.Column = 12 Then
                        MsgBox "FECHA ENTRADA no v�lida. Use formato dd/mm o dd-mm", vbCritical, "Error"
                    Else
                        MsgBox "FECHA SALIDA no v�lida. Use formato dd/mm o dd-mm", vbCritical, "Error"
                    End If
                    celda.ClearContents
                    celda.Select
                    GoTo Siguiente
                End If
            End Select

            ' ------------------ PARTE 2: Validaciones de congruencia --------------
            Select Case celda.Column

            '----------- Validaci�n FECHA PETICI�N: No futura y no mayor a 7 d�as ---
            Case 2
                fechaPeticion = CDate(celda.Value)
                If fechaPeticion > fechaActual Then
                    MsgBox "FECHA PETICI�N no puede ser posterior a FECHA ACTUAL.", vbExclamation, "Fecha Inv�lida"
                    celda.ClearContents
                    celda.Select
                    GoTo Siguiente
                End If
                If fechaPeticion < DateAdd("d", -7, fechaActual) Then
                    MsgBox "FECHA PETICI�N hace m�s de 7 d�as. Verificar si es correcto.", vbExclamation, "Fecha Inv�lida"
                    celda.Select
                    GoTo Siguiente
                End If

            '---------- Validaci�n FECHA ENTRADA: Rango y congruencia con FECHA PETICI�N --
            Case 12
                fechaEntrada = CDate(celda.Value)
                If Not IsDate(hoja.Cells(celda.Row, "B").Value) Then GoTo Siguiente
                fechaPeticion = CDate(hoja.Cells(celda.Row, "B").Value)

                If fechaEntrada > DateAdd("d", 30, fechaPeticion) Then
                    MsgBox "FECHA ENTRADA supera +30 d�as naturales la FECHA PETICI�N." & vbCrLf & _
                           "La solicitud quedara marcada como DESESTIMADA.", vbExclamation, "Aviso"
                    ' Guard clause: no sobreescribir si ya hay resolucion definitiva
                    Dim valorActualP As String
                    valorActualP = UCase(Trim(hoja.Cells(celda.Row, "P").Value))
                    If valorActualP <> "SI" And valorActualP <> "CONCEDIDA" And _
                       valorActualP <> "REEVALUADA" And valorActualP <> "RENUNCIA" Then
                        hoja.Cells(celda.Row, "P").Value = "DESESTIMADA"
                    End If
                    hoja.Cells(celda.Row, "M").Select
                    GoTo Siguiente
                End If
                If DateValue(fechaEntrada) < DateValue(fechaPeticion) Then
                    MsgBox "FECHA ENTRADA debe ser mayor o igual a FECHA PETICI�N.", vbCritical, "Error de Fecha"
                    celda.Select
                    GoTo Siguiente
                End If

                ' Si ya existe FECHA SALIDA, comprobar congruencia (SALIDA > ENTRADA)
                If IsDate(hoja.Cells(celda.Row, "M").Value) Then
                    fechaSalida = CDate(hoja.Cells(celda.Row, "M").Value)
                    If DateValue(fechaSalida) <= DateValue(fechaEntrada) Then
                        MsgBox "FECHA SALIDA debe ser posterior a FECHA ENTRADA.", vbCritical, "Error"
                        hoja.Cells(celda.Row, "M").ClearContents
                        hoja.Cells(celda.Row, "M").Select
                        GoTo Siguiente
                    End If
                End If

            '---------- Validaci�n FECHA SALIDA: Requiere ENTRADA y congruencia ------
            Case 13
                fechaSalida = CDate(celda.Value)
                If Not IsDate(hoja.Cells(celda.Row, "L").Value) Then
                    MsgBox "FECHA ENTRADA no existe, introduzca una fecha v�lida.", vbExclamation, "Error"
                    hoja.Cells(celda.Row, "L").Select
                    GoTo ActualizarN
                End If
                fechaEntrada = CDate(hoja.Cells(celda.Row, "L").Value)

                ' Determinar l�mite de estancia seg�n hoja
                Select Case hoja.Name
                Case "RESIDENCIA GIJ�N", "RESIDENCIA OVIEDO"
                    limiteDias = 7
                Case "RESIDENCIA SOTO"
                    limiteDias = 10
                Case Else
                    limiteDias = 7
                End Select

                ' Validaci�n de estancia prolongada
                If DateDiff("d", fechaEntrada, fechaSalida) > limiteDias Then
                    MsgBox "Atenci�n: FECHA SALIDA supera +" & limiteDias & " d�as la FECHA ENTRADA. Verificar validez.", _
                           vbExclamation, "Aviso de Estancia Prolongada"
                End If

                ' Validaci�n: SALIDA debe ser posterior a ENTRADA
                If DateValue(fechaSalida) <= DateValue(fechaEntrada) Then
                    MsgBox "FECHA SALIDA debe ser posterior a FECHA ENTRADA.", vbCritical, "Error"
                    celda.ClearContents
                    celda.Select
                    GoTo Siguiente
                End If
            End Select
        End If

ActualizarN:
        ' ------------------ PARTE 3: Actualiza columna N (das de estancia) y Sync Access -------
        If celda.Column = 12 Or celda.Column = 13 Then
            Set celdaEntrada = hoja.Cells(celda.Row, "L")
            Set celdaSalida = hoja.Cells(celda.Row, "M")
            Set celdaDias = hoja.Cells(celda.Row, "N")

            If IsDate(celdaEntrada.Value) And IsDate(celdaSalida.Value) Then
                celdaDias.Value = celdaSalida.Value - celdaEntrada.Value
                On Error Resume Next
                celdaDias.NumberFormat = "0"
                On Error GoTo ErrHandler
            Else
                celdaDias.Value = ""
            End If
        End If

        ' Sincronizar campo modificado con Access si existe NumOrden (Col A)
        If IsNumeric(hoja.Cells(celda.Row, 1).Value) And hoja.Cells(celda.Row, 1).Value > 0 Then
            Dim nOrd As Long
            nOrd = CLng(hoja.Cells(celda.Row, 1).Value)
            Select Case celda.Column
                Case 2:  modDatabase.ActualizarOrden nOrd, "FechaPeticion", CStr(celda.Value)
                Case 12: modDatabase.ActualizarOrden nOrd, "FechaEntrada", CStr(celda.Value)
                Case 13: modDatabase.ActualizarOrden nOrd, "FechaSalida", CStr(celda.Value)
            End Select
        End If

Siguiente:
    Next celda

CleanUp:
    ' Reactivar eventos siempre antes de salir.
    Application.enableEvents = True
    
    ' Reproteger la hoja si estaba protegida
    If hojaProtegida Then
        On Error Resume Next
        hoja.Protect password:=""
        On Error GoTo 0
    End If
    
    Exit Sub

'------------------------ Manejo de errores inesperados --------------------------
ErrHandler:
    ' Asegurarse de reactivar eventos y reproteger la hoja incluso en caso de error
    Application.enableEvents = True
    If hojaProtegida Then
        On Error Resume Next
        hoja.Protect password:=""
        On Error GoTo 0
    End If
    
    MsgBox "Error en validaci�n/formato de fechas: " & Err.Description, vbCritical, "�Oops!"
End Sub

