Attribute VB_Name = "BorradoDatosAlojamiento"
'=================================================================================
' Macro BorrarDatosAlojamientoOptimizado
'---------------------------------------------------------------------------------
' Borra de manera masiva y segura todos los datos de las hojas de alojamiento
' (RESIDENCIA GIJ�N, RESIDENCIA SOTO, RESIDENCIA OVIEDO), excepto aquellas columnas
' protegidas seg�n cada centro.
'
' Optimizaci�n:
'  - Calcula la �ltima fila realmente usada en toda la hoja (no solo columna A).
'  - Fuerza recalculado del UsedRange para evitar filas "fantasma" en el conteo.
'  - Desprotege la hoja si es necesario.
'  - Borra comentarios en columna P.
'  - Restaura siempre el estado de Excel ante cualquier error.
'  - Informa al usuario del n�mero real de filas procesadas y estado de protecci�n.
'
' Uso t�pico: asignar a un bot�n de la Cinta de Opciones o macro manual.
'=================================================================================
Option Explicit

Public Sub BorrarDatosAlojamientoOptimizado(control As IRibbonControl)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- Confirmaciones ---
    If MsgBox("Se va a proceder al borrado de la totalidad de los datos de la hoja '" & ws.Name & "'." & vbCrLf & vbCrLf & _
              "�Desea continuar?", vbExclamation + vbYesNo, "Confirmaci�n Inicial") <> vbYes Then Exit Sub

    If MsgBox("Realmente desea borrar la totalidad de los datos de la hoja '" & ws.Name & "'?" & vbCrLf & vbCrLf & _
              "El proceso es IRREVERSIBLE." & vbCrLf & vbCrLf & "�Confirma el borrado completo?", _
              vbCritical + vbYesNo, "Confirmaci�n Final - �ATENCI�N!") <> vbYes Then Exit Sub

    ' --- Columnas protegidas por hoja ---
    Dim protectedCols As Variant
    Select Case ws.Name
        Case "RESIDENCIA GIJ�N":   protectedCols = Array("U", "AB", "AC")
        Case "RESIDENCIA SOTO":    protectedCols = Array("Q", "R", "U", "AB", "AC")
        Case "RESIDENCIA OVIEDO":  protectedCols = Array("V", "AF", "AG")
        Case Else
            MsgBox "Esta macro solo puede ejecutarse en las hojas de alojamiento", vbCritical, "Error"
            Exit Sub
    End Select

    ' --- Optimizaci�n de Excel ---
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual

    ' --- Forzar recalculado del UsedRange para evitar filas "fantasma" ---
    Dim rangoUsado As Range
    Set rangoUsado = ws.UsedRange

    ' --- C�lculo robusto de la �ltima fila usada en cualquier columna �til ---
    Dim lastCol As Long: lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim lastRow As Long, i As Long, tempRow As Long
    lastRow = 2 ' Siempre deja la primera fila (encabezados)
    For i = 1 To lastCol
        tempRow = ws.Cells(ws.Rows.Count, i).End(xlUp).Row
        If tempRow > lastRow Then lastRow = tempRow
    Next i

    ' --- Si no hay datos, termina ---
    If lastRow < 2 Then lastRow = 2

    ' --- Desproteger hoja si estaba protegida ---
    If ws.ProtectContents Then
        On Error Resume Next
        ws.Unprotect ' Sin contrase�a
        On Error GoTo ErrHandler
    End If

    ' --- DisplayAlerts JUSTO antes de borrar (evita avisos de celdas bloqueadas) ---
    Application.DisplayAlerts = False

    ' --- Borrado masivo: borra solo columnas NO protegidas ---
    Dim col As Long
    For col = 1 To lastCol
        Dim colLetter As String
        colLetter = Split(ws.Cells(1, col).Address(True, False), "$")(0)

        Dim isProtected As Boolean
        isProtected = Not IsError(Application.Match(colLetter, protectedCols, 0))

        If Not isProtected Then
            ws.Range(ws.Cells(2, col), ws.Cells(lastRow, col)).ClearContents
        End If
    Next col

    ' --- Limpieza de comentarios columna P ---
    If lastRow >= 2 Then
        ws.Range("P2:P" & lastRow).ClearComments
    End If

    ' --- Restauraci�n del entorno Excel ---
    Application.enableEvents = True
    Application.calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.screenUpdating = True

    ' --- Mensaje de resumen ---
    MsgBox "Filas procesadas: " & (lastRow - 1) & vbCrLf & _
           "Estado protecci�n: " & IIf(ws.ProtectContents, "PROTEGIDA", "DESPROTEGIDA"), _
           vbInformation, "Limpieza Finalizada - " & ws.Name
    Exit Sub

ErrHandler:
    ' --- Restauraci�n segura ante error ---
    Application.DisplayAlerts = True
    Application.enableEvents = True
    Application.calculation = xlCalculationAutomatic
    Application.screenUpdating = True

    MsgBox "Error durante el borrado:" & vbCrLf & Err.Description, vbCritical, "Error en " & ws.Name
End Sub

