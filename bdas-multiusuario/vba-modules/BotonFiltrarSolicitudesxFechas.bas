Attribute VB_Name = "BotonFiltrarSolicitudesxFechas"
Option Explicit

'----------------------------------------------------------
' Alterna entre:
'  - Filtrar por intervalo de fechas (columna B)
'  - Restaurar y mostrar todas las filas
'
' Caracter�sticas / robustez a�adida:
'  - Guarda y restaura el estado original de Application.EnableEvents
'  - Maneja AutoFilter / ShowAllData de forma segura
'  - Desprotege y vuelve a proteger si la hoja estaba protegida
'  - Evita interferencias de Worksheet_SelectionChange (OVIEDO)
'    desactivando eventos justo antes de los Application.GoTo
'  - Manejo de errores centralizado que restablece estado de la app
'----------------------------------------------------------
Public Sub MostrarFilasPorIntervaloFechas(control As IRibbonControl)
    On Error GoTo Handler

    Dim ws As Worksheet
    Dim fechaInicio As Variant, fechaFin As Variant
    Dim lastRowB As Long
    Dim criterioInicio As String, criterioFin As String
    Dim hayFilasOcultas As Boolean
    Dim i As Long, filaPrimeraVisible As Long, filaUltimaVisible As Long
    Dim FILA_MAXIMA As Long
    Dim filasVisibles As Long
    Dim hojaEstabaProtegida As Boolean

    ' Guardar estados de la aplicaci�n para restaurarlos al final
    Dim prevEnableEvents As Boolean
    Dim prevScreenUpdating As Boolean
    prevEnableEvents = Application.enableEvents
    prevScreenUpdating = Application.screenUpdating

    ' S�lo actuar en las hojas permitidas
    Set ws = ActiveSheet
    Select Case ws.Name
        Case "RESIDENCIA GIJ�N", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO"
        Case Else
            Exit Sub
    End Select

    ' �ltima fila con dato en columna B (columna de fechas)
    lastRowB = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    FILA_MAXIMA = lastRowB
    If FILA_MAXIMA < 2 Then GoTo Salir   ' no hay datos �tiles

    ' Optimizaci�n y prevenci�n de eventos que interrumpan
    Application.screenUpdating = False
    If prevEnableEvents Then Application.enableEvents = False

    ' Desproteger hoja si est� protegida (sin contrase�a)
    hojaEstabaProtegida = ws.ProtectContents
    If hojaEstabaProtegida Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo Handler
    End If

    ' ---------------------------
    ' Detectar si ya hay filas ocultas / filtro aplicado
    ' ---------------------------
    hayFilasOcultas = False
    ' Si existe un filtro o est� en modo filtrado, lo consideramos
    If ws.AutoFilterMode Or ws.FilterMode Then hayFilasOcultas = True

    ' Si no se detect� por AutoFilter/FilterMode, inspeccionamos filas ocultas
    If Not hayFilasOcultas Then
        For i = 2 To FILA_MAXIMA
            If ws.Rows(i).Hidden Then
                hayFilasOcultas = True
                Exit For
            End If
        Next i
    End If

    If hayFilasOcultas Then
        ' -----------------------
        ' RESTAURAR: quitar filtro y mostrar todas las filas
        ' -----------------------
        On Error Resume Next
        ' Si hab�a un filtro aplicado, intentar mostrar todos los datos
        If ws.FilterMode Then ws.ShowAllData
        ' Si existe AutoFilter sin estar en modo filtrado, lo retiramos (comportamiento consistente)
        If ws.AutoFilterMode And Not ws.FilterMode Then ws.AutoFilterMode = False
        ' Asegurar que las filas est�n visibles
        ws.Rows("2:" & FILA_MAXIMA).Hidden = False
        On Error GoTo Handler

        ' Buscar la �ltima fila visible con dato en la columna B para situar el cursor
        filaUltimaVisible = 0
        For i = FILA_MAXIMA To 2 Step -1
            If Not ws.Rows(i).Hidden Then
                If Trim(CStr(ws.Cells(i, "B").Value)) <> "" Then
                    filaUltimaVisible = i
                    Exit For
                End If
            End If
        Next i

        If filaUltimaVisible <> 0 Then
            ' IMPORTANTE: desactivar eventos justo antes del GoTo para evitar que
            ' Worksheet_SelectionChange (OVIEDO) interfiera con la restauraci�n.
            ' Guardamos el estado actual de EnableEvents (ya lo tenemos en prevEnableEvents)
            ' y nos aseguramos de que est� FALSE durante el Application.Goto.
            Dim restoreEventsAfterGoto As Boolean
            restoreEventsAfterGoto = prevEnableEvents

            If Application.enableEvents Then Application.enableEvents = False
            Application.GoTo ws.Cells(filaUltimaVisible, "A"), True
            ' Restaurar al estado original
            If restoreEventsAfterGoto Then Application.enableEvents = True Else Application.enableEvents = False
        End If

    Else
        ' -----------------------
        ' FILTRAR: solicitar fechas y aplicar filtro sobre columna B
        ' -----------------------
        fechaInicio = Application.InputBox("Introduce la FECHA DE INICIO (dd/mm/yyyy):", "Intervalo de fechas", Type:=2)
        If fechaInicio = False Or Trim(CStr(fechaInicio)) = "" Then GoTo Salir

        fechaFin = Application.InputBox("Introduce la FECHA FINAL (dd/mm/yyyy):", "Intervalo de fechas", Type:=2)
        If fechaFin = False Or Trim(CStr(fechaFin)) = "" Then GoTo Salir

        If Not IsDate(fechaInicio) Or Not IsDate(fechaFin) Then
            MsgBox "Fechas no v�lidas.", vbCritical
            GoTo Salir
        End If

        If CDate(fechaInicio) > CDate(fechaFin) Then
            MsgBox "La fecha de inicio no puede ser posterior a la fecha final.", vbCritical
            GoTo Salir
        End If

        ' Criteria: usamos formato mm/dd/yyyy y Criteria2 con d�a siguiente para incluir el d�a final
        criterioInicio = Format(CDate(fechaInicio), "mm/dd/yyyy")
        criterioFin = Format(DateAdd("d", 1, CDate(fechaFin)), "mm/dd/yyyy")

        ' Asegurar que no haya restos de filtros anteriores
        On Error Resume Next
        If ws.FilterMode Then ws.ShowAllData
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        ws.Rows("2:" & FILA_MAXIMA).Hidden = False
        On Error GoTo Handler

        ' Aplicar filtro s�lo en el rango real usado (A:AB seg�n tu dise�o)
        ws.Range("A1:AB" & FILA_MAXIMA).AutoFilter Field:=2, _
            Criteria1:=">=" & criterioInicio, _
            Operator:=xlAnd, _
            Criteria2:="<" & criterioFin

        ' Contar filas visibles que tengan dato en B (tras aplicar el filtro)
        filasVisibles = 0
        filaPrimeraVisible = 0
        For i = 2 To FILA_MAXIMA
            If Not ws.Rows(i).Hidden Then
                If Trim(CStr(ws.Cells(i, "B").Value)) <> "" Then
                    filasVisibles = filasVisibles + 1
                    If filaPrimeraVisible = 0 Then filaPrimeraVisible = i
                End If
            End If
        Next i

        If filasVisibles = 0 Then
            MsgBox "No hay registros en el intervalo de fechas seleccionado.", vbInformation
            ' Restaurar todo por si acaso
            On Error Resume Next
            If ws.FilterMode Then ws.ShowAllData
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            ws.Rows("2:" & FILA_MAXIMA).Hidden = False
            On Error GoTo Handler
        Else
            ' Sit�a la vista en la primera fila visible con datos en B
            Dim restoreEventsAfterGoto2 As Boolean
            restoreEventsAfterGoto2 = prevEnableEvents

            If Application.enableEvents Then Application.enableEvents = False
            Application.GoTo ws.Cells(filaPrimeraVisible, "A"), True
            If restoreEventsAfterGoto2 Then Application.enableEvents = True Else Application.enableEvents = False
        End If
    End If

    ' -----------------------
    ' Volver a proteger la hoja si estaba protegida
    ' -----------------------
    If hojaEstabaProtegida Then
        On Error Resume Next
        ws.Protect AllowFormattingRows:=True
        On Error GoTo Handler
    End If

Salir:
    ' Restaurar estados de la aplicaci�n
    If prevEnableEvents Then Application.enableEvents = True Else Application.enableEvents = False
    Application.screenUpdating = prevScreenUpdating
    Exit Sub

Handler:
    ' Manejo centralizado de errores: intentar restaurar estado y notificar
    On Error Resume Next
    If hojaEstabaProtegida Then ws.Protect AllowFormattingRows:=True
    If prevEnableEvents Then Application.enableEvents = True Else Application.enableEvents = False
    Application.screenUpdating = prevScreenUpdating
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "MostrarFilasPorIntervaloFechas"
    Resume Salir
End Sub


