Attribute VB_Name = "ModuloMantenimientoFC"
Option Explicit

'=================================================================================
' SUB: ConsolidarFormatosCondicionales
' OBJETIVO:
'   - Limpia la fragmentaci�n del formato condicional en la hoja especificada.
'   - Ajusta las reglas para que se apliquen desde la fila 2 hasta la 1048576.
'   - Elimina las reglas duplicadas sobrantes para liberar memoria y CPU.
'=================================================================================
Public Sub ConsolidarFormatosCondicionales(ws As Object)
    Dim wsReal As Object
    Dim fc As Object
    Dim i As Long
    Dim rng As Range
    Dim cleanRange As Range
    Dim key As String
    Dim dict As Object
    
    If ws Is Nothing Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    ' Obtener la hoja base
    Set wsReal = ws.Range("A1").Worksheet
    
    ' Si la hoja no tiene formato condicional en sus celdas, salimos (CORREGIDO)
    If wsReal.Cells.FormatConditions.Count = 0 Then Exit Sub
    
    ' Guardar estado de eventos y pantalla
    Dim screenUpdatingOriginal As Boolean
    Dim eventsOriginal As Boolean
    screenUpdatingOriginal = Application.screenUpdating
    eventsOriginal = Application.enableEvents
    
    Application.screenUpdating = False
    Application.enableEvents = False
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' PASO 1: Ajustar los rangos de todas las reglas a las columnas completas (fila 2 a 1048576)
    For i = 1 To wsReal.Cells.FormatConditions.Count
        Set fc = wsReal.Cells.FormatConditions(i)
        Set rng = fc.AppliesTo
        
        ' Obtener el rango limpio de las columnas completas desde la fila 2 (excluyendo cabecera)
        Set cleanRange = Intersect(rng.EntireColumn, wsReal.Rows("2:1048576"))
        
        ' Modificar el rango de aplicaci�n de la regla al rango limpio consolidado
        If Not cleanRange Is Nothing Then
            On Error Resume Next
            fc.ModifyAppliesToRange cleanRange
            On Error GoTo ErrorHandler
        End If
    Next i
    
    ' PASO 2: Eliminar reglas duplicadas (que ahora tienen la misma f�rmula y rango)
    ' Recorremos de atr�s hacia adelante para poder eliminar de forma segura
    For i = wsReal.Cells.FormatConditions.Count To 1 Step -1
        Set fc = wsReal.Cells.FormatConditions(i)
        
        ' Generar una clave �nica para identificar duplicados basados en:
        ' Tipo de regla + Rango de Aplicaci�n
        key = fc.Type & "|" & fc.AppliesTo.Address & "|"
        
        ' Intentar agregar la f�rmula (si aplica al tipo de regla)
        On Error Resume Next
        key = key & fc.Formula1
        On Error GoTo ErrorHandler
        
        If dict.Exists(key) Then
            ' Si ya existe una regla id�ntica, eliminamos este fragmento duplicado
            fc.Delete
        Else
            ' Si es la primera vez que la vemos, la guardamos en el diccionario
            dict(key) = True
        End If
    Next i
    
CleanExit:
    Set dict = Nothing
    Application.enableEvents = eventsOriginal
    Application.screenUpdating = screenUpdatingOriginal
    Exit Sub
    
ErrorHandler:
    ' En caso de error, siempre restaurar la configuraci�n del usuario
    Resume CleanExit
End Sub
