Attribute VB_Name = "CalendarioRetroceder5dias"
'==============================================================================
' M�DULO: CalendarioRetroceder5d�as
' DESCRIPCI�N: Navegaci�n de calendario - Retrocede 5 d�as desde la fecha m�s cercana a la columna actual
' NOTA: Las fechas est�n en columnas 2, 5, 8, 11... (patr�n de 3 columnas)
' AUTOR: Bustiello
' FECHA: 2025-10-12
'==============================================================================

Option Explicit

Public Sub Retroceder5DiasDesdeFechaVisible(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim targetDate As Date
    Dim cell As Range
    Dim found As Boolean
    Dim referenciaFecha As Date
    Dim columnActual As Long
    Dim columnaFecha As Long
    
    ' Obtener hoja activa
    Set ws = ActiveSheet
    
    ' =========================================================================
    ' VALIDACI�N: Solo ejecutar en hojas de calendario
    ' =========================================================================
    If ws.Name <> "Calendario OVIEDO" And _
       ws.Name <> "Calendario GIJ�N" And _
       ws.Name <> "Calendario SOTO" Then
        Exit Sub
    End If
    
    ' =========================================================================
    ' DETERMINAR COLUMNA CON FECHA M�S CERCANA
    ' =========================================================================
    columnActual = ActiveCell.Column
    
    ' Buscar la columna con fecha m�s cercana (patr�n: 2, 5, 8, 11, 14...)
    columnaFecha = ObtenerColumnaFechaMasCercana(columnActual)
    
    ' Leer la fecha de esa columna
    On Error Resume Next
    If IsDate(ws.Cells(1, columnaFecha).Value) Then
        referenciaFecha = ws.Cells(1, columnaFecha).Value
    Else
        referenciaFecha = 0
    End If
    On Error GoTo ErrorHandler
    
    ' Validar que se obtuvo una fecha v�lida
    If referenciaFecha = 0 Or Not IsDate(referenciaFecha) Or Year(referenciaFecha) < 2000 Then
        MsgBox "No se pudo determinar la fecha de referencia." & vbCrLf & vbCrLf & _
               "Columna del cursor: " & ColumnNumberToLetter(columnActual) & vbCrLf & _
               "Columna con fecha m�s cercana: " & ColumnNumberToLetter(columnaFecha) & vbCrLf & _
               "Valor encontrado: " & ws.Cells(1, columnaFecha).Value & vbCrLf & vbCrLf & _
               "Por favor, aseg�rate de estar en el �rea del calendario.", _
               vbExclamation, "Fecha no v�lida"
        Exit Sub
    End If
    
    ' =========================================================================
    ' CALCULAR FECHA OBJETIVO (restar 5 d�as)
    ' =========================================================================
    targetDate = referenciaFecha - 5
    
    ' =========================================================================
    ' BUSCAR LA FECHA OBJETIVO EN LAS COLUMNAS CON FECHAS (patr�n 2, 5, 8...)
    ' =========================================================================
    found = BuscarYNavegarAFecha(ws, targetDate)
    
    ' =========================================================================
    ' MENSAJE SI NO SE ENCONTR� LA FECHA
    ' =========================================================================
    If Not found Then
        MsgBox "La fecha " & Format(targetDate, "dd/mm/yyyy") & " no se encuentra en el calendario.", _
               vbInformation, "Fecha no encontrada"
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "Error al retroceder 5 d�as:" & vbCrLf & vbCrLf & _
           "N�mero: " & Err.Number & vbCrLf & _
           "Descripci�n: " & Err.Description, _
           vbCritical, "Error de navegaci�n"
End Sub

'==============================================================================
' FUNCI�N: ObtenerColumnaFechaMasCercana
' DESCRIPCI�N: Calcula la columna con fecha m�s cercana seg�n el patr�n 2,5,8,11...
' PAR�METROS: colActual (Long) - Columna donde est� el cursor
' RETORNA: Long - Columna con fecha (2, 5, 8, 11, 14...)
'==============================================================================
Private Function ObtenerColumnaFechaMasCercana(colActual As Long) As Long
    ' Patr�n: columnas con fecha = 2 + (n � 3) donde n = 0, 1, 2, 3...
    ' Columnas: 2, 5, 8, 11, 14, 17, 20...
    
    If colActual < 2 Then
        ' Si est� en columna A, usar la primera fecha (columna B = 2)
        ObtenerColumnaFechaMasCercana = 2
    Else
        ' Calcular la columna de fecha m�s cercana
        Dim diferencia As Long
        diferencia = (colActual - 2) Mod 3
        
        Select Case diferencia
            Case 0  ' Ya est� en columna de fecha (2, 5, 8, 11...)
                ObtenerColumnaFechaMasCercana = colActual
            Case 1  ' Est� 1 columna despu�s de la fecha (3, 6, 9...)
                ObtenerColumnaFechaMasCercana = colActual - 1
            Case 2  ' Est� 2 columnas despu�s de la fecha (4, 7, 10...)
                ObtenerColumnaFechaMasCercana = colActual - 2
        End Select
    End If
End Function

'==============================================================================
' FUNCI�N: BuscarYNavegarAFecha
' DESCRIPCI�N: Busca una fecha en las columnas del patr�n y navega a ella
' PAR�METROS: ws (Worksheet), fechaBuscada (Date)
' RETORNA: Boolean - True si encontr� y naveg�, False si no encontr�
'==============================================================================
Private Function BuscarYNavegarAFecha(ws As Worksheet, fechaBuscada As Date) As Boolean
    On Error Resume Next
    
    Dim col As Long
    Dim valorCelda As Variant
    
    BuscarYNavegarAFecha = False
    
    ' Recorrer las columnas con fecha seg�n el patr�n: 2, 5, 8, 11, 14...
    For col = 2 To 1400 Step 3  ' Hasta columna ~APZ
        valorCelda = ws.Cells(1, col).Value
        
        If IsDate(valorCelda) Then
            If Year(valorCelda) >= 2000 And Year(valorCelda) <= 2100 Then
                If CDate(valorCelda) = fechaBuscada Then
                    ' Fecha encontrada: navegar a esa celda
                    Application.GoTo Reference:=ws.Cells(1, col), Scroll:=True
                    BuscarYNavegarAFecha = True
                    Exit Function
                End If
            End If
        End If
    Next col
    
End Function

'==============================================================================
' FUNCI�N: ColumnNumberToLetter
' DESCRIPCI�N: Convierte n�mero de columna a letra (1?A, 2?B, etc.)
'==============================================================================
Private Function ColumnNumberToLetter(colNum As Long) As String
    On Error Resume Next
    ColumnNumberToLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

