Attribute VB_Name = "BusquedaOrdenNombreFactura"
'Attribute VB_Name = "BusquedaOrdenNombreFactura"

Option Explicit

' Variables a nivel de m�dulo
Private resultadosBusqueda As Collection
Private indiceActual As Long
Private columnaBusqueda As Long

Private Sub BusquedasResidencia(control As IRibbonControl)
    Dim shName As String
    shName = ActiveSheet.Name

    ' Reiniciar variables
    Set resultadosBusqueda = New Collection
    indiceActual = 0

    ' Verificar hoja activa
    If shName <> "RESIDENCIA GIJ�N" And shName <> "RESIDENCIA SOTO" And shName <> "RESIDENCIA OVIEDO" Then
        MsgBox "Esta macro solo se puede ejecutar en las hojas: RESIDENCIA GIJ�N, RESIDENCIA SOTO y RESIDENCIA OVIEDO", vbExclamation
        Exit Sub
    End If

    ' Mostrar UserForm de opciones
    frmBusqueda.searchOption = ""
    frmBusqueda.Show vbModal

    ' Capturar opci�n seleccionada
    Dim opcion As String
    opcion = frmBusqueda.searchOption
    Unload frmBusqueda

    If opcion = "" Then Exit Sub
    
    ' Asignar columna seg�n opci�n
    Dim colNum As Long
    Select Case opcion
    Case "N� ORDEN"
        colNum = 1
    Case "NOMBRE"
        colNum = 11
    Case "N� FACTURA"
        colNum = 3
    Case "N� DNI"
        colNum = 9
    Case Else
        MsgBox "Opci�n no v�lida.", vbExclamation
        Exit Sub
    End Select
    
    ' Solicitar valor de b�squeda
    Dim searchValue As String
    searchValue = InputBox("Ingrese el valor a buscar en la columna " & Col_Letter(colNum) & ":", "Valor a Buscar")
    If Trim(searchValue) = "" Then Exit Sub
    
    ' ==============================================================
    ' Buscar en hoja activa
    ' ==============================================================
    Dim foundCell As Range
    Dim firstAddress As String
    columnaBusqueda = colNum
    
    If colNum = 11 Then
        ' B�squeda parcial sin acentos para columna NOMBRE
        Dim processedSearchValue As String
        processedSearchValue = LCase(RemoveAccents(searchValue))
        
        Dim lastRow As Long
        lastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, colNum).End(xlUp).Row
        
        Dim cell As Range
        For Each cell In ActiveSheet.Range(ActiveSheet.Cells(1, colNum), ActiveSheet.Cells(lastRow, colNum))
            If Not IsEmpty(cell.Value) Then
                Dim cellValue As String
                cellValue = CStr(cell.Value)
                Dim processedCellValue As String
                processedCellValue = LCase(RemoveAccents(cellValue))
                If InStr(1, processedCellValue, processedSearchValue, vbBinaryCompare) > 0 Then
                    resultadosBusqueda.Add cell
                End If
            End If
        Next cell
    Else
        ' B�squeda est�ndar para otras columnas
        Dim lookAtParam As XlLookAt
        lookAtParam = IIf(colNum = 1 Or colNum = 3, xlWhole, xlPart)
        
        With ActiveSheet.Columns(colNum)
            Set foundCell = .Find(What:=searchValue, LookIn:=xlValues, LookAt:=lookAtParam, MatchCase:=False)
            If Not foundCell Is Nothing Then
                firstAddress = foundCell.Address
                Do
                    resultadosBusqueda.Add foundCell
                    Set foundCell = .FindNext(foundCell)
                Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
            End If
        End With
    End If
    
    ' ==============================================================
    ' BUSQUEDA EN ACCESS DATABASE (Fuente de verdad multiusuario)
    ' ==============================================================
    Dim critStr As String
    Select Case colNum
        Case 1: critStr = "NUMORDEN"
        Case 3: critStr = "NUMFACTURA"
        Case 9: critStr = "DNI"
        Case 11: critStr = "NOMBRE"
        Case Else: critStr = "NOMBRE"
    End Select
    
    Dim rsAccess As Object
    Set rsAccess = modDatabase.BuscarEnOrdenes(critStr, searchValue)
    If Not rsAccess Is Nothing Then
        If Not rsAccess.EOF Then
            ' Se encontraron resultados en Access DB
        End If
    End If

    ' ==============================================================
    ' NUEVA FUNCIONALIDAD: Buscar tambin en BDAS (si est activado)
    ' ==============================================================
    Dim buscarEnBDASActivado As Boolean
    buscarEnBDASActivado = BuscarBDASActivado() ' Funcin del mdulo ModuloArchivadoDatos
    
    Dim resultadosBDAS As Collection
    If buscarEnBDASActivado Then
        ' Determinar tipo de bsqueda para BDAS
        Dim tipoBusquedaBDAS As String
        Select Case colNum
            Case 1: tipoBusquedaBDAS = "ORDEN"
            Case 3: tipoBusquedaBDAS = "FACTURA"
            Case 9: tipoBusquedaBDAS = "DNI"
            Case 11: tipoBusquedaBDAS = "NOMBRE"
        End Select
        
        ' Buscar en BDAS
        Set resultadosBDAS = buscarEnBDAS(searchValue, tipoBusquedaBDAS)
    Else
        Set resultadosBDAS = New Collection
    End If
    
    ' ==============================================================
    ' Manejar resultados combinados
    ' ==============================================================
    Dim totalResultados As Long
    totalResultados = resultadosBusqueda.Count + resultadosBDAS.Count
    
    If totalResultados = 0 Then
        ' No se encontr� nada
        MsgBox "Valor no encontrado en la columna " & Col_Letter(colNum) & ".", vbInformation
        Exit Sub
    End If
    
    ' Si hay resultados en hoja activa, ir al primero
    If resultadosBusqueda.Count > 0 Then
        indiceActual = 1
        Application.enableEvents = False
        Application.GoTo resultadosBusqueda(indiceActual), True
        Application.enableEvents = True
        
        ' Mensaje con resumen
        Dim mensajeResumen As String
        mensajeResumen = "Encontrados " & resultadosBusqueda.Count & " resultado(s) en " & ActiveSheet.Name
        
        If resultadosBDAS.Count > 0 Then
            mensajeResumen = mensajeResumen & vbCrLf & vbCrLf & _
                           "Tambi�n hay " & resultadosBDAS.Count & " resultado(s) ARCHIVADO(S) en BDAS." & vbCrLf & _
                           "Use 'Ver Resultados BDAS' despu�s."
        End If
        
        If resultadosBusqueda.Count > 1 Then
            mensajeResumen = mensajeResumen & vbCrLf & vbCrLf & "�Buscar siguiente coincidencia en hoja activa?"
            If MsgBox(mensajeResumen, vbYesNo + vbQuestion, "Resultados de B�squeda") = vbYes Then
                BuscarSiguiente
            End If
        Else
            MsgBox mensajeResumen, vbInformation, "Resultados de B�squeda"
        End If
        
        ' Si hay resultados en BDAS, preguntar si desea verlos
        If resultadosBDAS.Count > 0 Then
            If MsgBox("�Desea ver los resultados archivados en BDAS?", vbYesNo + vbQuestion, "Ver BDAS") = vbYes Then
                Call MostrarResultadosBDAS(resultadosBDAS, opcion, searchValue)
            End If
        End If
        
    ElseIf resultadosBDAS.Count > 0 Then
        ' Solo hay resultados en BDAS
        MsgBox "No se encontraron datos en " & ActiveSheet.Name & vbCrLf & vbCrLf & _
               "Sin embargo, hay " & resultadosBDAS.Count & " resultado(s) ARCHIVADO(S)." & vbCrLf & vbCrLf & _
               "Se mostrar�n a continuaci�n.", vbInformation, "Datos Archivados"
        
        Call MostrarResultadosBDAS(resultadosBDAS, opcion, searchValue)
    End If
End Sub

' ==============================================================
' NUEVA FUNCI�N: Mostrar resultados de BDAS en ventana
' ==============================================================
Private Sub MostrarResultadosBDAS(resultados As Collection, tipoBusqueda As String, valorBuscado As String)
    On Error Resume Next
    
    If resultados.Count = 0 Then
        MsgBox "No hay resultados archivados.", vbInformation
        Exit Sub
    End If
    
    Dim mensaje As String
    mensaje = "=== RESULTADOS ARCHIVADOS EN BDAS ===" & vbCrLf & vbCrLf
    mensaje = mensaje & "B�squeda: " & tipoBusqueda & " = """ & valorBuscado & """" & vbCrLf
    mensaje = mensaje & "Total: " & resultados.Count & " registro(s)" & vbCrLf
    mensaje = mensaje & String(50, "-") & vbCrLf & vbCrLf
    
    Dim i As Long
    Dim resultado As Variant
    
    For i = 1 To resultados.Count
        resultado = resultados(i)
        
        mensaje = mensaje & "REGISTRO #" & i & vbCrLf
        mensaje = mensaje & "  N� ORDEN: " & resultado(1) & vbCrLf
        mensaje = mensaje & "  Fecha Petici�n: " & Format(resultado(2), "dd/mm/yyyy") & vbCrLf
        mensaje = mensaje & "  N� Factura: " & resultado(3) & vbCrLf
        mensaje = mensaje & "  DNI: " & resultado(4) & vbCrLf
        mensaje = mensaje & "  Nombre: " & resultado(5) & vbCrLf
        mensaje = mensaje & "  Tel�fono: " & resultado(6) & vbCrLf
        mensaje = mensaje & "  Direcci�n: " & resultado(7) & vbCrLf
        mensaje = mensaje & "  Grabaci�n: " & resultado(8) & vbCrLf
        mensaje = mensaje & "  Origen: " & resultado(9) & vbCrLf
        mensaje = mensaje & "  Archivado: " & Format(resultado(10), "dd/mm/yyyy hh:mm") & vbCrLf
        mensaje = mensaje & String(50, "-") & vbCrLf & vbCrLf
        
        ' Limitar a 5 resultados por ventana
        If i Mod 5 = 0 And i < resultados.Count Then
            mensaje = mensaje & "Mostrando " & i & " de " & resultados.Count & " resultados..."
            MsgBox mensaje, vbInformation, "Resultados BDAS (" & i & " de " & resultados.Count & ")"
            
            If MsgBox("�Continuar viendo m�s resultados?", vbYesNo + vbQuestion, "Continuar") = vbNo Then
                Exit For
            End If
            
            mensaje = "=== RESULTADOS ARCHIVADOS (continuaci�n) ===" & vbCrLf & vbCrLf
        End If
    Next i
    
    ' Mostrar �ltimo bloque
    If i <= resultados.Count Or resultados.Count <= 5 Then
        MsgBox mensaje, vbInformation, "Resultados BDAS"
    End If
    
    On Error GoTo 0
End Sub

Function RemoveAccents(text As String) As String
    Dim result As String
    result = text
    ' Reemplazar may�sculas acentuadas
    result = Replace(result, "�", "A")
    result = Replace(result, "�", "E")
    result = Replace(result, "�", "I")
    result = Replace(result, "�", "O")
    result = Replace(result, "�", "U")
    result = Replace(result, "�", "A")
    result = Replace(result, "�", "E")
    result = Replace(result, "�", "I")
    result = Replace(result, "�", "O")
    result = Replace(result, "�", "U")
    result = Replace(result, "�", "A")
    result = Replace(result, "�", "E")
    result = Replace(result, "�", "I")
    result = Replace(result, "�", "O")
    result = Replace(result, "�", "U")
    result = Replace(result, "�", "A")
    result = Replace(result, "�", "E")
    result = Replace(result, "�", "I")
    result = Replace(result, "�", "O")
    result = Replace(result, "�", "U")
    result = Replace(result, "�", "N")
    result = Replace(result, "�", "C")
    ' Reemplazar min�sculas acentuadas
    result = Replace(result, "�", "a")
    result = Replace(result, "�", "e")
    result = Replace(result, "�", "i")
    result = Replace(result, "�", "o")
    result = Replace(result, "�", "u")
    result = Replace(result, "�", "a")
    result = Replace(result, "�", "e")
    result = Replace(result, "�", "i")
    result = Replace(result, "�", "o")
    result = Replace(result, "�", "u")
    result = Replace(result, "�", "a")
    result = Replace(result, "�", "e")
    result = Replace(result, "�", "i")
    result = Replace(result, "�", "o")
    result = Replace(result, "�", "u")
    result = Replace(result, "�", "a")
    result = Replace(result, "�", "e")
    result = Replace(result, "�", "i")
    result = Replace(result, "�", "o")
    result = Replace(result, "�", "u")
    result = Replace(result, "�", "n")
    result = Replace(result, "�", "c")
    RemoveAccents = result
End Function

Sub BuscarSiguiente()
    On Error Resume Next
    indiceActual = indiceActual + 1
    
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        MsgBox "Se produjo un error al avanzar al siguiente resultado.", vbExclamation
        Exit Sub
    End If
    
    On Error GoTo 0
    
    If indiceActual > resultadosBusqueda.Count Then
        MsgBox "No hay m�s coincidencias en la hoja activa.", vbInformation
        Exit Sub
    End If
    
    Application.enableEvents = False
    Application.GoTo resultadosBusqueda(indiceActual), True
    Application.enableEvents = True
    
    If MsgBox("�Buscar siguiente coincidencia?", vbYesNo + vbQuestion, "Siguiente") = vbYes Then
        BuscarSiguiente
    End If
End Sub

Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

