Attribute VB_Name = "BorradoSolicitudes"
'Attribute VB_Name = "BorradoSolicitudes"
'-------------------------------------------------------------------------------------------
' M�dulo:  BorradoSolicitudes
'-------------------------------------------------------------------------------------------
' Descripci�n:
'   - Permite eliminar filas completas de las hojas RESIDENCIA GIJ�N, RESIDENCIA SOTO
'     y RESIDENCIA OVIEDO bas�ndose en el N� ORDEN (columna A).
'   - Soporta selecci�n individual (varios N� ORDEN separados por coma) o por rango.
'   - Registra los borrados en el LOG correspondiente antes de eliminar.
'   - Elimina las filas de abajo hacia arriba para no alterar �ndices.
'
' Uso:
'   1) Ejecuta la macro BorrarSolicitudesPorOrden
'   2) Selecciona la(s) hoja(s) donde borrar
'   3) Elige el modo de selecci�n (individual o rango)
'   4) Introduce los N� ORDEN a eliminar
'   5) Confirma el borrado
'-------------------------------------------------------------------------------------------
Option Explicit

' Constantes para las hojas y sus LOG
Private Const HOJA_GIJON As String = "RESIDENCIA GIJ�N"
Private Const HOJA_SOTO As String = "RESIDENCIA SOTO"
Private Const HOJA_OVIEDO As String = "RESIDENCIA OVIEDO"
Private Const LOG_GIJON As String = "LOG_GIJ�N"
Private Const LOG_SOTO As String = "LOG_SOTO"
Private Const LOG_OVIEDO As String = "LOG_OVIEDO"


'-------------------------------------------------------------------------------------------
' Funci�n:  ObtenerNombreLOG
' Devuelve el nombre de la hoja LOG correspondiente a una hoja de residencia
'-------------------------------------------------------------------------------------------
Private Function ObtenerNombreLOG(ByVal nombreHoja As String) As String
    Select Case nombreHoja
        Case HOJA_GIJON
            ObtenerNombreLOG = LOG_GIJON
        Case HOJA_SOTO
            ObtenerNombreLOG = LOG_SOTO
        Case HOJA_OVIEDO
            ObtenerNombreLOG = LOG_OVIEDO
        Case Else
            ObtenerNombreLOG = ""
    End Select
End Function

'-------------------------------------------------------------------------------------------
' Funci�n: RegistrarBorradoEnLOG
' Registra en el LOG la eliminaci�n de una fila antes de borrarla
'-------------------------------------------------------------------------------------------
Private Sub RegistrarBorradoEnLOG(ByVal ws As Worksheet, ByVal fila As Long, ByVal nombreLOG As String)
    On Error Resume Next
    
    Dim wsLOG As Worksheet
    Dim ultimaFilaLOG As Long
    Dim numOrden As Variant
    Dim nombreSolicitante As Variant
    
    Set wsLOG = ThisWorkbook.Worksheets(nombreLOG)
    If wsLOG Is Nothing Then Exit Sub
    
    ' Obtener datos de la fila a eliminar
    numOrden = ws.Cells(fila, "A").Value
    nombreSolicitante = ws.Cells(fila, "K").Value   ' Columna K = Nombre solicitante
    
    ' Encontrar �ltima fila del LOG
    ultimaFilaLOG = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Registrar el borrado
    wsLOG.Cells(ultimaFilaLOG, "A").Value = Now
    wsLOG.Cells(ultimaFilaLOG, "B").Value = ThisWorkbook.usuarioActual
    wsLOG.Cells(ultimaFilaLOG, "C").Value = "BORRADO FILA"
    wsLOG.Cells(ultimaFilaLOG, "D").Value = "N� ORDEN:  " & numOrden & " | Solicitante: " & nombreSolicitante
    wsLOG.Cells(ultimaFilaLOG, "E").Value = "Fila " & fila & " eliminada"
    
    On Error GoTo 0
End Sub


'-------------------------------------------------------------------------------------------
' Funci�n:  BuscarFilasPorOrden
' Busca todas las filas que coincidan con los N� ORDEN especificados
' Devuelve una Collection con los n�meros de fila encontrados (ordenados de mayor a menor)
'-------------------------------------------------------------------------------------------
Private Function BuscarFilasPorOrden(ByVal ws As Worksheet, ByRef listaOrdenes As Collection) As Collection
    Dim filasEncontradas As Collection
    Set filasEncontradas = New Collection
    
    Dim ultimaFila As Long
    Dim fila As Long
    Dim valorCelda As Variant
    Dim ordenBuscado As Variant
    Dim filasTemp() As Long
    Dim contadorFilas As Long
    Dim i As Long, j As Long, temp As Long
    
    ' Obtener �ltima fila con datos en columna A
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Si no hay datos, devolver colecci�n vac�a
    If ultimaFila < 2 Then
        Set BuscarFilasPorOrden = filasEncontradas
        Exit Function
    End If
    
    ' Redimensionar array temporal
    ReDim filasTemp(1 To ultimaFila)
    contadorFilas = 0
    
    ' Recorrer todas las filas buscando coincidencias
    For fila = 2 To ultimaFila
        valorCelda = ws.Cells(fila, "A").Value
        
        If Not IsEmpty(valorCelda) Then
            ' Comparar con cada N� ORDEN de la lista
            For Each ordenBuscado In listaOrdenes
                If CStr(valorCelda) = CStr(ordenBuscado) Then
                    contadorFilas = contadorFilas + 1
                    filasTemp(contadorFilas) = fila
                    Exit For  ' Ya encontr� coincidencia, pasar a siguiente fila
                End If
            Next ordenBuscado
        End If
    Next fila
    
    ' Si no se encontraron filas, devolver colecci�n vac�a
    If contadorFilas = 0 Then
        Set BuscarFilasPorOrden = filasEncontradas
        Exit Function
    End If
    
    ' Ordenar de mayor a menor (para eliminar de abajo hacia arriba)
    For i = 1 To contadorFilas - 1
        For j = i + 1 To contadorFilas
            If filasTemp(i) < filasTemp(j) Then
                temp = filasTemp(i)
                filasTemp(i) = filasTemp(j)
                filasTemp(j) = temp
            End If
        Next j
    Next i
    
    ' Pasar al Collection
    For i = 1 To contadorFilas
        filasEncontradas.Add filasTemp(i)
    Next i
    
    Set BuscarFilasPorOrden = filasEncontradas
End Function

'-------------------------------------------------------------------------------------------
' Funci�n: GenerarListaOrdenesDesdeRango
' Genera una Collection con todos los N� ORDEN dentro de un rango num�rico
'-------------------------------------------------------------------------------------------
Private Function GenerarListaOrdenesDesdeRango(ByVal desde As Long, ByVal hasta As Long) As Collection
    Dim lista As Collection
    Set lista = New Collection
    
    Dim i As Long
    For i = desde To hasta
        lista.Add CStr(i)
    Next i
    
    Set GenerarListaOrdenesDesdeRango = lista
End Function

'-------------------------------------------------------------------------------------------
' Funci�n:  GenerarListaOrdenesDesdeTexto
' Genera una Collection con los N� ORDEN a partir de un texto separado por comas
'-------------------------------------------------------------------------------------------
Private Function GenerarListaOrdenesDesdeTexto(ByVal texto As String) As Collection
    Dim lista As Collection
    Set lista = New Collection
    
    Dim partes() As String
    Dim i As Long
    Dim valor As String
    
    partes = Split(texto, ",")
    
    For i = LBound(partes) To UBound(partes)
        valor = Trim(partes(i))
        If Len(valor) > 0 Then
            lista.Add valor
        End If
    Next i
    
    Set GenerarListaOrdenesDesdeTexto = lista
End Function



'-------------------------------------------------------------------------------------------
' Sub:   BorrarSolicitudesPorOrden
' Procedimiento principal para eliminar filas por N� ORDEN
'-------------------------------------------------------------------------------------------
Public Sub BorrarSolicitudesPorOrden()
    On Error GoTo ErrorHandler
    
    ' Variables para la hoja activa
    Dim ws As Worksheet
    Dim nombreHoja As String
    Dim nombreLOG As String
    
    ' Variables para selecci�n de N� ORDEN
    Dim modoSeleccion As String
    Dim listaOrdenes As Collection
    Dim textoOrdenes As String
    Dim ordenDesde As Variant, ordenHasta As Variant
    
    ' Variables para el proceso de borrado
    Dim filasAEliminar As Collection
    Dim fila As Variant
    Dim totalFilasEliminadas As Long
    
    '======================== VERIFICAR HOJA ACTIVA ========================
    Set ws = ActiveSheet
    nombreHoja = ws.Name
    
    ' Verificar que la hoja activa sea una de las permitidas
    If nombreHoja <> HOJA_GIJON And nombreHoja <> HOJA_SOTO And nombreHoja <> HOJA_OVIEDO Then
        MsgBox "Esta funcionalidad solo est� disponible en las hojas:" & vbCrLf & vbCrLf & _
               "- " & HOJA_GIJON & vbCrLf & _
               "- " & HOJA_SOTO & vbCrLf & _
               "- " & HOJA_OVIEDO & vbCrLf & vbCrLf & _
               "Por favor, sit�ese en una de esas hojas y vuelva a ejecutar.", _
               vbExclamation, "Borrado de Solicitudes - Hoja no v�lida"
        Exit Sub
    End If
    
    ' Obtener el LOG correspondiente
    nombreLOG = ObtenerNombreLOG(nombreHoja)
    
    '======================== SELECCI�N DE MODO (con validaci�n) ========================
    Do
        modoSeleccion = InputBox( _
            "HOJA ACTIVA:  " & nombreHoja & vbCrLf & vbCrLf & _
            "Seleccione el modo de selecci�n de N� ORDEN:" & vbCrLf & vbCrLf & _
            "1.  Individual (uno o varios N� ORDEN separados por coma)" & vbCrLf & _
            "2. Rango (desde un N� ORDEN hasta otro)" & vbCrLf & vbCrLf & _
            "Introduzca 1 o 2:", _
            "Borrado de Solicitudes - Modo de selecci�n", "1")
        
        ' Si el usuario cancela (cadena vac�a o puls� Cancelar)
        If StrPtr(modoSeleccion) = 0 Then
            MsgBox "Operaci�n cancelada por el usuario.", vbInformation
            Exit Sub
        End If
        
        ' Validar que sea 1 o 2
        If val(Trim(modoSeleccion)) = 1 Or val(Trim(modoSeleccion)) = 2 Then
            Exit Do
        Else
            MsgBox "N�mero no v�lido. Por favor, introduzca 1 o 2.", vbExclamation, "Valor incorrecto"
        End If
    Loop
    
    '======================== OBTENER N� ORDEN ========================
    Select Case val(Trim(modoSeleccion))
        Case 1
            ' Modo individual
            textoOrdenes = InputBox( _
                "Introduzca los N� ORDEN a eliminar:" & vbCrLf & vbCrLf & _
                "Separe m�ltiples valores con coma (ej: 100, 150, 200)" & vbCrLf & vbCrLf & _
                "ATENCI�N: Esta operaci�n no se puede deshacer.", _
                "Borrado de Solicitudes - N� ORDEN a eliminar")
            
            If StrPtr(textoOrdenes) = 0 Or Len(Trim(textoOrdenes)) = 0 Then
                MsgBox "No se han indicado N� ORDEN.  Operaci�n cancelada.", vbExclamation
                Exit Sub
            End If
            
            Set listaOrdenes = GenerarListaOrdenesDesdeTexto(textoOrdenes)
            
        Case 2
            ' Modo rango
            ordenDesde = InputBox( _
                "Introduzca el N� ORDEN DESDE:" & vbCrLf & vbCrLf & _
                "Ejemplo: 100", _
                "Borrado de Solicitudes - N� ORDEN inicial")
            
            If StrPtr(ordenDesde) = 0 Or Len(Trim(CStr(ordenDesde))) = 0 Then
                MsgBox "No se ha indicado N� ORDEN inicial. Operaci�n cancelada.", vbExclamation
                Exit Sub
            ElseIf Not IsNumeric(ordenDesde) Then
                MsgBox "El valor '" & ordenDesde & "' no es num�rico.  Operaci�n cancelada.", vbExclamation
                Exit Sub
            End If
            
            ordenHasta = InputBox( _
                "Introduzca el N� ORDEN HASTA:" & vbCrLf & vbCrLf & _
                "Ejemplo: 200", _
                "Borrado de Solicitudes - N� ORDEN final")
            
            If StrPtr(ordenHasta) = 0 Or Len(Trim(CStr(ordenHasta))) = 0 Then
                MsgBox "No se ha indicado N� ORDEN final.  Operaci�n cancelada.", vbExclamation
                Exit Sub
            ElseIf Not IsNumeric(ordenHasta) Then
                MsgBox "El valor '" & ordenHasta & "' no es num�rico.  Operaci�n cancelada.", vbExclamation
                Exit Sub
            End If
            
            If CLng(ordenDesde) > CLng(ordenHasta) Then
                MsgBox "El N� ORDEN DESDE no puede ser mayor que HASTA.  Operaci�n cancelada.", vbExclamation
                Exit Sub
            End If
            
            Set listaOrdenes = GenerarListaOrdenesDesdeRango(CLng(ordenDesde), CLng(ordenHasta))
    End Select
    
    If listaOrdenes.Count = 0 Then
        MsgBox "No se han especificado N� ORDEN v�lidos. Operaci�n cancelada.", vbExclamation
        Exit Sub
    End If
    
    '======================== BUSCAR FILAS ========================
    Set filasAEliminar = BuscarFilasPorOrden(ws, listaOrdenes)
    
    If filasAEliminar.Count = 0 Then
        MsgBox "No se han encontrado filas con los N� ORDEN especificados.", vbInformation
        Exit Sub
    End If
    
    '======================== CONSTRUIR MENSAJE DE CONFIRMACI�N ========================
    Dim mensajeConfirmacion As String
    Dim listaOrdenesEncontrados As String
    Dim ordenesEncontrados As Collection
    Dim ordenValor As Variant
    Dim contadorMostrar As Long
    
    ' Obtener los N� ORDEN realmente encontrados (en las filas a eliminar)
    Set ordenesEncontrados = New Collection
    For Each fila In filasAEliminar
        ordenValor = ws.Cells(CLng(fila), "A").Value
        On Error Resume Next
        ordenesEncontrados.Add CStr(ordenValor), CStr(ordenValor) ' Evita duplicados
        On Error GoTo ErrorHandler
    Next fila
    
    ' Construir la lista de N� ORDEN para mostrar (m�ximo 10 + �ltimo)
    listaOrdenesEncontrados = ""
    contadorMostrar = 0
    
    For Each ordenValor In ordenesEncontrados
        contadorMostrar = contadorMostrar + 1
        
        If contadorMostrar <= 10 Then
            If Len(listaOrdenesEncontrados) > 0 Then
                listaOrdenesEncontrados = listaOrdenesEncontrados & ", "
            End If
            listaOrdenesEncontrados = listaOrdenesEncontrados & ordenValor
        ElseIf contadorMostrar = ordenesEncontrados.Count Then
            ' Es el �ltimo, a�adir con puntos suspensivos
            listaOrdenesEncontrados = listaOrdenesEncontrados & ", .. ., " & ordenValor
        End If
    Next ordenValor
    
    ' Construir mensaje de confirmaci�n
    mensajeConfirmacion = "Se han encontrado las siguientes solicitudes para eliminar:" & vbCrLf & vbCrLf
    mensajeConfirmacion = mensajeConfirmacion & nombreHoja & ": " & listaOrdenesEncontrados & vbCrLf & vbCrLf
    mensajeConfirmacion = mensajeConfirmacion & "TOTAL: " & filasAEliminar.Count & " fila(s)" & vbCrLf & vbCrLf
    mensajeConfirmacion = mensajeConfirmacion & "�Desea continuar con el borrado?" & vbCrLf & vbCrLf
    mensajeConfirmacion = mensajeConfirmacion & "ATENCI�N: Esta operaci�n NO se puede deshacer."
    
    If MsgBox(mensajeConfirmacion, vbYesNo + vbExclamation, "Confirmar borrado") <> vbYes Then
        MsgBox "Operaci�n cancelada por el usuario.", vbInformation
        Exit Sub
    End If
    
    '======================== EJECUTAR BORRADO ========================
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual
    
    totalFilasEliminadas = 0
    
    ' Eliminar filas (ya vienen ordenadas de mayor a menor)
    For Each fila In filasAEliminar
        ' Registrar en LOG antes de eliminar
        If Len(nombreLOG) > 0 Then
            RegistrarBorradoEnLOG ws, CLng(fila), nombreLOG
        End If
        
        ' Eliminar la fila completa
        ws.Rows(CLng(fila)).Delete Shift:=xlUp
        
        totalFilasEliminadas = totalFilasEliminadas + 1
    Next fila
    
    '======================== RESTAURAR Y MOSTRAR RESULTADO ========================
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    
    ' Mensaje de resumen
    Dim mensajeResumen As String
    mensajeResumen = "BORRADO COMPLETADO" & vbCrLf & vbCrLf
    mensajeResumen = mensajeResumen & nombreHoja & ":  " & totalFilasEliminadas & " fila(s) eliminada(s)" & vbCrLf & vbCrLf
    mensajeResumen = mensajeResumen & "NOTA: Si utiliza calendarios o res�menes," & vbCrLf
    mensajeResumen = mensajeResumen & "es posible que deba actualizarlos manualmente."
    
    MsgBox mensajeResumen, vbInformation, "Borrado de Solicitudes - Resultado"
    
    Exit Sub
    
ErrorHandler:
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    MsgBox "Error inesperado:  " & Err.Description, vbCritical, "Error en Borrado de Solicitudes"
End Sub


'-------------------------------------------------------------------------------------------
' Sub:  BorradoParcialCeldas_Click
' Procedimiento wrapper para el bot�n RibbonX
'-------------------------------------------------------------------------------------------
Public Sub BorradoParcialCeldas_Click(control As IRibbonControl)
    BorrarSolicitudesPorOrden
End Sub

