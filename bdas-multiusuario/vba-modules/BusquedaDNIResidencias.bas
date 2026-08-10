Attribute VB_Name = "BusquedaDNIResidencias"
'Attribute VB_Name = "BusquedaDNIResidencias"
' =============================================================================
' M�DULO: BusquedaDNIResidencias
' PROP�SITO: Buscar DNI en hojas RESIDENCIA y BDAS, y copiar datos autom�ticamente
' =============================================================================
'
' CONFIGURACI�N EN HOJA CONFIG:
' -----------------------------------------------------------------------------
' | Celda | Prop�sito                          | Valores    | Este m�dulo |
' |-------|------------------------------------|-----------:|:-----------:|
' | B5    | Archivado autom�tico RGPD          | SI/NO      | NO usa      |
' | B6    | Activar b�squeda en hojas BDAS     | SI/NO      | S� usa      |
' -----------------------------------------------------------------------------
'
' NOTA: B5 controla el archivado autom�tico (ModuloArchivadoDatos, frmPanelRGPD)
'       B6 controla si al introducir un DNI se busca tambi�n en hojas BDAS
'
' HOJAS BDAS CONSULTADAS (cuando B6 = SI):
'   - BDAS GIJ�N
'   - BDAS SOTO
'   - BDAS OVIEDO
'
' =============================================================================
' =============================================================================
' PROCEDIMIENTO PRINCIPAL: ManejarCopiaDNI
' PROP�SITO: Buscar un DNI en diferentes hojas con orden de prioridad y copiar datos
'             o limpiar celdas cuando se elimina el DNI
' PARAMETROS:
'   - Target: celda(s) modificadas
'   - ws: hoja donde ocurri� el cambio
' PRIORIDADES DE B�SQUEDA:
'   1. Misma hoja donde se introduce el DNI
'   2. Hojas BDAS (si est� activado en CONFIG B6)
' =============================================================================
Public Sub ManejarCopiaDNI(ByVal Target As Range, ByVal ws As Worksheet)
    On Error GoTo ErrorHandler
    
    Dim celda As Range
    Dim buscarEnBDAS As Boolean
    Dim screenUpdating As Boolean
    Dim enableEvents As Boolean
    Dim calculation As XlCalculation
    
    ' -------------------------------------------------------------------------
    ' 1. Guardar estado actual y optimizar rendimiento
    ' -------------------------------------------------------------------------
    screenUpdating = Application.screenUpdating
    enableEvents = Application.enableEvents
    calculation = Application.calculation
    
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual
    
    ' -------------------------------------------------------------------------
    ' 2. Verificar si la b�squeda en BDAS est� activada
    ' -------------------------------------------------------------------------
    buscarEnBDAS = EstaBusquedaBDASActivada()
    
    ' -------------------------------------------------------------------------
    ' 3. Procesar cambios en columna I (DNI)
    ' -------------------------------------------------------------------------
    If Not Intersect(Target, ws.Columns("I")) Is Nothing Then
        For Each celda In Intersect(Target, ws.Columns("I"))
            If Len(Trim(celda.Value)) > 0 Then
                ' DNI introducido - buscar y copiar datos
                ' -----------------------------------------------------------------
                ' PRIORIDAD 1: MISMA HOJA
                ' -----------------------------------------------------------------
                If BuscarYCopiarEnMismaHoja(celda.Value, ws, celda.Row) Then
                    GoTo ContinueLoop
                End If
                
                ' -----------------------------------------------------------------
                ' PRIORIDAD 2: HOJAS BDAS (si est� activado)
                ' -----------------------------------------------------------------
                If buscarEnBDAS Then
                    If BuscarYCopiarEnBDAS(celda.Value, ws, celda.Row) Then
                        GoTo ContinueLoop
                    End If
                End If
            Else
                ' DNI eliminado - limpiar celdas
                LimpiarCeldasPorDNIEliminado ws, celda.Row
            End If
ContinueLoop:
        Next celda
    End If
    
CleanUp:
    ' -------------------------------------------------------------------------
    ' 4. Restaurar estado original
    ' -------------------------------------------------------------------------
    Application.calculation = calculation
    Application.enableEvents = enableEvents
    Application.screenUpdating = screenUpdating
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en ManejarCopiaDNI - " & ws.Name & ":  " & Err.Description & _
                IIf(Not celda Is Nothing, " (Celda " & celda.Address & ")", "")
    Resume CleanUp
End Sub

' =============================================================================
' FUNCI�N: EstaBusquedaBDASActivada
' PROP�SITO: Consultar si la b�squeda en BDAS est� activada (CONFIG B6)
' DEVUELVE: True si B6 contiene "S�" o "SI", False en caso contrario
' =============================================================================
Private Function EstaBusquedaBDASActivada() As Boolean
    On Error Resume Next
    Dim valorConfig As String
    Dim wsConfig As Worksheet
    
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    If wsConfig Is Nothing Then
        EstaBusquedaBDASActivada = False
        Exit Function
    End If
    
    valorConfig = UCase(Trim(wsConfig.Range("B6").Value))
    EstaBusquedaBDASActivada = (valorConfig = "S�" Or valorConfig = "SI")
    
    On Error GoTo 0
End Function

' =============================================================================
' FUNCI�N: BuscarYCopiarEnMismaHoja
' PROP�SITO: Buscar DNI en la misma hoja y copiar datos si encuentra
' DEVUELVE: True si encontr� y copi�, False en caso contrario
' =============================================================================
Private Function BuscarYCopiarEnMismaHoja(ByVal valorBuscado As Variant, _
                                          ByVal ws As Worksheet, _
                                          ByVal filaDestino As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ultimaFila As Long
    Dim pos As Variant
    
    ultimaFila = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    ' Buscar el DNI en la columna I, excluyendo la fila actual
        Dim rangoI As Range
    Set rangoI = ws.Range("I2:I" & ultimaFila)
    
    pos = Application.Match(valorBuscado, rangoI, 0)
    
    If Not IsError(pos) Then
        ' Ajustar posici�n porque el rango empieza en fila 2
        If CLng(pos) + 1 <> filaDestino Then
            ' Encontrado en otra fila de la misma hoja
            ' Ajustar posici�n porque el rango empieza en fila 2
            CopiarDatosDesdeMismaHoja ws, CLng(pos) + 1, filaDestino
            BuscarYCopiarEnMismaHoja = True
            Exit Function
        End If
    End If
    
    BuscarYCopiarEnMismaHoja = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en BuscarYCopiarEnMismaHoja: " & Err.Description
    BuscarYCopiarEnMismaHoja = False
End Function

' =============================================================================
' FUNCI�N: BuscarYCopiarEnBDAS
' PROP�SITO: Buscar DNI en hojas BDAS (columna D) y copiar datos si encuentra
' DEVUELVE: True si encontr� y copi�, False en caso contrario
' =============================================================================
Private Function BuscarYCopiarEnBDAS(ByVal valorBuscado As Variant, _
                                     ByVal wsDestino As Worksheet, _
                                     ByVal filaDestino As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim hojasBDAS As Variant
    Dim hojaBDA As Variant
    Dim wsBDA As Worksheet
    Dim ultimaFila As Long
    Dim pos As Variant
    
    ' Lista de hojas BDAS a buscar
    hojasBDAS = Array("BDAS GIJ�N", "BDAS SOTO", "BDAS OVIEDO")
    
    For Each hojaBDA In hojasBDAS
        On Error Resume Next
        Set wsBDA = ThisWorkbook.Sheets(CStr(hojaBDA))
        On Error GoTo ErrorHandler
        
        If Not wsBDA Is Nothing Then
            ' Buscar en columna D de la hoja BDAS
            ultimaFila = wsBDA.Cells(wsBDA.Rows.Count, "D").End(xlUp).Row
            
            If ultimaFila > 0 Then
                pos = Application.Match(valorBuscado, wsBDA.Range("D2:D" & ultimaFila), 0)
                
                If Not IsError(pos) Then
                    ' Encontrado en BDAS - copiar datos
                    ' Ajustar posici�n porque el rango empieza en fila 2
                    CopiarDatosDesdeBDAS wsBDA, CLng(pos) + 1, wsDestino, filaDestino
                    BuscarYCopiarEnBDAS = True
                    Exit Function
                End If
            End If
            
            Set wsBDA = Nothing
        End If
    Next hojaBDA
    
    BuscarYCopiarEnBDAS = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en BuscarYCopiarEnBDAS: " & Err.Description
    BuscarYCopiarEnBDAS = False
End Function

' =============================================================================
' PROCEDIMIENTO: CopiarDatosDesdeMismaHoja
' PROP�SITO: Copia datos de una fila a otra en la misma hoja
' PAR�METROS:
'   - ws: hoja donde copiar
'   - filaOrigen: fila de donde copiar
'   - filaDestino: fila a donde copiar
' =============================================================================
Private Sub CopiarDatosDesdeMismaHoja(ByVal ws As Worksheet, _
                                      ByVal filaOrigen As Long, _
                                      ByVal filaDestino As Long)
    On Error GoTo ErrorHandler
    
    Dim esOviedo As Boolean
    esOviedo = EsResidenciaOviedo(ws)
    
    With ws
        ' Copiar columnas comunes
        .Cells(filaDestino, "J").Value = .Cells(filaOrigen, "J").Value
        .Cells(filaDestino, "K").Value = .Cells(filaOrigen, "K").Value
        
        If esOviedo Then
            ' RESIDENCIA OVIEDO: J, K, W, X, Y, Z, AA
            .Cells(filaDestino, "W").Value = .Cells(filaOrigen, "W").Value
            .Cells(filaDestino, "X").Value = .Cells(filaOrigen, "X").Value
            .Cells(filaDestino, "Y").Value = .Cells(filaOrigen, "Y").Value
            .Cells(filaDestino, "Z").Value = .Cells(filaOrigen, "Z").Value
            .Cells(filaDestino, "AA").Value = .Cells(filaOrigen, "AA").Value
        Else
            ' RESIDENCIA GIJ�N/SOTO: J, K, V, W, X, Y, Z
            .Cells(filaDestino, "V").Value = .Cells(filaOrigen, "V").Value
            .Cells(filaDestino, "W").Value = .Cells(filaOrigen, "W").Value
            .Cells(filaDestino, "X").Value = .Cells(filaOrigen, "X").Value
            .Cells(filaDestino, "Y").Value = .Cells(filaOrigen, "Y").Value
            .Cells(filaDestino, "Z").Value = .Cells(filaOrigen, "Z").Value
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CopiarDatosDesdeMismaHoja: " & Err.Description
End Sub

' =============================================================================
' PROCEDIMIENTO: CopiarDatosDesdeBDAS
' PROP�SITO: Copia datos desde una hoja BDAS a una hoja RESIDENCIA
' PAR�METROS:
'   - wsBDA: hoja BDAS origen
'   - filaOrigen: fila en BDAS
'   - wsDestino: hoja RESIDENCIA destino
'   - filaDestino: fila en RESIDENCIA
' MAPEO BDAS -> RESIDENCIA GIJ�N/SOTO:
'   E->K, F->V, G->W, H->X, I->Y, J->Z
' MAPEO BDAS -> RESIDENCIA OVIEDO:
'   E->K, F->W, G->X, H->Y, I->Z, J->AA
' =============================================================================
Private Sub CopiarDatosDesdeBDAS(ByVal wsBDA As Worksheet, _
                                 ByVal filaOrigen As Long, _
                                 ByVal wsDestino As Worksheet, _
                                 ByVal filaDestino As Long)
    On Error GoTo ErrorHandler
    
    Dim esOviedo As Boolean
    esOviedo = EsResidenciaOviedo(wsDestino)
    
    With wsDestino
        If esOviedo Then
            ' MAPEO PARA RESIDENCIA OVIEDO
            .Cells(filaDestino, "K").Value = wsBDA.Cells(filaOrigen, "E").Value   ' E->K
            .Cells(filaDestino, "W").Value = wsBDA.Cells(filaOrigen, "F").Value     ' F->W
            .Cells(filaDestino, "X").Value = wsBDA.Cells(filaOrigen, "G").Value   ' G->X
            .Cells(filaDestino, "Y").Value = wsBDA.Cells(filaOrigen, "H").Value    ' H->Y
            .Cells(filaDestino, "Z").Value = wsBDA.Cells(filaOrigen, "I").Value    ' I->Z
            .Cells(filaDestino, "AA").Value = wsBDA.Cells(filaOrigen, "J").Value  ' J->AA
        Else
            ' MAPEO PARA RESIDENCIA GIJ�N/SOTO
            .Cells(filaDestino, "K").Value = wsBDA.Cells(filaOrigen, "E").Value    ' E->K
            .Cells(filaDestino, "V").Value = wsBDA.Cells(filaOrigen, "F").Value   ' F->V
            .Cells(filaDestino, "W").Value = wsBDA.Cells(filaOrigen, "G").Value    ' G->W
            .Cells(filaDestino, "X").Value = wsBDA.Cells(filaOrigen, "H").Value    ' H->X
            .Cells(filaDestino, "Y").Value = wsBDA.Cells(filaOrigen, "I").Value   ' I->Y
            .Cells(filaDestino, "Z").Value = wsBDA.Cells(filaOrigen, "J").Value     ' J->Z
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en CopiarDatosDesdeBDAS: " & Err.Description & _
                " (BDAS: " & wsBDA.Name & ", Destino: " & wsDestino.Name & ")"
End Sub

' =============================================================================
' PROCEDIMIENTO: LimpiarCeldasPorDNIEliminado
' PROP�SITO: Limpiar las celdas que fueron llenadas autom�ticamente cuando
'            se elimina el DNI
' PAR�METROS:
'   - ws: hoja donde limpiar
'   - fila: fila a limpiar
' =============================================================================
Private Sub LimpiarCeldasPorDNIEliminado(ByVal ws As Worksheet, ByVal fila As Long)
    On Error GoTo ErrorHandler
    
    Dim esOviedo As Boolean
    esOviedo = EsResidenciaOviedo(ws)
    
    With ws
        ' Limpiar columnas comunes
        .Cells(fila, "J").ClearContents
        .Cells(fila, "K").ClearContents
        
        If esOviedo Then
            ' RESIDENCIA OVIEDO: J, K, W, X, Y, Z, AA
            .Cells(fila, "W").ClearContents
            .Cells(fila, "X").ClearContents
            .Cells(fila, "Y").ClearContents
            .Cells(fila, "Z").ClearContents
            .Cells(fila, "AA").ClearContents
        Else
            ' RESIDENCIA GIJ�N/SOTO: J, K, V, W, X, Y, Z
            .Cells(fila, "V").ClearContents
            .Cells(fila, "W").ClearContents
            .Cells(fila, "X").ClearContents
            .Cells(fila, "Y").ClearContents
            .Cells(fila, "Z").ClearContents
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error en LimpiarCeldasPorDNIEliminado: " & Err.Description & _
                " (Hoja: " & ws.Name & ", Fila: " & fila & ")"
End Sub

' =============================================================================
' FUNCI�N:  EsResidenciaOviedo
' PROP�SITO: Determinar si una hoja pertenece a la residencia de Oviedo
' PAR�METROS:
'   - ws: hoja a evaluar
' DEVUELVE: True si el nombre de la hoja contiene "OVIEDO", False en caso contrario
' =============================================================================
Private Function EsResidenciaOviedo(ByVal ws As Worksheet) As Boolean
    EsResidenciaOviedo = (InStr(UCase(ws.Name), "OVIEDO") > 0)
End Function
