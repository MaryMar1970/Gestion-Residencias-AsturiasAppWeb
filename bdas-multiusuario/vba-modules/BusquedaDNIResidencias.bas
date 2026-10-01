Attribute VB_Name = "BusquedaDNIResidencias"
'Attribute VB_Name = "BusquedaDNIResidencias"
' =============================================================================
' Mï¿½DULO: BusquedaDNIResidencias
' PROPï¿½SITO: Buscar DNI en hojas RESIDENCIA y BDAS, y copiar datos automï¿½ticamente
' =============================================================================
'
' CONFIGURACIï¿½N EN HOJA CONFIG:
' -----------------------------------------------------------------------------
' | Celda | Propï¿½sito                          | Valores    | Este módulo |
' |-------|------------------------------------|-----------:|:-----------:|
' | B5    | Archivado automï¿½tico RGPD          | SI/NO      | NO usa      |
' | B6    | Activar bï¿½squeda en hojas BDAS     | SI/NO      | Sï¿½ usa      |
' -----------------------------------------------------------------------------
'
' NOTA: B5 controla el archivado automï¿½tico (ModuloArchivadoDatos, frmPanelRGPD)
'       B6 controla si al introducir un DNI se busca tambiï¿½n en hojas BDAS
'
' HOJAS BDAS CONSULTADAS (cuando B6 = SI):
'   - BDAS GIJÓN
'   - BDAS SOTO
'   - BDAS OVIEDO
'
' =============================================================================
' =============================================================================
' PROCEDIMIENTO PRINCIPAL: ManejarCopiaDNI
' PROPï¿½SITO: Buscar un DNI en diferentes hojas con orden de prioridad y copiar datos
'             o limpiar celdas cuando se elimina el DNI
' PARAMETROS:
'   - Target: celda(s) modificadas
'   - ws: hoja donde ocurriï¿½ el cambio
' PRIORIDADES DE Bï¿½SQUEDA:
'   1. Misma hoja donde se introduce el DNI
'   2. Hojas BDAS (si está activado en CONFIG B6)
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
    ' 2. Verificar si la bï¿½squeda en BDAS está activada
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
                ' PRIORIDAD 2: HOJAS BDAS (si está activado)
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
' FUNCIÓN: EstaBusquedaBDASActivada
' PROPï¿½SITO: Consultar si la bï¿½squeda en BDAS está activada (CONFIG B6)
' DEVUELVE: True si B6 contiene "Sï¿½" o "SI", False en caso contrario
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
    EstaBusquedaBDASActivada = (valorConfig = "Sï¿½" Or valorConfig = "SI")
    
    On Error GoTo 0
End Function

' =============================================================================
' FUNCIÓN: BuscarYCopiarEnMismaHoja
' PROPï¿½SITO: Buscar DNI en la misma hoja y copiar datos si encuentra
' DEVUELVE: True si encontrï¿½ y copiï¿½, False en caso contrario
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
        ' Ajustar posiciï¿½n porque el rango empieza en fila 2
        If CLng(pos) + 1 <> filaDestino Then
            ' Encontrado en otra fila de la misma hoja
            ' Ajustar posiciï¿½n porque el rango empieza en fila 2
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
' FUNCIÓN: BuscarYCopiarEnBDAS
' PROPÓSITO: Buscar DNI en Access DB (Activas + Histórico) y autocompletar
'            Nombre, Empleo, Situación, Rango, Teléfono, Dirección, CP, Población y Provincia
' DEVUELVE: True si encontró y copió, False en caso contrario
' =============================================================================
Private Function BuscarYCopiarEnBDAS(ByVal valorBuscado As Variant, _
                                     ByVal wsDestino As Worksheet, _
                                     ByVal filaDestino As Long) As Boolean
    On Error GoTo ErrorHandler
    
    Dim rs As Object
    Set rs = modDatabase.BuscarEnOrdenes("DNI", CStr(valorBuscado))
    
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            Dim esOviedo As Boolean
            esOviedo = EsResidenciaOviedo(wsDestino)
            
            Dim nomCompleto As String
            nomCompleto = Trim(CStr(rs("Nombre").Value))
            If Not IsNull(rs("Apellidos").Value) Then
                If Len(Trim(CStr(rs("Apellidos").Value))) > 0 Then
                    If InStr(nomCompleto, Trim(CStr(rs("Apellidos").Value))) = 0 Then
                        nomCompleto = nomCompleto & " " & Trim(CStr(rs("Apellidos").Value))
                    End If
                End If
            End If
            
            With wsDestino
                ' 1. Datos comunes profesionales y personales (Cols E, F, J, K)
                If Len(Trim(CStr(.Cells(filaDestino, "E").Value))) = 0 And Not IsNull(rs("Empleo").Value) Then
                    .Cells(filaDestino, "E").Value = rs("Empleo").Value
                End If
                If Len(Trim(CStr(.Cells(filaDestino, "F").Value))) = 0 And Not IsNull(rs("Situacion").Value) Then
                    .Cells(filaDestino, "F").Value = rs("Situacion").Value
                End If
                If Len(Trim(CStr(.Cells(filaDestino, "J").Value))) = 0 And Not IsNull(rs("Rango").Value) Then
                    .Cells(filaDestino, "J").Value = rs("Rango").Value
                End If
                If Len(Trim(CStr(.Cells(filaDestino, "K").Value))) = 0 And Len(nomCompleto) > 0 Then
                    .Cells(filaDestino, "K").Value = nomCompleto
                End If
                
                ' 2. Datos de contacto y domicilio según residencia
                If esOviedo Then
                    ' RESIDENCIA OVIEDO: W=Teléfono, X=Dirección, Y=CP, Z=Población, AA=Provincia
                    If Len(Trim(CStr(.Cells(filaDestino, "W").Value))) = 0 And Not IsNull(rs("Telefono").Value) Then
                        .Cells(filaDestino, "W").Value = rs("Telefono").Value
                    End If
                    If Len(Trim(CStr(.Cells(filaDestino, "X").Value))) = 0 And Not IsNull(rs("Direccion").Value) Then
                        .Cells(filaDestino, "X").Value = rs("Direccion").Value
                    End If
                    If Len(Trim(CStr(.Cells(filaDestino, "Y").Value))) = 0 And Not IsNull(rs("CodigoPostal").Value) Then
                        .Cells(filaDestino, "Y").Value = rs("CodigoPostal").Value
                    End If
                    If Len(Trim(CStr(.Cells(filaDestino, "Z").Value))) = 0 And Not IsNull(rs("Poblacion").Value) Then
                        .Cells(filaDestino, "Z").Value = rs("Poblacion").Value
                    End If
                    If Len(Trim(CStr(.Cells(filaDestino, "AA").Value))) = 0 And Not IsNull(rs("Provincia").Value) Then
                        .Cells(filaDestino, "AA").Value = rs("Provincia").Value
                    End If
                Else
                    ' RESIDENCIA GIJÓN / SOTO: V=Teléfono, W=Dirección, X=CP, Y=Población, Z=Provincia
                    If Len(Trim(CStr(.Cells(filaDestino, "V").Value))) = 0 And Not IsNull(rs("Telefono").Value) Then
                        .Cells(filaDestino, "V").Value = rs("Telefono").Value
                    End If
                    If Len(Trim(CStr(.Cells(filaDestino, "W").Value))) = 0 And Not IsNull(rs("Direccion").Value) Then
                        .Cells(filaDestino, "W").Value = rs("Direccion").Value
                    End If
                    If Len(Trim(CStr(.Cells(filaDestino, "X").Value))) = 0 And Not IsNull(rs("CodigoPostal").Value) Then
                        .Cells(filaDestino, "X").Value = rs("CodigoPostal").Value
                    End If
                    If Len(Trim(CStr(.Cells(filaDestino, "Y").Value))) = 0 And Not IsNull(rs("Poblacion").Value) Then
                        .Cells(filaDestino, "Y").Value = rs("Poblacion").Value
                    End If
                    If Len(Trim(CStr(.Cells(filaDestino, "Z").Value))) = 0 And Not IsNull(rs("Provincia").Value) Then
                        .Cells(filaDestino, "Z").Value = rs("Provincia").Value
                    End If
                End If
            End With
            
            rs.Close
            BuscarYCopiarEnBDAS = True
            Exit Function
        End If
        rs.Close
    End If
    
    BuscarYCopiarEnBDAS = False
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en BuscarYCopiarEnBDAS: " & Err.Description
    BuscarYCopiarEnBDAS = False
End Function

' =============================================================================
' PROCEDIMIENTO: CopiarDatosDesdeMismaHoja
' PROPÓSITO: Copia datos de una fila a otra en la misma hoja
' =============================================================================
Private Sub CopiarDatosDesdeMismaHoja(ByVal ws As Worksheet, _
                                      ByVal filaOrigen As Long, _
                                      ByVal filaDestino As Long)
    On Error GoTo ErrorHandler
    
    Dim esOviedo As Boolean
    esOviedo = EsResidenciaOviedo(ws)
    
    With ws
        ' Copiar columnas comunes (Empleo, Situación, Rango, Nombre)
        .Cells(filaDestino, "E").Value = .Cells(filaOrigen, "E").Value
        .Cells(filaDestino, "F").Value = .Cells(filaOrigen, "F").Value
        .Cells(filaDestino, "J").Value = .Cells(filaOrigen, "J").Value
        .Cells(filaDestino, "K").Value = .Cells(filaOrigen, "K").Value
        
        If esOviedo Then
            ' RESIDENCIA OVIEDO: W, X, Y, Z, AA
            .Cells(filaDestino, "W").Value = .Cells(filaOrigen, "W").Value
            .Cells(filaDestino, "X").Value = .Cells(filaOrigen, "X").Value
            .Cells(filaDestino, "Y").Value = .Cells(filaOrigen, "Y").Value
            .Cells(filaDestino, "Z").Value = .Cells(filaOrigen, "Z").Value
            .Cells(filaDestino, "AA").Value = .Cells(filaOrigen, "AA").Value
        Else
            ' RESIDENCIA GIJÓN/SOTO: V, W, X, Y, Z
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
' PROCEDIMIENTO: LimpiarCeldasPorDNIEliminado
' PROPÓSITO: Limpiar las celdas autocompletadas cuando se borra el DNI
' =============================================================================
Private Sub LimpiarCeldasPorDNIEliminado(ByVal ws As Worksheet, ByVal fila As Long)
    On Error GoTo ErrorHandler
    
    Dim esOviedo As Boolean
    esOviedo = EsResidenciaOviedo(ws)
    
    With ws
        ' Limpiar columnas comunes (Empleo, Situación, Rango, Nombre)
        .Cells(fila, "E").ClearContents
        .Cells(fila, "F").ClearContents
        .Cells(fila, "J").ClearContents
        .Cells(fila, "K").ClearContents
        
        If esOviedo Then
            ' RESIDENCIA OVIEDO: W, X, Y, Z, AA
            .Cells(fila, "W").ClearContents
            .Cells(fila, "X").ClearContents
            .Cells(fila, "Y").ClearContents
            .Cells(fila, "Z").ClearContents
            .Cells(fila, "AA").ClearContents
        Else
            ' RESIDENCIA GIJÓN/SOTO: V, W, X, Y, Z
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
' FUNCIÓN:  EsResidenciaOviedo
' PROPï¿½SITO: Determinar si una hoja pertenece a la residencia de Oviedo
' PARï¿½METROS:
'   - ws: hoja a evaluar
' DEVUELVE: True si el nombre de la hoja contiene "OVIEDO", False en caso contrario
' =============================================================================
Private Function EsResidenciaOviedo(ByVal ws As Worksheet) As Boolean
    EsResidenciaOviedo = (InStr(UCase(ws.Name), "OVIEDO") > 0)
End Function


