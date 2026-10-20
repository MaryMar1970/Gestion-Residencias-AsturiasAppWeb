' ======================================================================================
' EVITAR DUPLICIDAD SOLICITUDES RESIDENCIAS GIJÓN - SOTO - OVIEDO
' ======================================================================================
' Procedimiento que revisa si la solicitud actual ya existe en los últimos 30 días
' comparando DNI, fechas de entrada/salida y número de pax.
' Si encuentra una coincidencia, muestra un aviso con opciones para:
'   1) Continuar sin cambios
'   2) Marcar la resolución como DESESTIMADA y conservar la solicitud
'   3) Borrar toda la solicitud actual
Public Sub VerificarDuplicadoUltimos30Dias(ws As Worksheet, fila As Long)
    On Error GoTo ErrorHandler

    ' -----------------------------
    ' DECLARACIÓN DE VARIABLES
    ' -----------------------------
    Dim dni As String                 ' DNI de la solicitud actual, normalizado para comparación
    Dim entrada As Double             ' Fecha/hora de entrada de la solicitud actual en formato numérico Excel
    Dim Salida As Double              ' Fecha/hora de salida de la solicitud actual en formato numérico Excel
    Dim pax As Long                   ' Número de personas (pax) de la solicitud actual
    Dim fechaOrden As Variant         ' Fecha de petición/orden de la solicitud actual
    Dim currentFechaOrden As Date     ' Fecha de petición convertida a Date para validar rango
    Dim i As Long                     ' Variable para recorrer filas anteriores
    Dim celdaS As Range               ' Referencia a la celda S de la fila actual
    Dim entradaComparar As Double     ' Fecha/hora de entrada de una solicitud anterior
    Dim salidaComparar As Double      ' Fecha/hora de salida de una solicitud anterior
    Dim msgTexto As String            ' Texto del mensaje de advertencia al detectar duplicado
    Dim opcion As String              ' Opción elegida por el usuario en el InputBox
    Dim col As Long                   ' Variable auxiliar para limpiar celdas en la opción 3

    ' -----------------------------
    ' VALIDACIONES PREVIAS
    ' -----------------------------
    ' Si la fila no es válida (cabecera, cero, o fuera del rango de la hoja), se sale
    If fila <= 1 Or fila > ws.Rows.Count Then Exit Sub

    ' Se toma la celda S de la fila actual aunque en este bloque no se use después directamente
    Set celdaS = ws.Cells(fila, "S")

    ' Se obtiene la fecha de petición/orden de la fila actual desde la columna B
    fechaOrden = ws.Cells(fila, "B").Value

    ' Se obtiene el DNI de la fila actual, normalizado:
    ' - Trim elimina espacios
    ' - UCase pone en mayúsculas para comparar sin importar el formato
    dni = UCase(Trim(CStr(ws.Cells(fila, "I").Value)))

    ' Si las fechas de entrada o salida no son válidas, no se puede comparar duplicados
    If Not IsDate(ws.Cells(fila, "L").Value) Or _
       Not IsDate(ws.Cells(fila, "M").Value) Then Exit Sub

    ' Conversión de entrada y salida a valor numérico de fecha de Excel
    entrada = CDbl(CDate(ws.Cells(fila, "L").Value))
    Salida = CDbl(CDate(ws.Cells(fila, "M").Value))

    ' La columna O debe contener un valor numérico válido
    If Not IsNumeric(ws.Cells(fila, "O").Value) Then Exit Sub

    ' Si el número de pax es cero o negativo, se aborta
    If CLng(ws.Cells(fila, "O").Value) <= 0 Then Exit Sub
    pax = CLng(ws.Cells(fila, "O").Value)

    ' Si no hay DNI, no tiene sentido buscar duplicados
    If dni = "" Then Exit Sub

    ' La fecha de petición debe ser válida
    If Not IsDate(fechaOrden) Then Exit Sub

    ' -----------------------------
    ' CONSULTA DE DUPLICADOS EN ACCESS
    ' -----------------------------
    Dim resCod As String
    resCod = modDatabase.ObtenerClaveResidenciaDesdeHoja(ws.Name)
    If Len(resCod) = 0 Then resCod = "GIJON"

    
    If modDatabase.ExisteDuplicado(dni, resCod, CDate(ws.Cells(fila, "L").Value), CDate(ws.Cells(fila, "M").Value)) Then
        msgTexto = "ATENCION: SOLICITUD DUPLICADA EN ACCESS DNI: " & dni & vbCrLf & _
            String(45, "-") & vbCrLf & _
            "DNI            : " & dni & vbCrLf & _
            "Fecha entrada  : " & Format(CDate(ws.Cells(fila, "L").Value), "dd/mm/yyyy") & vbCrLf & _
            "Fecha salida   : " & Format(CDate(ws.Cells(fila, "M").Value), "dd/mm/yyyy") & vbCrLf & _
            String(45, "-") & vbCrLf & vbCrLf & _
            "Elija una opción:" & vbCrLf & _
            "  1 -> Continuar sin cambios" & vbCrLf & _
            "  2 -> Marcar resolución como DESESTIMADA y grabar solicitud" & vbCrLf & _
            "  3 -> Borrar toda la solicitud actual"

                    ' ==========================================================
                    ' BUCLE DE VALIDACIÓN DE LA OPCIÓN DEL USUARIO
                    ' ==========================================================
                    Do
                        ' Se muestra un InputBox con el mensaje de duplicidad.
                        ' El valor por defecto es "1"
                        opcion = Trim(InputBox(msgTexto, "Advertencia de Duplicidad", "1"))

                        ' Si el usuario cancela o deja vacío el cuadro, se interpreta como opción 1
                        If opcion = "" Then opcion = "1"

                        ' Solo son válidas las opciones 1, 2 o 3
                        If opcion <> "1" And opcion <> "2" And opcion <> "3" Then
                            MsgBox "Opción no válida. Introduzca 1, 2 o 3.", _
                                   vbExclamation, "Opción incorrecta"
                        End If
                    Loop While opcion <> "1" And opcion <> "2" And opcion <> "3"

                    ' ==========================================================
                    ' EJECUCIÓN DE LA OPCIÓN ELEGIDA
                    ' ==========================================================
                    Select Case opcion

                        Case "1"
                            ' Opción 1: no realiza cambios
                            ' Simplemente continúa sin modificar la solicitud actual

                        Case "2"
                            ' Opción 2: marcar la solicitud actual como DESESTIMADA
                            ' y limpiar ciertos campos según la hoja activa

                            Application.enableEvents = False
                            
                            ' Se cambia el estado de la resolución a DESESTIMADA
                            ws.Cells(fila, "P").Value = "DESESTIMADA"

                            If ws.CodeName = "Hoja24" Then
                                ' OVIEDO: Limpiar solo Factura (C), Habitaciones (T) y Pagado (AB)
                                ws.Cells(fila, "C").ClearContents
                                ws.Cells(fila, "T").ClearContents
                                ws.Cells(fila, "AB").ClearContents
                            Else
                                ' GIJÓN/SOTO: Limpiar solo Factura (C), Habitaciones (S) y Pagado (AA)
                                ws.Cells(fila, "C").ClearContents
                                ws.Cells(fila, "S").ClearContents
                                ws.Cells(fila, "AA").ClearContents
                            End If

                            Application.enableEvents = True

                            ' Selección de la celda correcta para disparar el aviso flotante
                            ' y llamar a la función de copia
                            Dim colAviso As String
                            If ws.CodeName = "Hoja24" Then
                                colAviso = "AF"         ' OVIEDO
                            Else
                                colAviso = "AB"         ' GIJÓN/SOTO
                            End If

                            ws.Cells(fila, colAviso).Select
                            Call Portapapeles.CopiarYAvisarSiNoVacio(ws.Cells(fila, colAviso).Value, frmAviso)

                        Case "3"
                            ' Opción 3: borrar la solicitud actual completa
                            ' Solo se eliminan los contenidos de las celdas que no contienen fórmulas

                            Application.enableEvents = False

                            ' Recorre todas las columnas usadas de la hoja
                            For col = 1 To ws.UsedRange.Columns.Count
                                ' Si la celda no contiene fórmula, se borra su contenido
                                ' para evitar eliminar fórmulas que podrían ser necesarias
                                If Not ws.Cells(fila, col).HasFormula Then
                                    ws.Cells(fila, col).ClearContents
                                End If
                            Next col

                            Application.enableEvents = True
                    End Select

                    ' Una vez tratada la duplicidad, se sale del procedimiento
                    Exit Sub
                End If

    Exit Sub

ErrorHandler:
    Application.enableEvents = True
    MsgBox "Error inesperado al verificar duplicados: " & Err.Description, vbCritical, "Control de Duplicados"
End Sub



