Attribute VB_Name = "EvitarDuplicidadSolicitudesGyS"
' ======================================================================================
' EVITAR DUPLICIDAD SOLICITUDES RESIDENCIAS GIJ�N - SOTO - OVIEDO
' ======================================================================================
' Procedimiento que revisa si la solicitud actual ya existe en los �ltimos 30 d�as
' comparando DNI, fechas de entrada/salida y n�mero de pax.
' Si encuentra una coincidencia, muestra un aviso con opciones para:
'   1) Continuar sin cambios
'   2) Marcar la resoluci�n como DESESTIMADA y conservar la solicitud
'   3) Borrar toda la solicitud actual
Public Sub VerificarDuplicadoUltimos30Dias(ws As Worksheet, fila As Long)
    On Error GoTo ErrorHandler

    ' -----------------------------
    ' DECLARACI�N DE VARIABLES
    ' -----------------------------
    Dim dni As String                 ' DNI de la solicitud actual, normalizado para comparaci�n
    Dim entrada As Double             ' Fecha/hora de entrada de la solicitud actual en formato num�rico Excel
    Dim Salida As Double              ' Fecha/hora de salida de la solicitud actual en formato num�rico Excel
    Dim pax As Long                   ' N�mero de personas (pax) de la solicitud actual
    Dim fechaOrden As Variant         ' Fecha de petici�n/orden de la solicitud actual
    Dim currentFechaOrden As Date     ' Fecha de petici�n convertida a Date para validar rango
    Dim i As Long                     ' Variable para recorrer filas anteriores
    Dim celdaS As Range               ' Referencia a la celda S de la fila actual
    Dim entradaComparar As Double     ' Fecha/hora de entrada de una solicitud anterior
    Dim salidaComparar As Double      ' Fecha/hora de salida de una solicitud anterior
    Dim msgTexto As String            ' Texto del mensaje de advertencia al detectar duplicado
    Dim opcion As String              ' Opci�n elegida por el usuario en el InputBox
    Dim col As Long                   ' Variable auxiliar para limpiar celdas en la opci�n 3

    ' -----------------------------
    ' VALIDACIONES PREVIAS
    ' -----------------------------
    ' Si la fila no es v�lida (cabecera, cero, o fuera del rango de la hoja), se sale
    If fila <= 1 Or fila > ws.Rows.Count Then Exit Sub

    ' Se toma la celda S de la fila actual aunque en este bloque no se use despu�s directamente
    Set celdaS = ws.Cells(fila, "S")

    ' Se obtiene la fecha de petici�n/orden de la fila actual desde la columna B
    fechaOrden = ws.Cells(fila, "B").Value

    ' Se obtiene el DNI de la fila actual, normalizado:
    ' - Trim elimina espacios
    ' - UCase pone en may�sculas para comparar sin importar el formato
    dni = UCase(Trim(CStr(ws.Cells(fila, "I").Value)))

    ' Si las fechas de entrada o salida no son v�lidas, no se puede comparar duplicados
    If Not IsDate(ws.Cells(fila, "L").Value) Or _
       Not IsDate(ws.Cells(fila, "M").Value) Then Exit Sub

    ' Conversi�n de entrada y salida a valor num�rico de fecha de Excel
    entrada = CDbl(CDate(ws.Cells(fila, "L").Value))
    Salida = CDbl(CDate(ws.Cells(fila, "M").Value))

    ' La columna O debe contener un valor num�rico v�lido
    If Not IsNumeric(ws.Cells(fila, "O").Value) Then Exit Sub

    ' Si el n�mero de pax es cero o negativo, se aborta
    If CLng(ws.Cells(fila, "O").Value) <= 0 Then Exit Sub
    pax = CLng(ws.Cells(fila, "O").Value)

    ' Si no hay DNI, no tiene sentido buscar duplicados
    If dni = "" Then Exit Sub

    ' La fecha de peticin debe ser vlida
    If Not IsDate(fechaOrden) Then Exit Sub

    ' -----------------------------
    ' CONSULTA DE DUPLICADOS EN ACCESS
    ' -----------------------------
    Dim resCod As String
    resCod = modDatabase.ObtenerResidenciaActiva()
    If Len(resCod) = 0 Then resCod = "GIJON"
    
    If modDatabase.ExisteDuplicado(dni, resCod, CDate(ws.Cells(fila, "L").Value), CDate(ws.Cells(fila, "M").Value)) Then
        msgTexto = "⚠️ SOLICITUD DUPLICADA EN ACCESS DNI: " & dni & vbCrLf & _
            String(45, "-") & vbCrLf & _
            "DNI            : " & dni & vbCrLf & _
            "Fecha entrada  : " & Format(CDate(ws.Cells(fila, "L").Value), "dd/mm/yyyy") & vbCrLf & _
            "Fecha salida   : " & Format(CDate(ws.Cells(fila, "M").Value), "dd/mm/yyyy") & vbCrLf & _
            String(45, "-") & vbCrLf & vbCrLf & _
            "Elija una opción:" & vbCrLf & _
            "  1 ➔ Continuar sin cambios" & vbCrLf & _
            "  2 ➔ Marcar resolución como DESESTIMADA y grabar solicitud" & vbCrLf & _
            "  3 ➔ Borrar toda la solicitud actual"

                    ' ==========================================================
                    ' BUCLE DE VALIDACIN DE LA OPCIN DEL USUARIO
                    ' ==========================================================
                    Do
                        ' Se muestra un InputBox con el mensaje de duplicidad.
                        ' El valor por defecto es "1"
                        opcion = Trim(InputBox(msgTexto, "Advertencia de Duplicidad", "1"))

                        ' Si el usuario cancela o deja vac�o el cuadro, se interpreta como opci�n 1
                        If opcion = "" Then opcion = "1"

                        ' Solo son v�lidas las opciones 1, 2 o 3
                        If opcion <> "1" And opcion <> "2" And opcion <> "3" Then
                            MsgBox "Opci�n no v�lida. Introduzca 1, 2 o 3.", _
                                   vbExclamation, "Opci�n incorrecta"
                        End If
                    Loop While opcion <> "1" And opcion <> "2" And opcion <> "3"

                    ' ==========================================================
                    ' EJECUCI�N DE LA OPCI�N ELEGIDA
                    ' ==========================================================
                    Select Case opcion

                        Case "1"
                            ' Opci�n 1: no realiza cambios
                            ' Simplemente contin�a sin modificar la solicitud actual

                        Case "2"
                            ' Opci�n 2: marcar la solicitud actual como DESESTIMADA
                            ' y limpiar ciertos campos seg�n la hoja activa

                            Application.enableEvents = False
                            
                            ' Se cambia el estado de la resoluci�n a DESESTIMADA
                            ws.Cells(fila, "P").Value = "DESESTIMADA"

                            If ws.CodeName = "Hoja24" Then
                                ' OVIEDO: Limpiar solo Factura (C), Habitaciones (T) y Pagado (AB)
                                ws.Cells(fila, "C").ClearContents
                                ws.Cells(fila, "T").ClearContents
                                ws.Cells(fila, "AB").ClearContents
                            Else
                                ' GIJ�N/SOTO: Limpiar solo Factura (C), Habitaciones (S) y Pagado (AA)
                                ws.Cells(fila, "C").ClearContents
                                ws.Cells(fila, "S").ClearContents
                                ws.Cells(fila, "AA").ClearContents
                            End If

                            Application.enableEvents = True

                            ' Selecci�n de la celda correcta para disparar el aviso flotante
                            ' y llamar a la funci�n de copia
                            Dim colAviso As String
                            If ws.CodeName = "Hoja24" Then
                                colAviso = "AF"         ' OVIEDO
                            Else
                                colAviso = "AB"         ' GIJ�N/SOTO
                            End If

                            ws.Cells(fila, colAviso).Select
                            Call Portapapeles.CopiarYAvisarSiNoVacio(ws.Cells(fila, colAviso).Value, frmAviso)

                        Case "3"
                            ' Opci�n 3: borrar la solicitud actual completa
                            ' Solo se eliminan los contenidos de las celdas que no contienen f�rmulas

                            Application.enableEvents = False

                            ' Recorre todas las columnas usadas de la hoja
                            For col = 1 To ws.UsedRange.Columns.Count
                                ' Si la celda no contiene f�rmula, se borra su contenido
                                ' para evitar eliminar f�rmulas que podr�an ser necesarias
                                If Not ws.Cells(fila, col).HasFormula Then
                                    ws.Cells(fila, col).ClearContents
                                End If
                            Next col

                            Application.enableEvents = True
                    End Select

                    ' Una vez tratada la duplicidad, se sale del procedimiento
                    Exit Sub
                End If
            End If
        End If

SiguienteFila:
        ' Etiqueta usada para saltar filas inv�lidas o incompletas
        ' y continuar con la siguiente iteraci�n del bucle
    Next i
    
    Exit Sub

ErrorHandler:
    Application.enableEvents = True
    MsgBox "Error inesperado al verificar duplicados: " & Err.Description, vbCritical, "Control de Duplicados"
End Sub

