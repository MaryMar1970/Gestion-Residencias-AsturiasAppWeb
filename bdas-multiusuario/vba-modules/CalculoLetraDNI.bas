Attribute VB_Name = "CalculoLetraDNI"
' =========================================================================
' SUBRUTINA: VerificarOCorregirUnaCeldaDNI
' USO:
'    Llamar desde Worksheet_Change, pasando la celda de la columna I que contiene el DNI/NIE.
'    Esta rutina valida y asiste al usuario en la correcci�n del DNI/NIE, manejando tanto DNIs est�ndar como NIEs (X, Y, Z).
' FLUJO GENERAL:
'    1. Analiza el contenido de la celda (puede ser solo n�mero, DNI completo, NIE completo o NIE sin letra).
'    2. Si la letra es incorrecta o falta, pregunta antes de modificar.
'    3. Si el formato es inv�lido, informa y limpia la celda.
' DEPENDENCIAS:
'    Requiere el m�dulo CalculoLetraDNIAux con funciones:
'      - CalcularLetraDNI
'      - ObtenerNumeroDNI
'      - ObtenerLetraDNI
'      - EsNumeroDNIValido
' =========================================================================
Public Sub VerificarOCorregirUnaCeldaDNI(ByVal celda As Range)
    On Error GoTo ErrDNI
    Dim dniCompleto As String          ' Valor introducido en la celda, ya en may�sculas y sin espacios
    Dim numeroDNI As String            ' Parte num�rica del DNI o NIE
    Dim letraCorrecta As String        ' Letra calculada que deber�a corresponder al n�mero
    Dim letraActual As String          ' Letra actualmente escrita en la celda (si la hay)
    Dim respuesta As VbMsgBoxResult    ' Respuesta del usuario ante los mensajes de confirmaci�n

    ' Normaliza la entrada: elimina espacios y convierte a may�sculas
    dniCompleto = UCase(Trim(celda.Value))

    ' =========================================================================
    ' CASO 1: Solo n�mero (8 d�gitos, sin letra)
    ' =========================================================================
    If Len(dniCompleto) = 8 And IsNumeric(dniCompleto) Then
        numeroDNI = dniCompleto
        ' Calcula la letra que corresponde a esos d�gitos
        letraCorrecta = CalculoLetraDNIAux.CalcularLetraDNI(numeroDNI)
        ' Pregunta al usuario si desea completar la letra autom�ticamente
        respuesta = MsgBox("Para el n�mero '" & numeroDNI & "' la letra correcta es '" & letraCorrecta & "'." & vbCrLf & _
                           "�Desea completar el DNI con la letra?", vbExclamation + vbYesNo, "Completar letra DNI")
        If respuesta = vbYes Then
            celda.Value = numeroDNI & UCase(letraCorrecta)
        End If
        Exit Sub
    End If

    ' =========================================================================
    ' CASO 2: DNI/NIE completo (9 caracteres: 8 d�gitos + 1 letra, o NIE tipo X1234567L)
    ' =========================================================================
    If Len(dniCompleto) = 9 Then
        ' Extrae la parte num�rica (admite NIE con X/Y/Z)
        numeroDNI = CalculoLetraDNIAux.ObtenerNumeroDNI(dniCompleto)
        ' Verifica validez del n�mero (8 d�gitos)
        If CalculoLetraDNIAux.EsNumeroDNIValido(numeroDNI) Then
            letraCorrecta = CalculoLetraDNIAux.CalcularLetraDNI(numeroDNI)
            letraActual = CalculoLetraDNIAux.ObtenerLetraDNI(dniCompleto)
            ' Si la letra no coincide con la calculada, pregunta si desea corregirla
            If UCase(letraActual) <> UCase(letraCorrecta) Then
                respuesta = MsgBox("Atenci�n: Para el n�mero '" & numeroDNI & "' la letra correcta es '" & letraCorrecta & "'." & vbCrLf & _
                                   "Verifique que los d�gitos sean los correctos." & vbCrLf & _
                                   "�Desea modificar la letra?", vbExclamation + vbYesNo, "Letra de DNI incorrecta")
                If respuesta = vbYes Then
                    celda.Value = numeroDNI & UCase(letraCorrecta)
                End If
                Exit Sub
            End If
            ' Si la letra es correcta pero est� en min�scula, corregir a may�sculas
            If letraActual <> UCase(letraActual) Then
                celda.Value = numeroDNI & UCase(letraActual)
            End If
            Exit Sub
        End If
    End If

    ' =========================================================================
    ' CASO 3: NIE incompleto o NIE con posible error en la letra (X, Y, Z inicial)
    '         Se maneja tanto NIE con letra como NIE s�lo con n�mero
    ' =========================================================================
    If (Left(dniCompleto, 1) = "X" Or Left(dniCompleto, 1) = "Y" Or Left(dniCompleto, 1) = "Z") Then
        numeroDNI = CalculoLetraDNIAux.ObtenerNumeroDNI(dniCompleto)
        If CalculoLetraDNIAux.EsNumeroDNIValido(numeroDNI) Then
            letraCorrecta = CalculoLetraDNIAux.CalcularLetraDNI(numeroDNI)
            If Len(dniCompleto) = 9 Then
                letraActual = CalculoLetraDNIAux.ObtenerLetraDNI(dniCompleto)
                ' Si la letra del NIE no coincide, ofrece corregir
                If UCase(letraActual) <> UCase(letraCorrecta) Then
                    respuesta = MsgBox("Atenci�n: Para el n�mero '" & numeroDNI & "' la letra correcta es '" & letraCorrecta & "'." & vbCrLf & _
                                       "Verifique que los d�gitos sean los correctos." & vbCrLf & _
                                       "�Desea modificar la letra?", vbExclamation + vbYesNo, "Letra de NIE incorrecta")
                    If respuesta = vbYes Then
                        celda.Value = Left(dniCompleto, 8) & UCase(letraCorrecta)
                    End If
                    Exit Sub
                End If
            ElseIf Len(dniCompleto) = 8 Then
                ' Si el NIE est� sin letra, ofrece completarla
                respuesta = MsgBox("Para el n�mero '" & dniCompleto & "' la letra correcta es '" & letraCorrecta & "'." & vbCrLf & _
                                   "�Desea completar el NIE con la letra?", vbExclamation + vbYesNo, "Completar letra NIE")
                If respuesta = vbYes Then
                    celda.Value = dniCompleto & UCase(letraCorrecta)
                End If
                Exit Sub
            End If
            Exit Sub
        End If
    End If

    ' =========================================================================
    ' CASO 4: Formato inv�lido
    '         Si no es un DNI/NIE v�lido, borra la celda y avisa.
    ' =========================================================================
    MsgBox "El n�mero de DNI no es v�lido (debe tener 8 d�gitos y ser num�rico, o NIE correcto). Por favor, corr�gelo.", vbCritical, "DNI Inv�lido"
    celda.ClearContents
    celda.Select
    Exit Sub

ErrDNI:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "VerificarOCorregirUnaCeldaDNI"
End Sub

