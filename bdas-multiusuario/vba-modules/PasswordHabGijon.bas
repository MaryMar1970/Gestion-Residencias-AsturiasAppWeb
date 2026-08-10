Attribute VB_Name = "PasswordHabGijon"
Option Explicit

'=================================================================================
' Funci�n: ObtenerContrasenasGijon
'---------------------------------------------------------------------------------
' OBJETIVO
'   Convertir una lista de habitaciones (separadas por comas) en la lista de
'   contrase�as correspondientes, buscando en la hoja "CONTRASE�AS GIJ�N".
'
' CONTEXTO (por qu� existe esta funci�n)
'   En la hoja "RESIDENCIA GIJ�N", en la columna S (N�M HAB.) puede aparecer:
'     - Habitaciones num�ricas: 1, 2, 3, 4, 5, 6, 7 ...
'     - Habitaciones alfanum�ricas / especiales: Of.1, Of.2, Of.3, Est.1, Est.2, Est.3 ...
'
'   Cuando el usuario pulsa la celda de la columna AC ("COPIAR TEXTO EMAIL"),
'   se genera un texto que (entre otras cosas) necesita traducir esas habitaciones
'   a sus contrase�as.
'
' ESTRATEGIA DE SOLUCI�N
'   1) Tratar cada habitaci�n como TEXTO (String) para soportar valores como "Of.1".
'   2) Hacer VLookup primero con clave de texto.
'   3) Si no se encuentra y la clave es num�rica, intentar VLookup como n�mero (CLng),
'      para cubrir el caso en que la tabla de contrase�as tenga "101" como n�mero real
'      (no como texto).
'   4) Evitar lanzar errores de VBA:
'       - Usamos Application.VLookup (devuelve un Variant que puede ser Error)
'         en lugar de WorksheetFunction.VLookup (lanza error en tiempo de ejecuci�n).
'
' REQUISITOS DE DATOS (MUY IMPORTANTE)
'   - Debe existir la hoja: "CONTRASE�AS GIJ�N" (nombre exacto)
'   - La tabla debe estar en columnas A:B:
'       Col A = clave habitaci�n (ej: 1, 2, 3, Of.1, Est.2, 101...)
'       Col B = contrase�a correspondiente
'   - El rango de b�squeda es A2:B14(ultimaFila). Se asume cabecera en fila 1.
'
' ENTRADA
'   strHabitaciones: String
'     - Cadena con habitaciones separadas por comas, ej:
'         "1, 2, 3"
'         "Of.1, Est.2, 4"
'     - Se toleran espacios alrededor de cada elemento.
'
' SALIDA
'   String con contrase�as separadas por comas, en el mismo orden que la entrada:
'     - Si una habitaci�n no existe, se devuelve "#N/D" en su posici�n.
'
' EJEMPLO
'   Entrada : "Of.1, 4, Est.2"
'   Salida  : "pwdOf1, pwd4, pwdEst2"   (si existen en la tabla)
'
'=================================================================================
Public Function ObtenerContrasenasGijon(ByVal strHabitaciones As String) As String

    '-----------------------------
    ' 0) Variables de trabajo
    '-----------------------------
    Dim arrHab() As String          ' Array de habitaciones separadas desde la cadena de entrada
    Dim arrPwd() As String          ' Array de contrase�as resultantes (mismo tama�o que arrHab)
    Dim shtPWD As Worksheet         ' Hoja donde est� la tabla de contrase�as
    Dim ultimaFila As Long          ' �ltima fila con datos en columna A (para acotar el rango)
    Dim i As Long                   ' �ndice del bucle

    Dim habKey As String            ' Habitaci�n actual ya normalizada (trim)
    Dim result As Variant           ' Resultado de VLookup (puede ser valor o error Variant)

    '-----------------------------------------------------
    ' 1) Obtener referencia a la hoja de contrase�as
    '-----------------------------------------------------
    ' NOTA:
    '   Si el nombre no coincide exactamente (acentos, espacios, may�sculas),
    '   Excel lanzar� error 9 "Subscript out of range".
    Set shtPWD = ThisWorkbook.Worksheets("CONTRASE�AS GIJ�N")

    '-----------------------------------------------------
    ' 2) Limpieza inicial de la entrada
    '-----------------------------------------------------
    ' - Elimina espacios al principio y al final.
    ' - Si queda vac�o, no hay nada que buscar.
    strHabitaciones = Trim$(strHabitaciones)
    If strHabitaciones = vbNullString Then
        ObtenerContrasenasGijon = vbNullString
        Exit Function
    End If

    '-----------------------------------------------------
    ' 3) Tokenizar (separar) la cadena por comas
    '-----------------------------------------------------
    ' Ej: "Of.1, Est.2, 4" -> arrHab(0)="Of.1" ; arrHab(1)=" Est.2" ; arrHab(2)=" 4"
    ' Despu�s normalizaremos cada token con Trim$.
    arrHab = Split(strHabitaciones, ",")

    '-----------------------------------------------------
    ' 4) Preparar array de salida con mismo tama�o
    '-----------------------------------------------------
    ReDim arrPwd(LBound(arrHab) To UBound(arrHab))

    '-----------------------------------------------------
    ' 5) Calcular �ltima fila en la tabla de contrase�as
    '-----------------------------------------------------
    ' Acota el rango A2:BultimaFila para que el VLookup no recorra m�s de lo necesario.
    ultimaFila = shtPWD.Cells(shtPWD.Rows.Count, "A").End(xlUp).Row

    ' Si no hay datos (solo cabecera o vac�o), devolvemos vac�o (o podr�amos devolver #N/D para todo)
    If ultimaFila < 2 Then
        ObtenerContrasenasGijon = vbNullString
        Exit Function
    End If

    '-----------------------------------------------------
    ' 6) Recorrer habitaciones y resolver contrase�a
    '-----------------------------------------------------
    For i = LBound(arrHab) To UBound(arrHab)

        ' 6.1) Normalizar el token (quita espacios extra)
        habKey = Trim$(arrHab(i))

        ' 6.2) Si el elemento viene vac�o (ej: "1,,2"), devolver #N/D en esa posici�n
        If habKey = vbNullString Then
            arrPwd(i) = "#N/D"

        Else
            '-------------------------------------------------------------
            ' FLUJO PRINCIPAL DE B�SQUEDA (tolerante a texto y n�meros)
            '-------------------------------------------------------------
            ' A) Buscar como TEXTO primero:
            '    - Esto cubre "Of.1", "Est.2" y tambi�n "4" cuando est� guardado como texto.
            ' B) Si falla y habKey es num�rico, intentar como N�MERO:
            '    - Esto cubre cuando la tabla tiene el 4 como n�mero real (no texto).
            '
            ' IMPORTANTE:
            '   Usamos Application.VLookup en vez de WorksheetFunction.VLookup para
            '   NO lanzar un error de ejecuci�n; Application.VLookup devuelve un Variant
            '   con error (IsError=True) cuando no encuentra la clave.
            '-------------------------------------------------------------

            ' A) Intento 1: clave como texto
            result = Application.VLookup( _
                        habKey, _
                        shtPWD.Range("A2:B14" & ultimaFila), _
                        2, _
                        False)

            ' B) Intento 2: si no aparece y es num�rico, buscar como n�mero
            If IsError(result) Then
                If IsNumeric(habKey) Then
                    ' CLng aqu� es seguro porque ya comprobamos IsNumeric.
                    result = Application.VLookup( _
                                CLng(habKey), _
                                shtPWD.Range("A2:B14" & ultimaFila), _
                                2, _
                                False)
                End If
            End If

            ' 6.3) Mapear el resultado al array de salida
            If IsError(result) Then
                ' No encontrado o tabla no v�lida para esa clave
                arrPwd(i) = "#N/D"
            Else
                ' Convertimos a String por seguridad (por si la contrase�a es num�rica)
                arrPwd(i) = CStr(result)
            End If
        End If
    Next i

    '-----------------------------------------------------
    ' 7) Unir el resultado en una cadena CSV de contrase�as
    '-----------------------------------------------------
    ' Respeta el orden original de habitaciones.
    ObtenerContrasenasGijon = Join(arrPwd, ",")

End Function

