Attribute VB_Name = "CalculoLetraDNIAux"
Option Explicit

' ===============================================================================
' M�dulo: CalculoLetraDNI
' Funcionalidad:
'   - Verifica y valida la letra del DNI/NIE en la columna I de la fila activa de una hoja.
'   - Si la letra NO coincide con la calculada, ofrece la opci�n de corregirla.
'   - Soporta DNI y NIE (X, Y, Z).
'   - Incluye funciones auxiliares para extracci�n y validaci�n.
' ===============================================================================

Private Const DNILetras As String = "TRWAGMYFPDXBNJZSQVHLCKE"  ' Secuencia oficial de letras del DNI

' ===============================================================================
' (A) VerificarLetraDNIEnHoja
' Verifica la letra del DNI/NIE en la columna I de la fila activa de la hoja especificada.
' Si la letra NO coincide con la calculada, ofrece corregirla con un mensaje al usuario.
' ===============================================================================
Public Sub VerificarLetraDNIEnHoja(ByVal nombreHoja As String)
    Dim ws As Worksheet
    Dim filaActual As Long
    Dim dniCompleto As String
    Dim numeroDNI As String
    Dim letraCorrecta As String
    Dim letraActual As String
    Dim respuesta As VbMsgBoxResult

    ' ---- Intentar obtener la hoja indicada ----
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "La hoja '" & nombreHoja & "' no existe.", vbExclamation, "Error"
        Exit Sub
    End If

    ' ---- Determinar la fila activa y extraer el valor de la columna I (DNI) ----
    filaActual = ws.Application.ActiveCell.Row
    dniCompleto = Trim(ws.Cells(filaActual, "I").Value)

    ' ---- Solo contin�a si hay alg�n valor en la celda ----
    If dniCompleto <> "" Then
        ' ---- Extraer el n�mero (soporta NIE) ----
        numeroDNI = ObtenerNumeroDNI(dniCompleto)
        ' ---- Validar: Deben ser 8 d�gitos num�ricos ----
        If EsNumeroDNIValido(numeroDNI) Then
            ' ---- Calcular la letra que corresponde y comparar con la introducida ----
            letraCorrecta = CalcularLetraDNI(numeroDNI)
            letraActual = ObtenerLetraDNI(dniCompleto)
            If letraActual <> letraCorrecta Then
                ' ---- Si la letra no coincide, ofrecer al usuario la opci�n de corregirla ----
                respuesta = MsgBox( _
                    "Atenci�n: Para el n�mero '" & numeroDNI & "' la letra correcta es '" & letraCorrecta & "'." & vbCrLf & _
                    "Verifique que los d�gitos sean los correctos." & vbCrLf & _
                    "�Desea modificar la letra?", _
                    vbExclamation + vbYesNo, "Letra de DNI incorrecta")
                If respuesta = vbYes Then
                    ' ---- Si el usuario acepta, corregir la letra autom�ticamente ----
                    ws.Cells(filaActual, "I").Value = numeroDNI & letraCorrecta
                End If
                ' ---- Si el usuario pulsa No, no se modifica nada ----
            End If
            ' ---- Si la letra coincide, se da por v�lido, no se hace nada ----
        Else
            ' ---- Si el n�mero no es v�lido, informar al usuario ----
            MsgBox "El n�mero de DNI no es v�lido (debe tener 8 d�gitos y ser num�rico). Por favor, corr�gelo.", vbCritical, "DNI Inv�lido"
        End If
    End If
End Sub

' ===============================================================================
' (B) CalcularLetraDNI
' Calcula la letra del DNI espa�ol a partir del n�mero (8 d�gitos).
' Devuelve la letra correspondiente seg�n la secuencia oficial.
' ===============================================================================
Public Function CalcularLetraDNI(ByVal numeroDNI As String) As String
    Dim num As Long
    On Error Resume Next
    num = CLng(numeroDNI)
    On Error GoTo 0
    If num >= 0 Then
        ' El c�lculo es el resto de dividir el n�mero por 23
        CalcularLetraDNI = Mid(DNILetras, (num Mod 23) + 1, 1)
    Else
        CalcularLetraDNI = ""
    End If
End Function

' ===============================================================================
' (C) ObtenerNumeroDNI
' Extrae solo los n�meros a partir del DNI/NIE completo (ej: "12345678Z" o "X1234567L").
' Convierte NIE inicial (X,Y,Z) a n�mero equivalente (X=0, Y=1, Z=2).
' ===============================================================================
Public Function ObtenerNumeroDNI(ByVal dniCompleto As String) As String
    Dim dni As String
    dni = UCase(dniCompleto)
    ' NIE: convierte letra inicial a n�mero para el c�lculo
    If Left(dni, 1) = "X" Then
        dni = "0" & Mid(dni, 2)
    ElseIf Left(dni, 1) = "Y" Then
        dni = "1" & Mid(dni, 2)
    ElseIf Left(dni, 1) = "Z" Then
        dni = "2" & Mid(dni, 2)
    End If
    ' Devuelve todo menos el �ltimo car�cter (la letra)
    ObtenerNumeroDNI = Left(dni, Len(dni) - 1)
End Function

' ===============================================================================
' (D) ObtenerLetraDNI
' Extrae la letra del DNI/NIE (�ltimo car�cter, convertido a may�scula).
' ===============================================================================
Public Function ObtenerLetraDNI(ByVal dniCompleto As String) As String
    ObtenerLetraDNI = UCase(Right(Trim(dniCompleto), 1))
End Function

' ===============================================================================
' (E) EsNumeroDNIValido
' Comprueba si el n�mero del DNI tiene exactamente 8 d�gitos num�ricos.
' ===============================================================================
Public Function EsNumeroDNIValido(ByVal numeroDNI As String) As Boolean
    EsNumeroDNIValido = (Len(numeroDNI) = 8) And (IsNumeric(numeroDNI))
End Function

' ===============================================================================
' (F) VerificarTodasLasResidencias
' Ejecuta la verificaci�n de la letra del DNI/NIE en las tres hojas principales.
' Muestra un mensaje al finalizar la comprobaci�n.
' ===============================================================================
Public Sub VerificarTodasLasResidencias()
    Call VerificarLetraDNIEnHoja("RESIDENCIA GIJ�N")
    Call VerificarLetraDNIEnHoja("RESIDENCIA SOTO")
    Call VerificarLetraDNIEnHoja("RESIDENCIA OVIEDO")
    MsgBox "Verificaci�n de DNI completada en todas las residencias.", vbInformation
End Sub

