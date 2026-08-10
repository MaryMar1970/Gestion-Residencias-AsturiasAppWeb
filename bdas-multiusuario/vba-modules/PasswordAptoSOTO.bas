Attribute VB_Name = "PasswordAptoSOTO"
Option Explicit

'=================================================================================
' Funci�n: ObtenerContrasenasSoto
' Descripci�n:
'   Dada una cadena con identificadores de apartamentos (por ejemplo "Ap.1,Ap.2,Ap.7"),
'   busca la contrase�a asociada a cada uno en la hoja "CONTRASE�AS SOTO" (columna B)
'   y devuelve una cadena con las contrase�as separadas por comas, en el mismo orden.
'   Si un identificador no se encuentra, devuelve "#N/D" en su lugar.
'
' Par�metro:
'   strHabitaciones (String): Apartamentos separados por comas ("Ap.1,Ap.2,...")
'
' Devuelve:
'   String: Contrase�as separadas por comas en el mismo orden de entrada.
'
' Ejemplo:
'   Si "Ap.1" ? 1234, "Ap.2" ? 5678, "Ap.3" no existe, entonces:
'   ObtenerContrasenasSoto("Ap.1,Ap.3,Ap.2") => "1234,#N/D,5678"
'=================================================================================
Public Function ObtenerContrasenasSoto(ByVal strHabitaciones As String) As String
    Dim arrHab() As String            ' Array para almacenar los apartamentos de entrada
    Dim arrPwd() As String            ' Array para almacenar las contrase�as obtenidas
    Dim shtPWD As Worksheet           ' Referencia a la hoja de contrase�as
    Dim ultimaFila As Long            ' �ltima fila con datos en la columna A
    Dim i As Long                     ' �ndice para iterar arrays
    Dim claveBusqueda As String        ' Apartamento a buscar en cada iteraci�n

    '---------------------------------------------------------------------
    ' 1. Referencia a la hoja donde est�n los datos de contrase�as
    '---------------------------------------------------------------------
    Set shtPWD = ThisWorkbook.Worksheets("CONTRASE�AS SOTO")
    
    '---------------------------------------------------------------------
    ' 2. Limpieza de la entrada: quita espacios antes/despu�s de la cadena
    '    Si est� vac�a, devuelve resultado vac�o
    '---------------------------------------------------------------------
    strHabitaciones = Trim(strHabitaciones)
    If strHabitaciones = vbNullString Then
        ObtenerContrasenasSoto = ""
        Exit Function
    End If

    '---------------------------------------------------------------------
    ' 3. Divide la cadena en un array usando la coma como separador
    '    Ejemplo: "Ap.1,Ap.2" ? arrHab(0)="Ap.1", arrHab(1)="Ap.2"
    '---------------------------------------------------------------------
    arrHab = Split(strHabitaciones, ",")
    
    '---------------------------------------------------------------------
    ' 4. Prepara el array de contrase�as, mismo rango de �ndices que arrHab
    '---------------------------------------------------------------------
    ReDim arrPwd(LBound(arrHab) To UBound(arrHab))

    '---------------------------------------------------------------------
    ' 5. Calcula la �ltima fila con datos en la columna A, para delimitar el rango de b�squeda
    '---------------------------------------------------------------------
    ultimaFila = shtPWD.Cells(shtPWD.Rows.Count, "A").End(xlUp).Row

    '---------------------------------------------------------------------
    ' 6. Para cada apartamento solicitado:
    '    - Limpia espacios
    '    - Busca la contrase�a en la hoja (columna B) usando VLOOKUP por texto exacto
    '    - Si hay error (no encuentra o error de tipo), asigna "#N/D"
    '---------------------------------------------------------------------
    For i = LBound(arrHab) To UBound(arrHab)
        On Error Resume Next
        claveBusqueda = Trim(arrHab(i))
        arrPwd(i) = Application.WorksheetFunction. _
                    VLookup(claveBusqueda, _
                            shtPWD.Range("A2:B" & ultimaFila), _
                            2, False)
        If Err.Number <> 0 Then
            arrPwd(i) = "#N/D"   ' Si no encuentra la clave, marca como no disponible
            Err.Clear
        End If
        On Error GoTo 0
    Next i

    '---------------------------------------------------------------------
    ' 7. Devuelve el resultado: todas las contrase�as, separadas por comas
    '---------------------------------------------------------------------
    ObtenerContrasenasSoto = Join(arrPwd, ",")
End Function

