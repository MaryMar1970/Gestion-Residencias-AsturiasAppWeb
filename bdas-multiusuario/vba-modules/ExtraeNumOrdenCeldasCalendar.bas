Attribute VB_Name = "ExtraeNumOrdenCeldasCalendar"
'=================================================================================
' M�dulo: ExtraerNumeroOrden
'
' Descripci�n:
'   Funci�n para extraer el n�mero de orden de una celda de texto, normalmente
'   usada en calendarios donde el N� ORDEN puede aparecer en la segunda l�nea de la celda,
'   y opcionalmente termina con " C" (comisi�n), que se elimina.
'=================================================================================

Option Explicit

'---------------------------------------------------------------------------------
' ExtraerNumeroOrden
'
' Par�metros:
'   - cellText: Texto de la celda del que se quiere extraer el n�mero de orden.
'
' Retorno:
'   - Devuelve el n�mero de orden como string, sin sufijo " C".
'
' L�gica:
'   1. Verifica que el contenido sea de tipo texto.
'   2. Intenta separar el texto por salto de l�nea (vbNewLine).
'   3. Si hay m�s de una l�nea, asume que el n�mero de orden est� en la segunda.
'      Si solo hay una, toma esa.
'   4. Elimina espacios y, si corresponde, el sufijo " C".
'   5. Si el contenido no es texto o est� vac�o, devuelve cadena vac�a.
'---------------------------------------------------------------------------------
Public Function ExtraerNumeroOrden(cellText As Variant) As String
    Dim parts() As String
    Dim orden As String

    ' 1. Verifica que el contenido de la celda sea texto y no est� vac�o
    If VarType(cellText) = vbString And Len(cellText) > 0 Then
        parts = Split(cellText, vbNewLine)
        If UBound(parts) >= 1 Then
            orden = parts(1)
        ElseIf UBound(parts) = 0 Then
            orden = parts(0)
        Else
            ExtraerNumeroOrden = ""
            Exit Function
        End If
        orden = Trim(orden)
        ' 4. Si termina en " C" (comisi�n), elimina ese fragmento
        If Len(orden) >= 2 Then
            If Right(orden, 2) = " C" Then
                orden = Trim(Left(orden, Len(orden) - 2))
            End If
        End If
        ExtraerNumeroOrden = orden
    Else
        ' Si el contenido no es texto o est� vac�o, devuelve cadena vac�a
        ExtraerNumeroOrden = ""
    End If
End Function
