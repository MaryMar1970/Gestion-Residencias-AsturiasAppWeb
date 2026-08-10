Attribute VB_Name = "TextoProperCaseyUpperCase"
' Funci�n para formatear una celda a ProperCase
Public Sub FormatearTextoProperCase(ByVal celda As Range)
    If Not IsEmpty(celda.Value) Then
        celda.Value = StrConv(celda.Value, vbProperCase)
    End If
End Sub

' Funci�n para formatear una celda a may�sculas
Public Sub FormatearTextoUpperCase(ByVal celda As Range)
    If Not IsEmpty(celda.Value) Then
        celda.Value = UCase(celda.Value)
    End If
End Sub

