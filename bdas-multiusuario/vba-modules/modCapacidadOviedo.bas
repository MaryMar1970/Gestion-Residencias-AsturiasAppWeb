Attribute VB_Name = "modCapacidadOviedo"
Option Explicit

' =========================================================
' Calcula la capacidad normalizada de una fila (solo OVIEDO)
' =========================================================
Public Function CapacidadTotalOviedo(ByVal habInd As Long, _
                                     ByVal habDob As Long, _
                                     ByVal supletorias As Long) As Long
    ' Individual = 1
    ' Doble = 2
    ' Supletoria = 1
    CapacidadTotalOviedo = habInd _
                         + (habDob * 2) _
                         + supletorias
End Function


' =========================================================
' Valida si el n�mero de supletorias es admisible seg�n
' las habitaciones dobles adjudicadas
' =========================================================
Public Function SupletoriasValidasOviedo(ByVal habDob As Long, _
                                        ByVal supletorias As Long, _
                                        ByVal listaHabitaciones As String) As Boolean
    Dim maxSupletorias As Long
    maxSupletorias = maxSupletoriasPermitidas(listaHabitaciones)

    ' No puede haber supletorias sin dobles
    If supletorias > 0 And habDob = 0 Then
        SupletoriasValidasOviedo = False
        Exit Function
    End If

    SupletoriasValidasOviedo = (supletorias <= maxSupletorias)
End Function

' =========================================================
' Devuelve el m�ximo de camas supletorias permitidas
' seg�n las habitaciones asignadas
' =========================================================
Private Function maxSupletoriasPermitidas(ByVal habitaciones As String) As Long
    Dim arr As Variant, h As Variant
    arr = Split(habitaciones, ",")

    For Each h In arr
        Select Case Trim(h)
            Case "1", "2", "3"
                maxSupletoriasPermitidas = maxSupletoriasPermitidas + 1
            Case "12"
                maxSupletoriasPermitidas = maxSupletoriasPermitidas + 2
        End Select
    Next h
End Function

