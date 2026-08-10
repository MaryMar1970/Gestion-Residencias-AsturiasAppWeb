Attribute VB_Name = "AccesoPlantillas"
Option Explicit

' Callback para mostrar la plantilla de GIJ�N
Public Sub MostrarPlantillaGijon(control As IRibbonControl)
    MostrarPlantilla "PLANTILLAS GIJ�N"
End Sub

' Callback para mostrar la plantilla de OVIEDO
Public Sub MostrarPlantillaOviedo(control As IRibbonControl)
    MostrarPlantilla "PLANTILLAS OVIEDO"
End Sub

' Callback para mostrar la plantilla de SOTO
Public Sub MostrarPlantillaSoto(control As IRibbonControl)
    MostrarPlantilla "PLANTILLAS SOTO"
End Sub

' --- Funci�n auxiliar ---
Private Sub MostrarPlantilla(nombreHoja As String)
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No se encontr� la hoja: " & nombreHoja, vbExclamation
        Exit Sub
    End If
    
    Application.screenUpdating = False
    ws.visible = xlSheetVisible
    ws.Activate
    Application.screenUpdating = True
End Sub

