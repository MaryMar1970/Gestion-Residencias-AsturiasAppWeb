Attribute VB_Name = "BotonTarifasUnidadDescanso"
' Muestra la hoja TARIFAS UNIDAD
Sub MostrarTarifaUnidad(control As IRibbonControl)
    On Error Resume Next
    With ThisWorkbook.Sheets("TARIFAS UNIDAD")
        .visible = xlSheetVisible
        .Activate
    End With
End Sub

' Muestra la hoja Turnos-Precios
Sub MostrarTarifaDescanso(control As IRibbonControl)
    On Error Resume Next
    With ThisWorkbook.Sheets("Turnos-Precios")
        .visible = xlSheetVisible
        .Activate
    End With
End Sub
