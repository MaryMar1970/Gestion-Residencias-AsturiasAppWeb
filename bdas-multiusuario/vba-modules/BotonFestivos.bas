Attribute VB_Name = "BotonFestivos"
Sub AsignarFestivo(control As IRibbonControl)
    On Error Resume Next
    With ThisWorkbook.Sheets("Festivos")
        .visible = xlSheetVisible
        .Activate
    End With
End Sub

