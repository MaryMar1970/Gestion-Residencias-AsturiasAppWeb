Attribute VB_Name = "BotonListaNegra"
Sub MostrarListaNegra(control As IRibbonControl)
    On Error Resume Next
    With ThisWorkbook.Sheets("LISTA NEGRA")
        .visible = xlSheetVisible
        .Activate
    End With
End Sub
