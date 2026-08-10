Attribute VB_Name = "BotonSobreGijon"
Sub MostrarSobreGijon(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SOBRE GIJ�N")
    ws.visible = xlSheetVisible
    ws.Activate
End Sub
