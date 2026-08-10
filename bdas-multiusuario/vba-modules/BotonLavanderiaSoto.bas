Attribute VB_Name = "BotonLavanderiaSoto"
Sub MostrarLavanderiaSoto(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Lavander�a Soto")
    ws.visible = xlSheetVisible
    ws.Activate
End Sub

