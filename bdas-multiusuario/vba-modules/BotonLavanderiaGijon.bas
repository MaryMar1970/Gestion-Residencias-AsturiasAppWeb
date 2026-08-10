Attribute VB_Name = "BotonLavanderiaGijon"
Sub MostrarLavanderiaGijon(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Lavander�a Gij�n")
    ws.visible = xlSheetVisible
    ws.Activate
End Sub


