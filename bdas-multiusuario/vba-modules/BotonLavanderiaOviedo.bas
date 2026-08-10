Attribute VB_Name = "BotonLavanderiaOviedo"
Sub MostrarLavanderiaOviedo(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Lavander�a Oviedo")
    ws.visible = xlSheetVisible
    ws.Activate
End Sub

