Attribute VB_Name = "BotonSobreOviedo"
Sub MostrarSobreOviedo(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SOBRE OVIEDO")
    ws.visible = xlSheetVisible
    ws.Activate
End Sub
