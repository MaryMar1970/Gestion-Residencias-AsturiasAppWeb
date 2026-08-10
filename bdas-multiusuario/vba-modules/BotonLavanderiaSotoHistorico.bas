Attribute VB_Name = "BotonLavanderiaSotoHistorico"
Sub MostrarHistoricoLavanderiaSoto(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HISTORICO_LAVANDERIA_SOTO")
    ws.visible = xlSheetVisible
    ws.Activate
End Sub


