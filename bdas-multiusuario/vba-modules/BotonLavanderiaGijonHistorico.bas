Attribute VB_Name = "BotonLavanderiaGijonHistorico"
Sub MostrarHistoricoLavanderiaGijon(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HISTORICO_LAVANDERIA_GIJON")
    ws.visible = xlSheetVisible
    ws.Activate
End Sub


