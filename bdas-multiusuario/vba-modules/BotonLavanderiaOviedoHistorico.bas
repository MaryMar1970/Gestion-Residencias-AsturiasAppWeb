Attribute VB_Name = "BotonLavanderiaOviedoHistorico"
Sub MostrarHistoricoLavanderiaoviedo(control As IRibbonControl)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("HISTORICO_LAVANDERIA_OVIEDO")
    ws.visible = xlSheetVisible
    ws.Activate
End Sub
