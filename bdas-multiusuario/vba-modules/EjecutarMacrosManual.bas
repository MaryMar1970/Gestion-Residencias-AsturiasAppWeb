Attribute VB_Name = "EjecutarMacrosManual"
Sub RestaurarEventosYEstados(control As IRibbonControl)
    Application.enableEvents = True
    Application.calculation = xlCalculationAutomatic
    Application.screenUpdating = True
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    MsgBox "Estados de Excel restaurados. Los eventos y la actualizaci�n est�n activos.", vbInformation
End Sub
