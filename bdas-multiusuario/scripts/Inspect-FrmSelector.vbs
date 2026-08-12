Dim xl, wb, comp
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
xl.EnableEvents = False
Set wb = xl.Workbooks.Open("H:\ResidenciaBD\BDAS_Multiusuario.xlsm", 0, True)
Set comp = wb.VBProject.VBComponents("frmSelectorResidencia")
WScript.Echo "Type: " & comp.Type
WScript.Echo "Code:"
WScript.Echo comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
wb.Close False
xl.Quit
