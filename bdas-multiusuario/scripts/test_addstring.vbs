Dim fso, f, codeText, excel, wb, proj, comp

Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile("H:\ResidenciaApp\bdas-multiusuario\scripts\modDatabase.bas", 1)
codeText = f.ReadAll()
f.Close()

Set excel = CreateObject("Excel.Application")
excel.Visible = True
excel.WindowState = -4137
excel.DisplayAlerts = False

Set wb = excel.Workbooks.Add()
Set proj = wb.VBProject

On Error Resume Next
Set comp = proj.VBComponents.Item("modDatabase")
If Not comp Is Nothing Then proj.VBComponents.Remove comp
On Error GoTo 0

WScript.Echo "Creating new standard module modDatabase..."
Set comp = proj.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
comp.Name = "modDatabase"

WScript.Echo "Adding code from string..."
On Error Resume Next
comp.CodeModule.AddFromString codeText
If Err.Number <> 0 Then
    WScript.Echo "ERROR AddFromString: " & Err.Description & " (Code: " & Hex(Err.Number) & ")"
Else
    WScript.Echo "SUCCESS: Code added to " & comp.Name & "! Lines: " & comp.CodeModule.CountOfLines
End If

wb.Close False
excel.Quit
