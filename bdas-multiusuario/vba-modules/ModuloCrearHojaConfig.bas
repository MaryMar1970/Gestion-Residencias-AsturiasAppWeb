Attribute VB_Name = "ModuloCrearHojaConfig"
Sub CrearHojaConfiguracion()
    Dim wsConfig As Worksheet
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Worksheets("CONFIG")
    On Error GoTo 0
    
    ' Si ya existe, salir
    If Not wsConfig Is Nothing Then
        MsgBox "La hoja CONFIG ya existe.", vbInformation
        Exit Sub
    End If
    
    ' Crear hoja
    Set wsConfig = ThisWorkbook.Worksheets.Add
    wsConfig.Name = "CONFIG"
    
    ' Encabezados y configuraci�n
    With wsConfig
        .Cells(1, 1).Value = "CONFIGURACI�N DEL SISTEMA"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 14
        
        .Cells(3, 1).Value = "Par�metro"
        .Cells(3, 2).Value = "Valor"
        .Cells(3, 3).Value = "Descripci�n"
        .Range("A3:C3").Font.Bold = True
        .Range("A3:C3").Interior.color = RGB(217, 217, 217)
        
        .Cells(4, 1).Value = "D�as Retenci�n"
        .Cells(4, 2).Value = 90
        .Cells(4, 3).Value = "D�as antes de archivar datos personales"
        
        .Cells(5, 1).Value = "Buscar en BDAS"
        .Cells(5, 2).Value = "S�"
        .Cells(5, 3).Value = "Incluir BDAS en b�squedas (S�/NO)"
        
        ' Ajustar anchos
        .Columns("A:A").ColumnWidth = 20
        .Columns("B:B").ColumnWidth = 15
        .Columns("C:C").ColumnWidth = 50
    End With
    
    ' Ocultar hoja
    wsConfig.visible = xlSheetVeryHidden
    
    ' Proteger
    wsConfig.Protect password:=ModuloConfigSegura.ObtenerPasswordHojas(), UserInterfaceOnly:=True
    
    MsgBox "Hoja CONFIG creada correctamente." & vbCrLf & vbCrLf & _
           "Configuraci�n:" & vbCrLf & _
           "- D�as retenci�n: 90" & vbCrLf & _
           "- Buscar en BDAS: S�", vbInformation, "Configuraci�n Creada"
End Sub

