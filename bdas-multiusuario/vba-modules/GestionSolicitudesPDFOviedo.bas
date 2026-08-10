Attribute VB_Name = "GestionSolicitudesPDFOviedo"
Option Explicit

' Constantes para almacenar el setting (estos valores se usar�n en el registro de Office)
Public Const APP_NAME As String = "MiBusquedaPDFOviedo"
Public Const KEY_FOLDER As String = "CarpetaPDFOviedo"

Public gFolderPathOviedo As String

' Macro para configurar la carpeta de b�squeda
Sub ConfigurarCarpetaOviedo(control As IRibbonControl)
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Seleccione la carpeta de b�squeda de archivos PDF"
        If .Show = -1 Then
            gFolderPathOviedo = .SelectedItems(1)
            ' Guardar la carpeta en el registro
            SaveSetting APP_NAME, KEY_FOLDER, "Path", gFolderPathOviedo
            MsgBox "Carpeta configurada:" & vbCrLf & gFolderPathOviedo, vbInformation
        Else
            MsgBox "No se seleccion� ninguna carpeta.", vbExclamation
        End If
    End With
End Sub

' Macro para buscar archivos PDF seg�n el nombre ingresado
Sub BuscarPDFOviedo(control As IRibbonControl)
    ' Recupera la carpeta configurada (si existe)
    gFolderPathOviedo = GetSetting(APP_NAME, KEY_FOLDER, "Path", "")
    If gFolderPathOviedo = "" Then
        MsgBox "Primero debe configurar la carpeta de b�squeda.", vbExclamation
        Exit Sub
    End If
    
    Dim nombreBusqueda As String
    nombreBusqueda = InputBox("Ingrese el nombre del archivo PDF a buscar (ejemplo: 1234 o 12345678A):", "Buscar PDF Oviedo")
    If nombreBusqueda = "" Then Exit Sub
    
    ' Validar que el nombre ingresado cumpla el patr�n:
    ' - 4 d�gitos (ejemplo: 1234) o
    ' - 8 d�gitos seguidos de una letra (ejemplo: 12345678A)
    Dim validInput As Boolean
    validInput = False
    If Len(nombreBusqueda) = 4 Then
        If IsNumeric(nombreBusqueda) Then validInput = True
    ElseIf Len(nombreBusqueda) = 9 Then
        If IsNumeric(Left(nombreBusqueda, 8)) And (Mid(nombreBusqueda, 9, 1) Like "[A-Za-z]") Then
            validInput = True
        End If
    End If
    
    If Not validInput Then
        MsgBox "El formato del nombre no es correcto." & vbCrLf & _
               "Debe ser de 4 d�gitos o 8 d�gitos seguidos de una letra.", vbExclamation
        Exit Sub
    End If
    
    ' Buscar archivos PDF en la carpeta configurada cuyo nombre (sin extensi�n) coincida exactamente
    Dim fName As String
    Dim resultados As Collection
    Set resultados = New Collection
    fName = Dir(gFolderPathOviedo & "\*.pdf")
    
    Do While fName <> ""
        Dim baseName As String
        baseName = Left(fName, Len(fName) - 4)   ' Quita la extensi�n ".pdf"
        
        If InStr(1, baseName, nombreBusqueda, vbTextCompare) > 0 Then
            resultados.Add fName
        End If
        
        fName = Dir
    Loop
    
    If resultados.Count > 0 Then
        ' Cargar los resultados en el UserForm y mostrarlo
        ResultadosOviedo.lResultadosOviedo.Clear
        Dim i As Long
        For i = 1 To resultados.Count
            ResultadosOviedo.lResultadosOviedo.AddItem resultados(i)
        Next i
        ResultadosOviedo.Show
    Else
        MsgBox "No se encontraron archivos que coincidan con: " & nombreBusqueda, vbInformation
    End If
End Sub


