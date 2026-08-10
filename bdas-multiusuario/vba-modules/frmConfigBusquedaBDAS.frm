VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConfigBusquedaBDAS 
   Caption         =   "Configuraci�n b�squeda DNI en BDAS"
   ClientHeight    =   2370
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3510
   OleObjectBlob   =   "frmConfigBusquedaBDAS.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmConfigBusquedaBDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAplicar_Click()
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    If wsConfig Is Nothing Then
        MsgBox "No se encuentra la hoja CONFIG.", vbExclamation
    Else
        If Me.chkActivarBusquedaBDAS.Value = True Then
            wsConfig.Range("B6").Value = "SI"
        Else
            wsConfig.Range("B6").Value = "NO"
        End If
        MsgBox "Cambios aplicados correctamente.", vbInformation
        Unload Me
    End If
End Sub

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    If wsConfig Is Nothing Then
        Me.chkActivarBusquedaBDAS.Enabled = False
        MsgBox "No se encuentra la hoja CONFIG.", vbExclamation
    Else
        Dim valorB6 As String
        valorB6 = UCase(Trim(wsConfig.Range("B6").Value))
        Me.chkActivarBusquedaBDAS.Value = (valorB6 = "SI" Or valorB6 = "S�")
    End If
End Sub
