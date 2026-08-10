VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResultadosOviedo 
   Caption         =   "UserForm1"
   ClientHeight    =   7680
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   13965
   OleObjectBlob   =   "ResultadosOviedo.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ResultadosOviedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --- C�digo en frmResultados ---

Private Sub lResultadosOviedo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim filePath As String
    filePath = gFolderPathOviedo & "\" & lResultadosOviedo.Value
    If filePath <> "" Then
        Shell "cmd /c start """" """ & filePath & """", vbNormalFocus
    End If
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

