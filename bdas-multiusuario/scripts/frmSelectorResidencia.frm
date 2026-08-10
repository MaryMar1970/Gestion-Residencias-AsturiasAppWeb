VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectorResidencia
   Caption         =   "BDAS - Selecci" & Chr(243) & "n de Residencia"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblBienvenida
      Height          =   360
      Left            =   240
      Top             =   240
      Width           =   4320
      Caption         =   "Bienvenido/a"
      Font.Bold       =   -1
      Font.Size       =   12
      TextAlign       =   2
   End
   Begin MSForms.Label lblInstruccion
      Height          =   360
      Left            =   240
      Top             =   720
      Width           =   4320
      Caption         =   "Selecciona la residencia con la que vas a trabajar:"
      TextAlign       =   2
   End
   Begin MSForms.ListBox lstResidencias
      Height          =   1560
      Left            =   480
      ListStyle       =   1
      TabIndex        =   0
      Top             =   1200
      Width           =   3840
   End
   Begin MSForms.CommandButton btnEntrar
      Caption         =   "Entrar"
      Default         =   -1
      Height          =   480
      Left            =   1680
      TabIndex        =   1
      Top             =   3000
      Width           =   1440
   End
End
Attribute VB_Name = "frmSelectorResidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim residencias As Collection
    Dim i As Long
    Dim nombreDisplay As String
    
    lblBienvenida.Caption = "Bienvenido/a, " & modDatabase.ObtenerNombreCompleto()
    
    lstResidencias.Clear
    Set residencias = modDatabase.ObtenerResidenciasAsignadas()
    
    For i = 1 To residencias.Count
        Select Case residencias(i)
            Case "GIJON": nombreDisplay = "Residencia de GIJ" & Chr(211) & "N"
            Case "SOTO": nombreDisplay = "Residencia de SOTO DEL BARCO"
            Case "OVIEDO": nombreDisplay = "Residencia de OVIEDO"
            Case Else: nombreDisplay = residencias(i)
        End Select
        lstResidencias.AddItem nombreDisplay
    Next i
    
    If lstResidencias.ListCount > 0 Then
        lstResidencias.ListIndex = 0
    End If
End Sub

Private Sub btnEntrar_Click()
    Dim residencias As Collection
    Dim selIndex As Long
    
    If lstResidencias.ListIndex < 0 Then
        MsgBox "Selecciona una residencia antes de continuar.", vbExclamation
        Exit Sub
    End If
    
    Set residencias = modDatabase.ObtenerResidenciasAsignadas()
    selIndex = lstResidencias.ListIndex + 1
    
    modDatabase.EstablecerResidenciaActiva residencias(selIndex)
    
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub lstResidencias_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnEntrar_Click
End Sub
