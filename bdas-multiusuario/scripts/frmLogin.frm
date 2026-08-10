VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin
   Caption         =   "BDAS - Inicio de Sesi" & Chr(243) & "n"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblTitulo
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   4920
      Caption         =   "BDAS - Residencias"
      Font.Bold       =   -1
      Font.Size       =   14
      TextAlign       =   2
   End
   Begin MSForms.Label lblSubtitulo
      Height          =   360
      Left            =   240
      Top             =   720
      Width           =   4920
      Caption         =   "Inicio de Sesi" & Chr(243) & "n"
      Font.Size       =   10
      TextAlign       =   2
   End
   Begin MSForms.Label lblUsuario
      Height          =   240
      Left            =   480
      Top             =   1320
      Width           =   1440
      Caption         =   "Usuario:"
   End
   Begin MSForms.TextBox txtUsuario
      Height          =   360
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   2880
   End
   Begin MSForms.Label lblClave
      Height          =   240
      Left            =   480
      Top             =   1920
      Width           =   1440
      Caption         =   "Contrase" & Chr(241) & "a:"
   End
   Begin MSForms.TextBox txtClave
      Height          =   360
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   2880
   End
   Begin MSForms.Label lblError
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   480
      Top             =   2520
      Visible         =   0
      Width           =   4440
   End
   Begin MSForms.CommandButton btnEntrar
      Caption         =   "Entrar"
      Default         =   -1
      Height          =   480
      Left            =   1200
      TabIndex        =   2
      Top             =   3240
      Width           =   1440
   End
   Begin MSForms.CommandButton btnSalir
      Cancel          =   -1
      Caption         =   "Salir"
      Height          =   480
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   1440
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Cancelado As Boolean

Private Sub UserForm_Initialize()
    txtUsuario.Text = ""
    txtClave.Text = ""
    lblError.Visible = False
    m_Cancelado = False
    txtUsuario.SetFocus
End Sub

Private Sub btnEntrar_Click()
    Dim usuario As String
    Dim clave As String
    
    usuario = Trim(txtUsuario.Text)
    clave = txtClave.Text
    
    If Len(usuario) = 0 Then
        MostrarError "Introduce tu nombre de usuario."
        txtUsuario.SetFocus
        Exit Sub
    End If
    
    If Len(clave) = 0 Then
        MostrarError "Introduce tu contrase" & Chr(241) & "a."
        txtClave.SetFocus
        Exit Sub
    End If
    
    btnEntrar.Enabled = False
    btnEntrar.Caption = "Validando..."
    DoEvents
    
    If modDatabase.ValidarCredenciales(usuario, clave) Then
        Me.Hide
    Else
        MostrarError "Usuario o contrase" & Chr(241) & "a incorrectos."
        txtClave.Text = ""
        txtClave.SetFocus
    End If
    
    btnEntrar.Enabled = True
    btnEntrar.Caption = "Entrar"
End Sub

Private Sub btnSalir_Click()
    m_Cancelado = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        m_Cancelado = True
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub MostrarError(ByVal mensaje As String)
    lblError.Caption = mensaje
    lblError.Visible = True
End Sub

Public Property Get Cancelado() As Boolean
    Cancelado = m_Cancelado
End Property

Private Sub txtUsuario_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lblError.Visible = False
End Sub

Private Sub txtClave_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lblError.Visible = False
End Sub
