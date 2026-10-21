VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "BDAS - Inicio de Sesion"
   ClientHeight    =   2820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4155
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Cancelado As Boolean
Private WithEvents txtUsuario As MSForms.TextBox
Private WithEvents txtClave As MSForms.TextBox
Private WithEvents btnEntrar As MSForms.CommandButton
Private WithEvents btnSalir As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim ctl As Object

    ' Configuración de tarjeta compacta integrada en la escena
    Me.Caption = "BDAS - Identificaci" & Chr(243) & "n"
    Me.Width = 240
    Me.Height = 165
    Me.BackColor = RGB(255, 255, 255)
    Me.StartUpPosition = 0 ' Posicionamiento manual exacto

    ' Título institucional estilizado
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    ctl.Caption = "BDAS - Acceso al Sistema"
    ctl.Left = 10: ctl.Top = 10: ctl.Width = 220: ctl.Height = 18
    ctl.TextAlign = 2
    ctl.BackStyle = 0
    ctl.ForeColor = RGB(30, 45, 60)
    With ctl.Font: .Size = 10: .Bold = True: End With

    ' Etiqueta y campo de Usuario
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblUsuario", True)
    ctl.Caption = "Usuario:"
    ctl.Left = 18: ctl.Top = 38: ctl.Width = 55: ctl.Height = 16
    ctl.BackStyle = 0
    ctl.ForeColor = RGB(60, 60, 60)
    With ctl.Font: .Size = 9: .Bold = True: End With

    Set txtUsuario = Me.Controls.Add("Forms.TextBox.1", "txtUsuario", True)
    txtUsuario.Left = 78: txtUsuario.Top = 36: txtUsuario.Width = 140: txtUsuario.Height = 20
    txtUsuario.SpecialEffect = 2
    txtUsuario.TabIndex = 0
    txtUsuario.TabStop = True

    ' Etiqueta y campo de Contraseña
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblClave", True)
    ctl.Caption = "Clave:"
    ctl.Left = 18: ctl.Top = 66: ctl.Width = 55: ctl.Height = 16
    ctl.BackStyle = 0
    ctl.ForeColor = RGB(60, 60, 60)
    With ctl.Font: .Size = 9: .Bold = True: End With

    Set txtClave = Me.Controls.Add("Forms.TextBox.1", "txtClave", True)
    txtClave.Left = 78: txtClave.Top = 64: txtClave.Width = 140: txtClave.Height = 20
    txtClave.PasswordChar = "*"
    txtClave.SpecialEffect = 2
    txtClave.TabIndex = 1
    txtClave.TabStop = True

    ' Mensaje de error
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblError", True)
    ctl.Caption = ""
    ctl.Left = 12: ctl.Top = 88: ctl.Width = 216: ctl.Height = 16
    ctl.TextAlign = 2
    ctl.BackStyle = 0
    ctl.ForeColor = RGB(200, 0, 0)
    With ctl.Font: .Size = 8: .Bold = False: End With
    ctl.visible = False

    ' Botón Entrar (verde institucional sobrio con texto blanco)
    Set btnEntrar = Me.Controls.Add("Forms.CommandButton.1", "btnEntrar", True)
    btnEntrar.Caption = "Entrar"
    btnEntrar.Left = 32: btnEntrar.Top = 108: btnEntrar.Width = 82: btnEntrar.Height = 24
    btnEntrar.BackColor = RGB(35, 75, 45)
    btnEntrar.ForeColor = RGB(255, 255, 255)
    With btnEntrar.Font: .Size = 9: .Bold = True: End With
    btnEntrar.Default = False
    btnEntrar.TabIndex = 2
    btnEntrar.TabStop = True

    ' Botón Salir
    Set btnSalir = Me.Controls.Add("Forms.CommandButton.1", "btnSalir", True)
    btnSalir.Caption = "Salir"
    btnSalir.Left = 126: btnSalir.Top = 108: btnSalir.Width = 82: btnSalir.Height = 24
    btnSalir.BackColor = RGB(242, 242, 245)
    btnSalir.ForeColor = RGB(50, 50, 50)
    With btnSalir.Font: .Size = 9: End With
    btnSalir.Cancel = True
    btnSalir.TabIndex = 3
    btnSalir.TabStop = True

    m_Cancelado = False
    
    ' Calcular posición inicial anclada
    PosicionarFormulario
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    PosicionarFormulario
    txtUsuario.SetFocus
    On Error GoTo 0
End Sub

' ==============================================================================
' Anclaje milimétrico del formulario en la base entre los dos Guardias Civiles
' ==============================================================================
Private Sub PosicionarFormulario()
    On Error Resume Next
    
    ' 1. Centrado horizontal exacto con la ventana de Excel (eje central de la fotografía)
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    
    ' 2. Posicionamiento vertical: en la base de la escena (52% de la altura útil de la hoja)
    Dim topGrilla As Double
    topGrilla = Application.Top + (Application.Height - Application.UsableHeight)
    Me.Top = topGrilla + (Application.UsableHeight * 0.52)
    
    ' Salvaguarda para evitar solapes con la barra de tareas en resoluciones muy compactas
    If (Me.Top + Me.Height) > (Application.Top + Application.Height - 35) Then
        Me.Top = Application.Top + Application.Height - Me.Height - 45
    End If
    If Me.Top < (topGrilla + 10) Then Me.Top = topGrilla + 10
    
    On Error GoTo 0
End Sub

Private Sub txtUsuario_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        txtClave.SetFocus
    End If
End Sub

Private Sub txtClave_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        ProcesarLogin
    End If
End Sub


Private Sub btnEntrar_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        ProcesarLogin
    End If
End Sub

Private Sub btnEntrar_Click()
    ProcesarLogin
End Sub

Private Sub ProcesarLogin()
    Dim usuario As String
    Dim clave As String
    Dim btnE As Object, lblE As Object, txU As Object, txC As Object

    Set btnE = Me.Controls("btnEntrar")
    Set lblE = Me.Controls("lblError")
    Set txU = Me.Controls("txtUsuario")
    Set txC = Me.Controls("txtClave")

    usuario = Trim(txU.text)
    clave = txC.text

    lblE.visible = False

    If Len(usuario) = 0 Then
        lblE.Caption = "Introduce tu nombre de usuario."
        lblE.visible = True
        txU.SetFocus
        Exit Sub
    End If
    If Len(clave) = 0 Then
        lblE.Caption = "Introduce tu contrase" & Chr(241) & "a."
        lblE.visible = True
        txC.SetFocus
        Exit Sub
    End If

    btnE.Enabled = False
    btnE.Caption = "Validando..."
    DoEvents

    If modDatabase.ValidarCredenciales(usuario, clave) Then
        Me.Hide
    Else
        lblE.Caption = "Usuario o contrase" & Chr(241) & "a incorrectos."
        lblE.visible = True
        txC.text = ""
        txC.SetFocus
    End If

    btnE.Enabled = True
    btnE.Caption = "Entrar"
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

Public Property Get Cancelado() As Boolean
    Cancelado = m_Cancelado
End Property
