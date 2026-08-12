Option Explicit

Private m_Cancelado As Boolean
Private WithEvents btnEntrar As MSForms.CommandButton
Private WithEvents btnSalir As MSForms.CommandButton

Private Sub UserForm_Initialize()
    Dim ctl As Object

    Me.Caption = "BDAS - Inicio de Sesi" & Chr(243) & "n"
    Me.Width = 250
    Me.Height = 210
    Me.StartUpPosition = 1

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    ctl.Caption = "BDAS - Residencias"
    ctl.Left = 10: ctl.Top = 10: ctl.Width = 224: ctl.Height = 24
    ctl.TextAlign = 2
    With ctl.Font: .Size = 14: .Bold = True: End With

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblSubtitulo", True)
    ctl.Caption = "Inicio de Sesi" & Chr(243) & "n"
    ctl.Left = 10: ctl.Top = 34: ctl.Width = 224: ctl.Height = 18
    ctl.TextAlign = 2

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblUsuario", True)
    ctl.Caption = "Usuario:"
    ctl.Left = 18: ctl.Top = 62: ctl.Width = 60: ctl.Height = 18

    Set ctl = Me.Controls.Add("Forms.TextBox.1", "txtUsuario", True)
    ctl.Left = 84: ctl.Top = 60: ctl.Width = 140: ctl.Height = 22
    ctl.TabIndex = 0

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblClave", True)
    ctl.Caption = "Contrase" & Chr(241) & "a:"
    ctl.Left = 18: ctl.Top = 92: ctl.Width = 60: ctl.Height = 18

    Set ctl = Me.Controls.Add("Forms.TextBox.1", "txtClave", True)
    ctl.Left = 84: ctl.Top = 90: ctl.Width = 140: ctl.Height = 22
    ctl.PasswordChar = "*"
    ctl.TabIndex = 1

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblError", True)
    ctl.Caption = ""
    ctl.Left = 18: ctl.Top = 118: ctl.Width = 206: ctl.Height = 18
    ctl.ForeColor = RGB(200, 0, 0)
    ctl.Visible = False

    Set btnEntrar = Me.Controls.Add("Forms.CommandButton.1", "btnEntrar", True)
    btnEntrar.Caption = "Entrar"
    btnEntrar.Left = 36: btnEntrar.Top = 142: btnEntrar.Width = 80: btnEntrar.Height = 26
    btnEntrar.Default = True
    btnEntrar.TabIndex = 2

    Set btnSalir = Me.Controls.Add("Forms.CommandButton.1", "btnSalir", True)
    btnSalir.Caption = "Salir"
    btnSalir.Left = 130: btnSalir.Top = 142: btnSalir.Width = 80: btnSalir.Height = 26
    btnSalir.Cancel = True
    btnSalir.TabIndex = 3

    m_Cancelado = False
    Me.Controls("txtUsuario").SetFocus
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
    Set txU  = Me.Controls("txtUsuario")
    Set txC  = Me.Controls("txtClave")

    usuario = Trim(txU.Text)
    clave   = txC.Text

    lblE.Visible = False

    If Len(usuario) = 0 Then
        lblE.Caption = "Introduce tu nombre de usuario."
        lblE.Visible = True
        txU.SetFocus
        Exit Sub
    End If
    If Len(clave) = 0 Then
        lblE.Caption = "Introduce tu contrase" & Chr(241) & "a."
        lblE.Visible = True
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
        lblE.Visible = True
        txC.Text = ""
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
