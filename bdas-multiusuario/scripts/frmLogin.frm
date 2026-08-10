Option Explicit

Private m_Cancelado As Boolean

Private Sub UserForm_Initialize()
    Dim ctl As Object

    Me.Caption = "BDAS" & Chr(160) & Chr(8212) & Chr(160) & "Inicio de Sesi" & Chr(243) & "n"
    Me.Width = 220
    Me.Height = 170
    Me.StartUpPosition = 1

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblTitulo", True)
    ctl.Caption = "BDAS" & Chr(160) & Chr(8212) & Chr(160) & "Residencias"
    ctl.Left = 6: ctl.Top = 6: ctl.Width = 198: ctl.Height = 24
    ctl.TextAlign = 2
    With ctl.Font: .Size = 14: .Bold = True: End With

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblSubtitulo", True)
    ctl.Caption = "Inicio de Sesi" & Chr(243) & "n"
    ctl.Left = 6: ctl.Top = 33: ctl.Width = 198: ctl.Height = 18
    ctl.TextAlign = 2

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblUsuario", True)
    ctl.Caption = "Usuario:"
    ctl.Left = 18: ctl.Top = 63: ctl.Width = 54: ctl.Height = 18

    Set ctl = Me.Controls.Add("Forms.TextBox.1", "txtUsuario", True)
    ctl.Left = 78: ctl.Top = 60: ctl.Width = 126: ctl.Height = 21
    ctl.TabIndex = 0

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblClave", True)
    ctl.Caption = "Contrase" & Chr(241) & "a:"
    ctl.Left = 18: ctl.Top = 90: ctl.Width = 54: ctl.Height = 18

    Set ctl = Me.Controls.Add("Forms.TextBox.1", "txtClave", True)
    ctl.Left = 78: ctl.Top = 87: ctl.Width = 126: ctl.Height = 21
    ctl.PasswordChar = "*"
    ctl.TabIndex = 1

    Set ctl = Me.Controls.Add("Forms.Label.1", "lblError", True)
    ctl.Caption = ""
    ctl.Left = 18: ctl.Top = 117: ctl.Width = 186: ctl.Height = 18
    ctl.ForeColor = RGB(200, 0, 0)
    ctl.Visible = False

    Set ctl = Me.Controls.Add("Forms.CommandButton.1", "btnEntrar", True)
    ctl.Caption = "Entrar"
    ctl.Left = 36: ctl.Top = 141: ctl.Width = 72: ctl.Height = 24
    ctl.Default = True
    ctl.TabIndex = 2

    Set ctl = Me.Controls.Add("Forms.CommandButton.1", "btnSalir", True)
    ctl.Caption = "Salir"
    ctl.Left = 120: ctl.Top = 141: ctl.Width = 72: ctl.Height = 24
    ctl.Cancel = True
    ctl.TabIndex = 3

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
