Attribute VB_Name = "modCrearFormularios"
Option Explicit
' ==============================================================================
' Modulo auxiliar TEMPORAL.
' ImportarVBA.vbs lo inyecta, ejecuta CrearFormularios, y luego lo elimina.
' El codigo de los formularios esta incrustado aqui directamente (sin leer
' ficheros) para evitar bloqueos en la llamada COM de VBScript.
' ==============================================================================

Private Function CodigoFrmLogin() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private m_Cancelado As Boolean" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Dim ctl As Object" & vbCrLf
    c = c & "    Me.Caption = ""BDAS - Inicio de Sesi"" & Chr(243) & ""n""" & vbCrLf
    c = c & "    Me.Width = 220" & vbCrLf
    c = c & "    Me.Height = 170" & vbCrLf
    c = c & "    Me.StartUpPosition = 1" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblTitulo"", True)" & vbCrLf
    c = c & "    ctl.Caption = ""BDAS - Residencias""" & vbCrLf
    c = c & "    ctl.Left = 6: ctl.Top = 6: ctl.Width = 198: ctl.Height = 24" & vbCrLf
    c = c & "    ctl.TextAlign = 2" & vbCrLf
    c = c & "    With ctl.Font: .Size = 14: .Bold = True: End With" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblSubtitulo"", True)" & vbCrLf
    c = c & "    ctl.Caption = ""Inicio de Sesi"" & Chr(243) & ""n""" & vbCrLf
    c = c & "    ctl.Left = 6: ctl.Top = 33: ctl.Width = 198: ctl.Height = 18" & vbCrLf
    c = c & "    ctl.TextAlign = 2" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblUsuario"", True)" & vbCrLf
    c = c & "    ctl.Caption = ""Usuario:""" & vbCrLf
    c = c & "    ctl.Left = 18: ctl.Top = 63: ctl.Width = 54: ctl.Height = 18" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.TextBox.1"", ""txtUsuario"", True)" & vbCrLf
    c = c & "    ctl.Left = 78: ctl.Top = 60: ctl.Width = 126: ctl.Height = 21" & vbCrLf
    c = c & "    ctl.TabIndex = 0" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblClave"", True)" & vbCrLf
    c = c & "    ctl.Caption = ""Contrase"" & Chr(241) & ""a:""" & vbCrLf
    c = c & "    ctl.Left = 18: ctl.Top = 90: ctl.Width = 54: ctl.Height = 18" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.TextBox.1"", ""txtClave"", True)" & vbCrLf
    c = c & "    ctl.Left = 78: ctl.Top = 87: ctl.Width = 126: ctl.Height = 21" & vbCrLf
    c = c & "    ctl.PasswordChar = ""*""" & vbCrLf
    c = c & "    ctl.TabIndex = 1" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblError"", True)" & vbCrLf
    c = c & "    ctl.Caption = """"" & vbCrLf
    c = c & "    ctl.Left = 18: ctl.Top = 117: ctl.Width = 186: ctl.Height = 18" & vbCrLf
    c = c & "    ctl.ForeColor = RGB(200, 0, 0)" & vbCrLf
    c = c & "    ctl.Visible = False" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.CommandButton.1"", ""btnEntrar"", True)" & vbCrLf
    c = c & "    ctl.Caption = ""Entrar""" & vbCrLf
    c = c & "    ctl.Left = 36: ctl.Top = 141: ctl.Width = 72: ctl.Height = 24" & vbCrLf
    c = c & "    ctl.Default = True: ctl.TabIndex = 2" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.CommandButton.1"", ""btnSalir"", True)" & vbCrLf
    c = c & "    ctl.Caption = ""Salir""" & vbCrLf
    c = c & "    ctl.Left = 120: ctl.Top = 141: ctl.Width = 72: ctl.Height = 24" & vbCrLf
    c = c & "    ctl.Cancel = True: ctl.TabIndex = 3" & vbCrLf
    c = c & "    m_Cancelado = False" & vbCrLf
    c = c & "    Me.Controls(""txtUsuario"").SetFocus" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub btnEntrar_Click()" & vbCrLf
    c = c & "    ProcesarLogin" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub ProcesarLogin()" & vbCrLf
    c = c & "    Dim usuario As String, clave As String" & vbCrLf
    c = c & "    Dim btnE As Object, lblE As Object, txU As Object, txC As Object" & vbCrLf
    c = c & "    Set btnE = Me.Controls(""btnEntrar"")" & vbCrLf
    c = c & "    Set lblE = Me.Controls(""lblError"")" & vbCrLf
    c = c & "    Set txU  = Me.Controls(""txtUsuario"")" & vbCrLf
    c = c & "    Set txC  = Me.Controls(""txtClave"")" & vbCrLf
    c = c & "    usuario = Trim(txU.Text): clave = txC.Text" & vbCrLf
    c = c & "    lblE.Visible = False" & vbCrLf
    c = c & "    If Len(usuario) = 0 Then" & vbCrLf
    c = c & "        lblE.Caption = ""Introduce tu nombre de usuario.""" & vbCrLf
    c = c & "        lblE.Visible = True: txU.SetFocus: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    If Len(clave) = 0 Then" & vbCrLf
    c = c & "        lblE.Caption = ""Introduce tu contrase"" & Chr(241) & ""a.""" & vbCrLf
    c = c & "        lblE.Visible = True: txC.SetFocus: Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    btnE.Enabled = False: btnE.Caption = ""Validando..."": DoEvents" & vbCrLf
    c = c & "    If modDatabase.ValidarCredenciales(usuario, clave) Then" & vbCrLf
    c = c & "        Me.Hide" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        lblE.Caption = ""Usuario o contrase"" & Chr(241) & ""a incorrectos.""" & vbCrLf
    c = c & "        lblE.Visible = True: txC.Text = """": txC.SetFocus" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    btnE.Enabled = True: btnE.Caption = ""Entrar""" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub btnSalir_Click()" & vbCrLf
    c = c & "    m_Cancelado = True: Me.Hide" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    c = c & "    If CloseMode = vbFormControlMenu Then" & vbCrLf
    c = c & "        m_Cancelado = True: Cancel = True: Me.Hide" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Public Property Get Cancelado() As Boolean" & vbCrLf
    c = c & "    Cancelado = m_Cancelado" & vbCrLf
    c = c & "End Property" & vbCrLf
    CodigoFrmLogin = c
End Function

Private Function CodigoFrmSelector() As String
    Dim c As String
    c = "Option Explicit" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    Dim residencias As Object, i As Long, nd As String, ctl As Object" & vbCrLf
    c = c & "    Me.Caption = ""BDAS - Selecci"" & Chr(243) & ""n de Residencia""" & vbCrLf
    c = c & "    Me.Width = 200: Me.Height = 160: Me.StartUpPosition = 1" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblBienvenida"", True)" & vbCrLf
    c = c & "    ctl.Left = 6: ctl.Top = 6: ctl.Width = 180: ctl.Height = 24" & vbCrLf
    c = c & "    ctl.TextAlign = 2" & vbCrLf
    c = c & "    With ctl.Font: .Size = 12: .Bold = True: End With" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblInstruccion"", True)" & vbCrLf
    c = c & "    ctl.Caption = ""Selecciona tu residencia:""" & vbCrLf
    c = c & "    ctl.Left = 6: ctl.Top = 33: ctl.Width = 180: ctl.Height = 18" & vbCrLf
    c = c & "    ctl.TextAlign = 2" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.ListBox.1"", ""lstResidencias"", True)" & vbCrLf
    c = c & "    ctl.Left = 18: ctl.Top = 57: ctl.Width = 156: ctl.Height = 54" & vbCrLf
    c = c & "    ctl.TabIndex = 0" & vbCrLf
    c = c & "    Set ctl = Me.Controls.Add(""Forms.CommandButton.1"", ""btnEntrar"", True)" & vbCrLf
    c = c & "    ctl.Caption = ""Entrar""" & vbCrLf
    c = c & "    ctl.Left = 57: ctl.Top = 123: ctl.Width = 84: ctl.Height = 24" & vbCrLf
    c = c & "    ctl.Default = True: ctl.TabIndex = 1" & vbCrLf
    c = c & "    Me.Controls(""lblBienvenida"").Caption = ""Bienvenido/a, "" & modDatabase.ObtenerNombreCompleto()" & vbCrLf
    c = c & "    Set residencias = modDatabase.ObtenerResidenciasAsignadas()" & vbCrLf
    c = c & "    Me.Controls(""lstResidencias"").Clear" & vbCrLf
    c = c & "    For i = 1 To residencias.Count" & vbCrLf
    c = c & "        Select Case residencias(i)" & vbCrLf
    c = c & "            Case ""GIJON"":  nd = ""Residencia de GIJ"" & ChrW(211) & ""N""" & vbCrLf
    c = c & "            Case ""SOTO"":   nd = ""Residencia de SOTO DEL BARCO""" & vbCrLf
    c = c & "            Case ""OVIEDO"": nd = ""Residencia de OVIEDO""" & vbCrLf
    c = c & "            Case Else:     nd = residencias(i)" & vbCrLf
    c = c & "        End Select" & vbCrLf
    c = c & "        Me.Controls(""lstResidencias"").AddItem nd" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "    If Me.Controls(""lstResidencias"").ListCount > 0 Then Me.Controls(""lstResidencias"").ListIndex = 0" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub btnEntrar_Click()" & vbCrLf
    c = c & "    Dim residencias As Object, selIndex As Long" & vbCrLf
    c = c & "    If Me.Controls(""lstResidencias"").ListIndex < 0 Then" & vbCrLf
    c = c & "        MsgBox ""Selecciona una residencia."", vbExclamation, ""BDAS""" & vbCrLf
    c = c & "        Exit Sub" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    Set residencias = modDatabase.ObtenerResidenciasAsignadas()" & vbCrLf
    c = c & "    selIndex = Me.Controls(""lstResidencias"").ListIndex + 1" & vbCrLf
    c = c & "    modDatabase.EstablecerResidenciaActiva residencias(selIndex)" & vbCrLf
    c = c & "    Me.Hide" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub lstResidencias_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    c = c & "    btnEntrar_Click" & vbCrLf
    c = c & "End Sub" & vbCrLf
    c = c & "" & vbCrLf
    c = c & "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf
    c = c & "    If CloseMode = vbFormControlMenu Then" & vbCrLf
    c = c & "        Cancel = True: Me.Hide" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    CodigoFrmSelector = c
End Function

Sub CrearFormularios()
    Dim vbp As Object
    Dim comp As Object

    Set vbp = ThisWorkbook.VBProject

    ' ---- Eliminar formularios anteriores (ignorar errores si no existen) ----
    On Error Resume Next
    vbp.VBComponents.Remove vbp.VBComponents.Item("frmLogin")
    vbp.VBComponents.Remove vbp.VBComponents.Item("frmSelectorResidencia")

    ' ---- Crear frmLogin ----
    ' NOTA: No usar comp.Properties() -- causa inestabilidad via COM externo.
    ' El Caption/Width/Height se establecen en UserForm_Initialize del propio form.
    Set comp = vbp.VBComponents.Add(3)   ' 3 = vbext_ct_MSForm (UserForm)
    comp.Name = "frmLogin"
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString CodigoFrmLogin()
    End With

    ' ---- Crear frmSelectorResidencia ----
    Set comp = vbp.VBComponents.Add(3)
    comp.Name = "frmSelectorResidencia"
    With comp.CodeModule
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .AddFromString CodigoFrmSelector()
    End With
    On Error GoTo 0
End Sub
