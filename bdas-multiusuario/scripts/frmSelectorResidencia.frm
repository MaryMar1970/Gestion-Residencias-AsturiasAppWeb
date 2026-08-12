Option Explicit

Private WithEvents btnEntrar As MSForms.CommandButton
Private WithEvents lstResidencias As MSForms.ListBox

Private Sub UserForm_Initialize()
    Dim residencias As Object  ' Collection
    Dim i As Long
    Dim nombreDisplay As String
    Dim lblB As Object
    Dim ctl As Object

    Me.Caption = "BDAS - Selecci" & Chr(243) & "n de Residencia"
    Me.Width = 240
    Me.Height = 200
    Me.StartUpPosition = 1

    ' --- Label Bienvenida ---
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblBienvenida", True)
    ctl.Left = 10: ctl.Top = 10: ctl.Width = 204: ctl.Height = 24
    ctl.Caption = "Bienvenido/a"
    ctl.TextAlign = 2
    With ctl.Font: .Size = 12: .Bold = True: End With

    ' --- Label Instruccion ---
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblInstruccion", True)
    ctl.Left = 10: ctl.Top = 34: ctl.Width = 204: ctl.Height = 20
    ctl.Caption = "Selecciona tu residencia:"
    ctl.TextAlign = 2

    ' --- ListBox Residencias ---
    Set lstResidencias = Me.Controls.Add("Forms.ListBox.1", "lstResidencias", True)
    lstResidencias.Left = 24: lstResidencias.Top = 58: lstResidencias.Width = 176: lstResidencias.Height = 65
    lstResidencias.TabIndex = 0

    ' --- Boton Entrar ---
    Set btnEntrar = Me.Controls.Add("Forms.CommandButton.1", "btnEntrar", True)
    btnEntrar.Caption = "Entrar"
    btnEntrar.Left = 72: btnEntrar.Top = 134: btnEntrar.Width = 80: btnEntrar.Height = 26
    btnEntrar.Default = True
    btnEntrar.TabIndex = 1

    ' --- Rellenar datos ---
    Set lblB = Me.Controls("lblBienvenida")
    lblB.Caption = "Bienvenido/a, " & modDatabase.ObtenerNombreCompleto()

    lstResidencias.Clear
    Set residencias = modDatabase.ObtenerResidenciasAsignadas()

    For i = 1 To residencias.Count
        Select Case residencias(i)
            Case "GIJON":  nombreDisplay = "Residencia de GIJ" & Chr(211) & "N"
            Case "SOTO":   nombreDisplay = "Residencia de SOTO DEL BARCO"
            Case "OVIEDO": nombreDisplay = "Residencia de OVIEDO"
            Case Else:     nombreDisplay = residencias(i)
        End Select
        lstResidencias.AddItem nombreDisplay
    Next i

    If lstResidencias.ListCount > 0 Then lstResidencias.ListIndex = 0
End Sub

Private Sub btnEntrar_Click()
    ProcesarSeleccion
End Sub

Private Sub lstResidencias_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ProcesarSeleccion
End Sub

Private Sub ProcesarSeleccion()
    Dim residencias As Object
    Dim selIndex As Long

    If lstResidencias.ListIndex < 0 Then
        MsgBox "Selecciona una residencia antes de continuar.", vbExclamation, "BDAS"
        Exit Sub
    End If

    Set residencias = modDatabase.ObtenerResidenciasAsignadas()
    selIndex = lstResidencias.ListIndex + 1  ' Collection base 1

    Me.Hide
    modDatabase.EstablecerResidenciaActiva residencias(selIndex)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
