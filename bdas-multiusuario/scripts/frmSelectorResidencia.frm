Option Explicit

Private Sub UserForm_Initialize()
    Dim residencias As Object  ' Collection
    Dim i As Long
    Dim nombreDisplay As String
    Dim lstR As Object
    Dim lblB As Object
    Dim ctl As Object

    Me.Caption = "BDAS" & Chr(160) & Chr(8212) & Chr(160) & "Selecci" & Chr(243) & "n de Residencia"
    Me.Width = 200
    Me.Height = 160
    Me.StartUpPosition = 1

    ' --- Label Bienvenida ---
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblBienvenida", True)
    ctl.Left = 6: ctl.Top = 6: ctl.Width = 180: ctl.Height = 24
    ctl.Caption = "Bienvenido/a"
    ctl.TextAlign = 2
    With ctl.Font: .Size = 12: .Bold = True: End With

    ' --- Label Instruccion ---
    Set ctl = Me.Controls.Add("Forms.Label.1", "lblInstruccion", True)
    ctl.Left = 6: ctl.Top = 33: ctl.Width = 180: ctl.Height = 24
    ctl.Caption = "Selecciona tu residencia:"
    ctl.TextAlign = 2

    ' --- ListBox Residencias ---
    Set ctl = Me.Controls.Add("Forms.ListBox.1", "lstResidencias", True)
    ctl.Left = 18: ctl.Top = 60: ctl.Width = 156: ctl.Height = 54
    ctl.TabIndex = 0

    ' --- Boton Entrar ---
    Set ctl = Me.Controls.Add("Forms.CommandButton.1", "btnEntrar", True)
    ctl.Caption = "Entrar"
    ctl.Left = 57: ctl.Top = 123: ctl.Width = 84: ctl.Height = 24
    ctl.Default = True
    ctl.TabIndex = 1

    ' --- Rellenar datos ---
    Set lblB = Me.Controls("lblBienvenida")
    lblB.Caption = "Bienvenido/a, " & modDatabase.ObtenerNombreCompleto()

    Set lstR = Me.Controls("lstResidencias")
    lstR.Clear
    Set residencias = modDatabase.ObtenerResidenciasAsignadas()

    For i = 1 To residencias.Count
        Select Case residencias(i)
            Case "GIJON":  nombreDisplay = "Residencia de GIJ" & Chr(211) & "N"
            Case "SOTO":   nombreDisplay = "Residencia de SOTO DEL BARCO"
            Case "OVIEDO": nombreDisplay = "Residencia de OVIEDO"
            Case Else:     nombreDisplay = residencias(i)
        End Select
        lstR.AddItem nombreDisplay
    Next i

    If lstR.ListCount > 0 Then lstR.ListIndex = 0
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
    Dim lstR As Object

    Set lstR = Me.Controls("lstResidencias")

    If lstR.ListIndex < 0 Then
        MsgBox "Selecciona una residencia antes de continuar.", vbExclamation, "BDAS"
        Exit Sub
    End If

    Set residencias = modDatabase.ObtenerResidenciasAsignadas()
    selIndex = lstR.ListIndex + 1  ' Collection base 1

    modDatabase.EstablecerResidenciaActiva residencias(selIndex)
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
