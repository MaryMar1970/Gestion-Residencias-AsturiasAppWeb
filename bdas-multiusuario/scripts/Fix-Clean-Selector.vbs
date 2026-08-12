Sub SetCleanSelectorCode(p)
    Dim xl, wb, proj, comp, codeMod, cleanCode
    
    WScript.Echo "Inyectando codigo limpio en: " & p
    
    cleanCode = "Option Explicit" & vbCrLf & _
                "" & vbCrLf & _
                "Private WithEvents btnEntrar As MSForms.CommandButton" & vbCrLf & _
                "Private WithEvents lstResidencias As MSForms.ListBox" & vbCrLf & _
                "" & vbCrLf & _
                "Private Sub UserForm_Initialize()" & vbCrLf & _
                "    Dim residencias As Object  ' Collection" & vbCrLf & _
                "    Dim i As Long" & vbCrLf & _
                "    Dim nombreDisplay As String" & vbCrLf & _
                "    Dim lblB As Object" & vbCrLf & _
                "    Dim ctl As Object" & vbCrLf & _
                "" & vbCrLf & _
                "    Me.Caption = ""BDAS - Selecci"" & Chr(243) & ""n de Residencia""" & vbCrLf & _
                "    Me.Width = 240" & vbCrLf & _
                "    Me.Height = 200" & vbCrLf & _
                "    Me.StartUpPosition = 1" & vbCrLf & _
                "" & vbCrLf & _
                "    ' --- Label Bienvenida ---" & vbCrLf & _
                "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblBienvenida"", True)" & vbCrLf & _
                "    ctl.Left = 10: ctl.Top = 10: ctl.Width = 204: ctl.Height = 24" & vbCrLf & _
                "    ctl.Caption = ""Bienvenido/a""" & vbCrLf & _
                "    ctl.TextAlign = 2" & vbCrLf & _
                "    With ctl.Font: .Size = 12: .Bold = True: End With" & vbCrLf & _
                "" & vbCrLf & _
                "    ' --- Label Instruccion ---" & vbCrLf & _
                "    Set ctl = Me.Controls.Add(""Forms.Label.1"", ""lblInstruccion"", True)" & vbCrLf & _
                "    ctl.Left = 10: ctl.Top = 34: ctl.Width = 204: ctl.Height = 20" & vbCrLf & _
                "    ctl.Caption = ""Selecciona tu residencia:""" & vbCrLf & _
                "    ctl.TextAlign = 2" & vbCrLf & _
                "" & vbCrLf & _
                "    ' --- ListBox Residencias ---" & vbCrLf & _
                "    Set lstResidencias = Me.Controls.Add(""Forms.ListBox.1"", ""lstResidencias"", True)" & vbCrLf & _
                "    lstResidencias.Left = 24: lstResidencias.Top = 58: lstResidencias.Width = 176: lstResidencias.Height = 65" & vbCrLf & _
                "    lstResidencias.TabIndex = 0" & vbCrLf & _
                "" & vbCrLf & _
                "    ' --- Boton Entrar ---" & vbCrLf & _
                "    Set btnEntrar = Me.Controls.Add(""Forms.CommandButton.1"", ""btnEntrar"", True)" & vbCrLf & _
                "    btnEntrar.Caption = ""Entrar""" & vbCrLf & _
                "    btnEntrar.Left = 72: btnEntrar.Top = 134: btnEntrar.Width = 80: btnEntrar.Height = 26" & vbCrLf & _
                "    btnEntrar.Default = True" & vbCrLf & _
                "    btnEntrar.TabIndex = 1" & vbCrLf & _
                "" & vbCrLf & _
                "    ' --- Rellenar datos ---" & vbCrLf & _
                "    Set lblB = Me.Controls(""lblBienvenida"")" & vbCrLf & _
                "    lblB.Caption = ""Bienvenido/a, "" & modDatabase.ObtenerNombreCompleto()" & vbCrLf & _
                "" & vbCrLf & _
                "    lstResidencias.Clear" & vbCrLf & _
                "    Set residencias = modDatabase.ObtenerResidenciasAsignadas()" & vbCrLf & _
                "" & vbCrLf & _
                "    For i = 1 To residencias.Count" & vbCrLf & _
                "        Select Case residencias(i)" & vbCrLf & _
                "            Case ""GIJON"":  nombreDisplay = ""Residencia de GIJ"" & Chr(211) & ""N""" & vbCrLf & _
                "            Case ""SOTO"":   nombreDisplay = ""Residencia de SOTO DEL BARCO""" & vbCrLf & _
                "            Case ""OVIEDO"": nombreDisplay = ""Residencia de OVIEDO""" & vbCrLf & _
                "            Case Else:     nombreDisplay = residencias(i)" & vbCrLf & _
                "        End Select" & vbCrLf & _
                "        lstResidencias.AddItem nombreDisplay" & vbCrLf & _
                "    Next i" & vbCrLf & _
                "" & vbCrLf & _
                "    If lstResidencias.ListCount > 0 Then lstResidencias.ListIndex = 0" & vbCrLf & _
                "End Sub" & vbCrLf & _
                "" & vbCrLf & _
                "Private Sub btnEntrar_Click()" & vbCrLf & _
                "    ProcesarSeleccion" & vbCrLf & _
                "End Sub" & vbCrLf & _
                "" & vbCrLf & _
                "Private Sub lstResidencias_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf & _
                "    ProcesarSeleccion" & vbCrLf & _
                "End Sub" & vbCrLf & _
                "" & vbCrLf & _
                "Private Sub ProcesarSeleccion()" & vbCrLf & _
                "    Dim residencias As Object" & vbCrLf & _
                "    Dim selIndex As Long" & vbCrLf & _
                "" & vbCrLf & _
                "    If lstResidencias.ListIndex < 0 Then" & vbCrLf & _
                "        MsgBox ""Selecciona una residencia antes de continuar."", vbExclamation, ""BDAS""" & vbCrLf & _
                "        Exit Sub" & vbCrLf & _
                "    End If" & vbCrLf & _
                "" & vbCrLf & _
                "    Set residencias = modDatabase.ObtenerResidenciasAsignadas()" & vbCrLf & _
                "    selIndex = lstResidencias.ListIndex + 1  ' Collection base 1" & vbCrLf & _
                "" & vbCrLf & _
                "    modDatabase.EstablecerResidenciaActiva residencias(selIndex)" & vbCrLf & _
                "    Me.Hide" & vbCrLf & _
                "End Sub" & vbCrLf & _
                "" & vbCrLf & _
                "Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)" & vbCrLf & _
                "    If CloseMode = vbFormControlMenu Then" & vbCrLf & _
                "        Cancel = True" & vbCrLf & _
                "        Me.Hide" & vbCrLf & _
                "    End If" & vbCrLf & _
                "End Sub"
                
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    
    Set wb = xl.Workbooks.Open(p, 0, False)
    Set proj = wb.VBProject
    Set comp = proj.VBComponents.Item("frmSelectorResidencia")
    Set codeMod = comp.CodeModule
    
    If codeMod.CountOfLines > 0 Then
        codeMod.DeleteLines 1, codeMod.CountOfLines
    End If
    codeMod.InsertLines 1, cleanCode
    
    wb.Save
    wb.Close False
    xl.Quit
End Sub

Dim fso, bdasDir, folder, file, masterPath, deployPath
Set fso = CreateObject("Scripting.FileSystemObject")
bdasDir = "H:\ResidenciaApp\bdas-multiusuario\"

Set folder = fso.GetFolder(bdasDir)
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "xlsm" Then
        masterPath = file.Path
        Exit For
    End If
Next

deployPath = "H:\ResidenciaBD\BDAS_Multiusuario.xlsm"

SetCleanSelectorCode masterPath
SetCleanSelectorCode deployPath

WScript.Echo "CODIGO LIMPIO INYECTADO CON EXITO."
