VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectorResidencia
   Caption         =   "BDAS — Selección de Residencia"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "frmSelectorResidencia.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectorResidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' UserForm: frmSelectorResidencia
' Propósito: Permite al usuario seleccionar su residencia activa cuando tiene
'            acceso a más de una residencia.
'
' INSTRUCCIONES PARA CREAR ESTE FORMULARIO EN VBA:
' 1. Abrir el Editor VBA (ALT+F11)
' 2. Insertar > UserForm
' 3. Renombrar a "frmSelectorResidencia" (en la ventana de Propiedades)
' 4. Añadir los siguientes controles:
'
'    Control          | Nombre          | Propiedades clave
'    -----------------|-----------------|------------------------------------------
'    Label            | lblBienvenida   | Caption="Bienvenido, [nombre]", Font=12pt Bold
'    Label            | lblInstruccion  | Caption="Selecciona la residencia con la que vas a trabajar:"
'    ListBox          | lstResidencias  | ListStyle=1 (fmListStyleOption), TabIndex=0
'    CommandButton    | btnEntrar       | Caption="Entrar", Default=True, TabIndex=1
'
' 5. Pegar este código en el módulo del formulario
' ==============================================================================

Option Explicit

Private Sub UserForm_Initialize()
    Dim residencias As Collection
    Dim i As Long
    Dim nombreDisplay As String
    
    Me.Caption = "BDAS — Selección de Residencia"
    Me.StartUpPosition = 1  ' CenterOwner
    
    ' Mostrar nombre del usuario
    lblBienvenida.Caption = "Bienvenido/a, " & modDatabase.ObtenerNombreCompleto()
    
    ' Cargar residencias asignadas en la lista
    lstResidencias.Clear
    Set residencias = modDatabase.ObtenerResidenciasAsignadas()
    
    For i = 1 To residencias.Count
        Select Case residencias(i)
            Case "GIJON": nombreDisplay = "Residencia de GIJÓN"
            Case "SOTO": nombreDisplay = "Residencia de SOTO DEL BARCO"
            Case "OVIEDO": nombreDisplay = "Residencia de OVIEDO"
            Case Else: nombreDisplay = residencias(i)
        End Select
        lstResidencias.AddItem nombreDisplay
    Next i
    
    ' Seleccionar la primera por defecto
    If lstResidencias.ListCount > 0 Then
        lstResidencias.ListIndex = 0
    End If
End Sub

Private Sub btnEntrar_Click()
    Dim residencias As Collection
    Dim selIndex As Long
    
    ' Verificar que hay una selección
    If lstResidencias.ListIndex < 0 Then
        MsgBox "Selecciona una residencia antes de continuar.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener la residencia correspondiente al índice seleccionado
    Set residencias = modDatabase.ObtenerResidenciasAsignadas()
    selIndex = lstResidencias.ListIndex + 1  ' Collection es base 1
    
    ' Establecer la residencia activa en modDatabase
    modDatabase.EstablecerResidenciaActiva residencias(selIndex)
    
    ' Cerrar el formulario
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Si cierran con la X, no seleccionar ninguna residencia
    ' (LoginUsuario() detectará que m_ResidenciaActiva está vacía)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub lstResidencias_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Doble clic = seleccionar y entrar
    btnEntrar_Click
End Sub
