VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBusqueda 
   Caption         =   "Selecci�n b�squeda"
   ClientHeight    =   3075
   ClientLeft      =   240
   ClientTop       =   900
   ClientWidth     =   4005
   OleObjectBlob   =   "frmBusqueda.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Variable p�blica para almacenar la opci�n seleccionada
Public searchOption As String

Private Sub cmbOpcion_Change()

End Sub

Private Sub UserForm_Initialize()
    ' Agregar las opciones al ComboBox
    With cmbOpciones
        .Clear
        .AddItem "NOMBRE"
        .AddItem "N� DNI"
        .AddItem "N� ORDEN"
        .AddItem "N� FACTURA"
        .ListIndex = 0                           ' Selecci�n por defecto
    End With
End Sub

' Bot�n Aceptar
Private Sub cmdAceptar_Click()
    searchOption = cmbOpciones.Value
    Me.Hide                                      ' Oculta el formulario en lugar de descargarlo
End Sub

' Bot�n Cancelar
Private Sub cmdCancelar_Click()
    searchOption = ""
    Me.Hide                                      ' Oculta el formulario en lugar de descargarlo
End Sub

