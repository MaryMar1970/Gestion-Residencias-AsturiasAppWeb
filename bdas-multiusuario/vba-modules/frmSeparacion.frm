VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSeparacion 
   Caption         =   "N� D�AS SEPARACI�N RESERVAS "
   ClientHeight    =   2535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3270
   OleObjectBlob   =   "frmSeparacion.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmSeparacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' UserForm: frmSeparacion
' =============================================================================
' Permite al usuario establecer el n�mero de d�as de separaci�n para una residencia
' concreta (GIJ�N, OVIEDO o SOTO). Al aceptar, valida el valor ingresado y lo almacena
' en la variable apropiada. Si el valor no es num�rico, muestra un mensaje de error.
' =============================================================================

' -----------------------------------------------------------------------------
' Variable p�blica que indica la residencia sobre la que se va a operar.
' Debe establecerse antes de mostrar el formulario.
' Valores esperados: "GIJ�N", "OVIEDO", "SOTO"
' -----------------------------------------------------------------------------
Public residenciaActual As String

' -----------------------------------------------------------------------------
' Evento: Click en el bot�n Aceptar
' Valida el dato ingresado y almacena el valor seg�n la residencia seleccionada.
' -----------------------------------------------------------------------------
Private Sub cmdAceptar_Click()
    Dim valor As Long
    On Error GoTo ErrorHandler ' Si ocurre error al convertir, salta a ErrorHandler
    valor = CLng(txtSeparacion.Value) ' Intenta convertir el valor ingresado a n�mero entero
    
    ' Seg�n la residencia seleccionada, almacena el valor en la variable correspondiente
    Select Case residenciaActual
        Case "GIJ�N":   SeparacionGijon = valor
        Case "OVIEDO":  SeparacionOviedo = valor
        Case "SOTO":    SeparacionSoto = valor
    End Select

    Unload Me ' Cierra el formulario tras guardar el valor
    Exit Sub

ErrorHandler:
    MsgBox "Introduce un n�mero v�lido de d�as.", vbExclamation
End Sub

' -----------------------------------------------------------------------------
' Evento: Cambio en el cuadro de texto (no implementado, reservado para extender)
' -----------------------------------------------------------------------------
Private Sub txtSeparacion_Change()
    ' (Puedes agregar validaciones en tiempo real aqu� si lo necesitas)
End Sub

' -----------------------------------------------------------------------------
' Evento: Inicializaci�n del formulario
' Establece el valor predeterminado del cuadro de texto en 0 d�as.
' -----------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    txtSeparacion.Value = 0
End Sub

