VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAviso 
   Caption         =   "Valor copiado al portapapeles:"
   ClientHeight    =   975
   ClientLeft      =   240
   ClientTop       =   930
   ClientWidth     =   6150
   OleObjectBlob   =   "frmAviso.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' UserForm: frmAviso
' =============================================================================
' Este formulario muestra un aviso flotante temporal en pantalla, con cierre
' autom�tico despu�s de un tiempo configurable (timer), o cierre inmediato si
' el usuario hace clic en el formulario o en el label del aviso.
'
' El temporizador se programa usando Application.OnTime, que s�lo ejecuta
' procedimientos p�blicos en m�dulos est�ndar: por eso, la macro que cierra
' este formulario (CerrarAvisoFlotante) debe estar en un m�dulo est�ndar.
' =============================================================================

Option Explicit

' -----------------------------------------------------------------------------
' Variable privada para almacenar el instante programado de cierre autom�tico.
' -----------------------------------------------------------------------------
Private avisarTimer As Double

' -----------------------------------------------------------------------------
' M�todo p�blico para mostrar el aviso.
' - texto: Texto que se mostrar� en el label del formulario.
' - segundos (opcional): Tiempo en segundos hasta el cierre autom�tico.
' -----------------------------------------------------------------------------
Public Sub MostrarAviso(ByVal texto As String, Optional ByVal segundos As Double = 1)
    ' Establece el texto del aviso.
    Me.lblAviso.Caption = texto
    
    ' Fuerza el repintado del formulario para actualizar el texto.
    Me.Repaint
    
    ' Muestra el formulario de manera no modal (permite seguir usando Excel).
    Me.Show vbModeless
    
    ' ===========================
    ' SELECCI�N Y PROGRAMACI�N DEL TIMER
    ' ===========================
    ' Calcula el instante futuro en el que se cerrar� el formulario.
    ' (Now + segundos/86400) convierte los segundos en la fracci�n de d�a que usa VBA.
    avisarTimer = Now + (segundos / 86400)
    
    ' Programa la ejecuci�n del procedimiento de cierre autom�tico.
    ' "CerrarAvisoFlotante" debe estar en un m�dulo est�ndar.
    Application.OnTime avisarTimer, "CerrarAvisoFlotante"
    ' ===========================
    ' FIN SELECCI�N TIMER
    ' ===========================
End Sub

' -----------------------------------------------------------------------------
' Cierra el formulario (llamado tanto por el temporizador como por clic de usuario).
' -----------------------------------------------------------------------------
Public Sub CerrarAviso()
    On Error Resume Next
    Unload Me
End Sub

' -----------------------------------------------------------------------------
' Cierre inmediato si el usuario hace clic en cualquier parte del formulario.
' -----------------------------------------------------------------------------
Private Sub UserForm_Click()
    CerrarAviso
End Sub

' -----------------------------------------------------------------------------
' Cierre inmediato si el usuario hace clic en el label del aviso.
' -----------------------------------------------------------------------------
Private Sub lblAviso_Click()
    CerrarAviso
End Sub

' =============================================================================
' Notas de uso:
' - El temporizador se selecciona y programa en MostrarAviso, l�nea avisarTimer = ...
' - Es fundamental que el procedimiento CerrarAvisoFlotante est� en un m�dulo est�ndar.
' - La propiedad WordWrap del label lblAviso debe estar en True para saltos de l�nea.
' =============================================================================
