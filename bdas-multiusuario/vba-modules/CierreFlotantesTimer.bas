Attribute VB_Name = "CierreFlotantesTimer"
' =============================================================================
' M�dulo est�ndar: modCierreFlotantes
' =============================================================================
' Proporciona procedimientos p�blicos para el cierre autom�tico de UserForms
' mediante Application.OnTime, ya que OnTime s�lo puede invocar macros ubicadas
' en m�dulos est�ndar, no en formularios.
' Incluye cierres para frmAviso y frmReservasUnidad.
' =============================================================================

Option Explicit

' Cierra el UserForm frmAviso (llamado por Application.OnTime)
Public Sub CerrarAvisoFlotante()
    On Error Resume Next
    frmAviso.CerrarAviso
End Sub

' Cierra el UserForm frmReservasUnidad (llamado por Application.OnTime)
Public Sub CerrarReservasUnidadFlotante()
    On Error Resume Next
    frmReservasUnidad.CerrarFormulario
End Sub
