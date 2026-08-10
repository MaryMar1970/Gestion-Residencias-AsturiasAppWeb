Attribute VB_Name = "SeparacionDiasReservas"
'=================================================================================
' M�dulo: SeparacionDiasReserva
' Descripci�n:
'   - Permite definir los d�as de separaci�n entre reservas para cada residencia.
'   - Los valores se almacenan en la hoja oculta "Auxiliar" (celdas C3, D3 y E3).
'   - Incluye validaci�n, manejo de errores y persistencia entre sesiones.
'=================================================================================

Public SeparacionGijon As Long
Public SeparacionOviedo As Long
Public SeparacionSoto As Long

'---------------------------------------------------------------------------------
' CargarSeparaciones
'   Carga los valores almacenados en "Auxiliar" al abrir el archivo.
'   Se debe llamar desde Workbook_Open
'---------------------------------------------------------------------------------
Public Sub CargarSeparaciones()
    On Error Resume Next
    Dim wsAux As Worksheet
    Set wsAux = ThisWorkbook.Sheets("Auxiliar")
    
    ' Leer valores desde hoja "Auxiliar", si son v�lidos
    If IsNumeric(wsAux.Range("C3").Value) Then SeparacionGijon = CLng(wsAux.Range("C3").Value)
    If IsNumeric(wsAux.Range("D3").Value) Then SeparacionOviedo = CLng(wsAux.Range("D3").Value)
    If IsNumeric(wsAux.Range("E3").Value) Then SeparacionSoto = CLng(wsAux.Range("E3").Value)
End Sub

'---------------------------------------------------------------------------------
' CambiarSeparacionGijon
'   Pide al usuario un nuevo valor para la separaci�n de Gij�n.
'   Valida el dato, lo guarda en memoria y en la hoja "Auxiliar"
'---------------------------------------------------------------------------------
Sub CambiarSeparacionGijon(control As IRibbonControl)
    Dim respuesta As Variant
    respuesta = InputBox("Introduce d�as de separaci�n entre entrada/salida:", "Separaci�n reservas Gij�n", SeparacionGijon)
    
    If respuesta = "" Then Exit Sub ' Cancelado

    If IsNumeric(respuesta) And val(respuesta) >= 0 Then
        SeparacionGijon = CLng(respuesta)
        ThisWorkbook.Sheets("Auxiliar").Range("C3").Value = SeparacionGijon
        MsgBox "Separaci�n para Gij�n actualizada a " & SeparacionGijon & " d�as.", vbInformation
    Else
        MsgBox "Entrada no v�lida. Debes introducir un n�mero entero mayor o igual que cero.", vbExclamation
    End If
End Sub

'---------------------------------------------------------------------------------
' CambiarSeparacionOviedo
'---------------------------------------------------------------------------------
Sub CambiarSeparacionOviedo(control As IRibbonControl)
    Dim respuesta As Variant
    respuesta = InputBox("Introduce d�as de separaci�n entre entrada/salida:", "Separaci�n reservas Oviedo", SeparacionOviedo)

    If respuesta = "" Then Exit Sub

    If IsNumeric(respuesta) And val(respuesta) >= 0 Then
        SeparacionOviedo = CLng(respuesta)
        ThisWorkbook.Sheets("Auxiliar").Range("D3").Value = SeparacionOviedo
        MsgBox "Separaci�n para Oviedo actualizada a " & SeparacionOviedo & " d�as.", vbInformation
    Else
        MsgBox "Entrada no v�lida. Debes introducir un n�mero entero mayor o igual que cero.", vbExclamation
    End If
End Sub

'---------------------------------------------------------------------------------
' CambiarSeparacionSoto
'---------------------------------------------------------------------------------
Sub CambiarSeparacionSoto(control As IRibbonControl)
    Dim respuesta As Variant
    respuesta = InputBox("Introduce d�as de separaci�n entre entrada/salida:", "Separaci�n reservas Soto", SeparacionSoto)

    If respuesta = "" Then Exit Sub

    If IsNumeric(respuesta) And val(respuesta) >= 0 Then
        SeparacionSoto = CLng(respuesta)
        ThisWorkbook.Sheets("Auxiliar").Range("E3").Value = SeparacionSoto
        MsgBox "Separaci�n para Soto actualizada a " & SeparacionSoto & " d�as.", vbInformation
    Else
        MsgBox "Entrada no v�lida. Debes introducir un n�mero entero mayor o igual que cero.", vbExclamation
    End If
End Sub

