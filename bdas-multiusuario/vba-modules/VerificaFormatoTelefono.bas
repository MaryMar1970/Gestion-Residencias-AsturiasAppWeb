Attribute VB_Name = "VerificaFormatoTelefono"
'=================================================================================
' M�dulo: ValidacionTelefono
'
' Descripci�n:
'   Valida el formato de un n�mero de tel�fono en una celda espec�fica de una hoja.
'   - Solo act�a en la columna de tel�fonos, seg�n la hoja.
'   - El tel�fono debe tener 9 d�gitos, comenzar por 6, 7, 8 o 9, y ser num�rico.
'   - Si el formato es incorrecto, borra la celda, muestra un mensaje y devuelve el foco.
'=================================================================================

Option Explicit

'---------------------------------------------------------------------------------
' VerificaFormatoTelefonoCelda
'   - ws: hoja de trabajo en la que se valida el tel�fono
'   - targetCell: celda que contiene el tel�fono a validar
'
'   Seg�n la hoja, verifica la columna correcta (V=22 para GIJ�N/SOTO, W=23 para OVIEDO).
'   Si el formato no es v�lido, borra el contenido, muestra un aviso y devuelve el foco.
'---------------------------------------------------------------------------------
Public Sub VerificaFormatoTelefonoCelda(ws As Worksheet, ByVal targetCell As Range)
    Dim telefono As String           ' Tel�fono a validar (como texto)
    Dim valido As Boolean            ' Resultado de la validaci�n de formato
    Dim columna As Long              ' N�mero de columna de la celda objetivo
    Dim nombreHoja As String         ' Nombre de la hoja activa

    nombreHoja = ws.Name
    columna = targetCell.Column

    '---------------------------------------------------------------------
    ' 1. Determinar si la celda est� en la columna correcta seg�n la hoja
    '---------------------------------------------------------------------
    Select Case nombreHoja
        Case "RESIDENCIA GIJ�N", "RESIDENCIA SOTO"
            If columna <> 22 Then Exit Sub ' Solo columna V (22)
        Case "RESIDENCIA OVIEDO"
            If columna <> 23 Then Exit Sub ' Solo columna W (23)
        Case Else
            Exit Sub                      ' No valida en otras hojas
    End Select

    '---------------------------------------------------------------------
    ' 2. No hacer nada si la celda est� vac�a (permite celdas en blanco)
    '---------------------------------------------------------------------
    telefono = Trim(targetCell.Value)
    If telefono = "" Then Exit Sub

    '---------------------------------------------------------------------
    ' 3. Validar formato:
    '    - 9 d�gitos exactos
    '    - Primer d�gito 6, 7, 8 o 9
    '    - Todos los caracteres son num�ricos
    '---------------------------------------------------------------------
    valido = telefono Like "[6789]########" And Len(telefono) = 9 And IsNumeric(telefono)

    '---------------------------------------------------------------------
    ' 4. Si el formato es incorrecto:
    '    - Borra la celda
    '    - Muestra mensaje informativo
    '    - Devuelve el foco a la celda
    '    - Desactiva temporalmente eventos para evitar bucles
    '---------------------------------------------------------------------
        If Not valido Then
        On Error GoTo SafeExit
        Application.enableEvents = False
        MsgBox "El n�mero de tel�fono debe contener 9 d�gitos y comenzar por 6, 7, 8 o 9.", vbExclamation, "Formato de Tel�fono Incorrecto"
        targetCell.ClearContents
        targetCell.Select
    End If
SafeExit:
    Application.enableEvents = True
End Sub
