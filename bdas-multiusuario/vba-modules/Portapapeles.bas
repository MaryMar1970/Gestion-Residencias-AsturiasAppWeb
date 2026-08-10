Attribute VB_Name = "Portapapeles"
Option Explicit
' ===============================================================================
' M�dulo: modPortapapeles
' Funci�n: Copiar texto al portapapeles usando API de Windows (Unicode, robusto)
' Uso recomendado: CopiarYAvisarSiNoVacio "Texto a copiar", frmAviso
' ===============================================================================

#If VBA7 Then
    '--- Declaraciones para sistemas de 64 bits (VBA7) ---
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal lpDest As LongPtr, ByVal lpSource As LongPtr, ByVal cbCopy As LongPtr)
    Private Const GMEM_MOVEABLE As Long = &H2           ' Memoria movible (necesario para portapapeles)
    Private Const CF_UNICODETEXT As Long = 13           ' Formato de texto UNICODE para portapapeles
#Else
    '--- Declaraciones para sistemas de 32 bits (legacy VBA) ---
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByVal lpDest As Long, ByVal lpSource As Long, ByVal cbCopy As Long)
    Private Const GMEM_MOVEABLE As Long = &H2
    Private Const CF_UNICODETEXT As Long = 13
#End If

' ===============================================================================
' Funci�n principal: CopiarTextoAlPortapapelesAPI
' Descripci�n:
'   Copia el texto recibido al portapapeles de Windows en formato UNICODE.
'   Usa API de Windows para m�xima compatibilidad y robustez.
' Par�metro:
'   - text: texto a copiar (String)
' Devuelve:
'   - True si tuvo �xito, False si hubo error
' Uso:
'   If CopiarTextoAlPortapapelesAPI("texto") Then MsgBox "OK" Else MsgBox "Error"
' ===============================================================================
Public Function CopiarTextoAlPortapapelesAPI(ByVal text As String) As Boolean
    Dim hGlobal As LongPtr, lpGlobal As LongPtr
    Dim bytes() As Byte, nBytes As LongPtr

    On Error GoTo ErrCopy

    ' --- 1. Convertir el texto a UNICODE (UTF-16), con terminador NULL ---
    bytes = text & vbNullChar
    nBytes = (UBound(bytes) + 1)

    ' --- 2. Reservar un bloque de memoria global movible ---
    hGlobal = GlobalAlloc(GMEM_MOVEABLE, nBytes)
    If hGlobal = 0 Then GoTo ErrCopy

    ' --- 3. Bloquear la memoria y obtener el puntero ---
    lpGlobal = GlobalLock(hGlobal)
    If lpGlobal = 0 Then
        GlobalFree hGlobal
        GoTo ErrCopy
    End If

    ' --- 4. Copiar los bytes del texto en la memoria reservada ---
    CopyMemory lpGlobal, VarPtr(bytes(0)), nBytes
    Call GlobalUnlock(hGlobal)

    ' --- 5. Abrir el portapapeles del sistema ---
    If OpenClipboard(0&) = 0 Then
        GlobalFree hGlobal
        GoTo ErrCopy
    End If

    ' --- 6. Vaciar el portapapeles y establecer el nuevo contenido ---
    Call EmptyClipboard
    If SetClipboardData(CF_UNICODETEXT, hGlobal) = 0 Then
        CloseClipboard
        GlobalFree hGlobal
        GoTo ErrCopy
    End If

    ' --- 7. Cerrar el portapapeles ---
    Call CloseClipboard

    ' --- 8. �xito ---
    CopiarTextoAlPortapapelesAPI = True
    Exit Function

ErrCopy:
    ' --- 9. En caso de error, devolver False (no copiado) ---
    CopiarTextoAlPortapapelesAPI = False
End Function

' ===============================================================================
' Wrapper: CopiarYAvisarSiNoVacio
' Descripci�n:
'   - Limpia el texto de saltos de l�nea y espacios.
'   - Solo copia si, tras limpiar, el texto no est� vac�o.
'   - Si se pasa un formulario de aviso (ej: frmAviso), muestra el texto copiado.
' Par�metros:
'   - texto: texto a copiar (String)
'   - frmAviso (opcional): formulario con m�todo .MostrarAviso para mostrar el aviso flotante
' Uso:
'   CopiarYAvisarSiNoVacio texto, frmAviso
' ===============================================================================
Public Sub CopiarYAvisarSiNoVacio(ByVal texto As String, Optional ByRef frmAviso As Object = Nothing)
    Dim textoLimpio As String

    ' --- 1. Eliminar saltos de l�nea y espacios al principio/final ---
    textoLimpio = Replace(texto, Chr(10), "")
    textoLimpio = Replace(textoLimpio, Chr(13), "")
    textoLimpio = Trim(textoLimpio)

    ' --- 2. Solo contin�a si el texto NO est� vac�o ---
    If textoLimpio <> "" Then
        On Error Resume Next
        ' --- 3. Copiar al portapapeles usando la funci�n principal ---
        CopiarTextoAlPortapapelesAPI textoLimpio
        ' --- 4. Si se pas� un formulario de aviso, mostrar el texto ---
        If Not frmAviso Is Nothing Then
            frmAviso.MostrarAviso textoLimpio
        End If
        On Error GoTo 0
    End If
End Sub

