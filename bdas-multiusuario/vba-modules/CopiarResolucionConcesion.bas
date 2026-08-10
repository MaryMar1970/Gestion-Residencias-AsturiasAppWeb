Attribute VB_Name = "CopiarResolucionConcesion"
Option Explicit

'========================================================================================
' Declaraciones API para manipulaci�n avanzada del portapapeles y memoria.
' Compatibilidad con VBA7 (64 bits) y versiones anteriores (32 bits).
'========================================================================================
#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

' Constantes para la gesti�n de memoria y formatos de portapapeles
Const GHND As Long = &H42
Const CF_TEXT As Long = 1
Const CF_UNICODETEXT As Long = 13   ' Formato Unicode para compatibilidad con caracteres especiales (�)

'========================================================================================
' Subrutina principal: CopiarResolucion
' Copia al portapapeles el contenido visible del rango A1:F40 de la hoja activa,
' omitiendo una celda espec�fica seg�n la hoja, y limpia el valor de C1 antes de copiar.
'========================================================================================
Sub CopiarResolucion()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Limpia el contenido de la celda C1 cada vez que se ejecuta la macro.
    ws.Range("C1").ClearContents

    ' Verifica si la hoja activa es una hoja autorizada para usar la macro.
    If ws.Name <> "CONCESI�N GIJ�N" And ws.Name <> "CONCESI�N SOTO" And ws.Name <> "CONCESI�N OVIEDO" Then
        MsgBox "Esta macro solo se puede usar en las hojas de Concesi�n.", vbExclamation
        Exit Sub
    End If

    ' Determina la celda a omitir seg�n la hoja activa.
    Dim celdaOmitir As String
    Select Case ws.Name
        Case "CONCESI�N GIJ�N"
            celdaOmitir = "$C$9"
        Case "CONCESI�N SOTO"
            celdaOmitir = "$C$8"
        Case "CONCESI�N OVIEDO"
            celdaOmitir = ""
        Case Else
            celdaOmitir = ""
    End Select

    ' Obtiene el rango visible A1:F40 (sin incluir filas o columnas ocultas)
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.Range("A1:F40").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "No hay celdas visibles en el rango especificado.", vbExclamation
        Exit Sub
    End If

    ' Construcci�n del texto a copiar, respetando saltos de fila y omisi�n de celda
    Dim strTexto As String
    Dim fila As Range, celda As Range
    Dim filaTexto As String

    For Each fila In rng.Rows
        filaTexto = ""
        For Each celda In fila.Cells
            ' Solo incluye celdas visibles y no ocultas
            If Not celda.EntireRow.Hidden And Not celda.EntireColumn.Hidden Then
                ' Omite la celda indicada si corresponde
                If celdaOmitir <> "" And celda.Address = celdaOmitir Then
                    filaTexto = filaTexto & vbTab
                Else
                    ' Usa .Text para conservar formatos, s�mbolos y moneda (�)
                    filaTexto = filaTexto & celda.text & vbTab
                End If
            End If
        Next celda
        ' Elimina el tabulador final de la fila y a�ade salto de l�nea
        If Len(filaTexto) > 0 Then
            filaTexto = Left(filaTexto, Len(filaTexto) - 1)
            strTexto = strTexto & filaTexto & vbCrLf
        End If
    Next fila

    ' Elimina el salto de l�nea extra al final del texto si corresponde
    If Len(strTexto) > 0 Then
        strTexto = Left(strTexto, Len(strTexto) - 2)
    End If

    ' Env�a el resultado al portapapeles usando la funci�n API
    CopyToClipboard_API strTexto

    MsgBox "Datos copiados al portapapeles.", vbInformation
End Sub

'========================================================================================
' Subrutina auxiliar: CopyToClipboard_API
' Copia texto Unicode al portapapeles de Windows utilizando llamadas API para m�xima compatibilidad.
'========================================================================================
Sub CopyToClipboard_API(texto As String)
    Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr, lngResult As LongPtr
    Dim lngLength As Long

    ' Calcula la cantidad de bytes necesarios para almacenar el texto Unicode (2 bytes por car�cter + null)
    lngLength = (Len(texto) + 1) * 2

    ' Reserva memoria global para el texto
    hGlobalMemory = GlobalAlloc(GHND, lngLength)
    If hGlobalMemory = 0 Then
        MsgBox "No se pudo asignar memoria para el portapapeles.", vbCritical
        Exit Sub
    End If

    ' Bloquea la memoria para obtener un puntero seguro
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    If lpGlobalMemory = 0 Then
        MsgBox "No se pudo bloquear la memoria para el portapapeles.", vbCritical
        Exit Sub
    End If

    ' Copia el texto Unicode a la memoria reservada
    CopyMemory ByVal lpGlobalMemory, ByVal StrPtr(texto), lngLength

    ' Libera el bloqueo de la memoria (la memoria sigue reservada para el portapapeles)
    GlobalUnlock hGlobalMemory

    ' Intenta abrir el portapapeles de Windows
    If OpenClipboard(0) = 0 Then
        MsgBox "No se pudo abrir el portapapeles.", vbCritical
        Exit Sub
    End If

    ' Vac�a el portapapeles antes de escribir (buena pr�ctica)
    If EmptyClipboard() = 0 Then
        MsgBox "No se pudo vaciar el portapapeles.", vbCritical
        CloseClipboard
        Exit Sub
    End If

    ' Copia la memoria al portapapeles en formato Unicode
    lngResult = SetClipboardData(CF_UNICODETEXT, hGlobalMemory)
    If lngResult = 0 Then
        MsgBox "No se pudo establecer los datos en el portapapeles.", vbCritical
        CloseClipboard
        Exit Sub
    End If

    ' Cierra el portapapeles correctamente
    CloseClipboard
End Sub
