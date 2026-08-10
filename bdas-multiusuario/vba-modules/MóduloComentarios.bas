Attribute VB_Name = "M�duloComentarios"
' Macro para insertar o modificar comentario en la celda activa
Sub ComentarioCeldaInsertar(control As IRibbonControl)
    Dim ws As Worksheet
    Dim celda As Range
    Set ws = ActiveSheet
    Set celda = ActiveCell

    If celda.Locked Then
        MsgBox "La celda est� bloqueada. No puedes insertar comentarios aqu�.", vbExclamation
        Exit Sub
    End If

    Dim comentario As String
    On Error Resume Next
    comentario = celda.Comment.text
    On Error GoTo 0

    comentario = InputBox("Escribe o edita el comentario:", "Insertar Comentario", comentario)
    If comentario = "" Then Exit Sub

    On Error GoTo ErrHandler

    ' DESPROTEGER HOJA temporalmente
    ws.Unprotect password:=""

    celda.ClearComments
    celda.AddComment text:=comentario

    ' VOLVER A PROTEGER HOJA con UserInterfaceOnly para que macros puedan seguir funcionando
    ws.Protect password:="", UserInterfaceOnly:=True

    Exit Sub

ErrHandler:
    MsgBox "Error al insertar comentario: " & Err.Description, vbCritical
End Sub

' Macro para borrar comentario en la celda activa
Sub ComentarioCeldaBorrar(control As IRibbonControl)
    Dim ws As Worksheet
    Dim celda As Range
    Set ws = ActiveSheet
    Set celda = ActiveCell

    On Error GoTo ErrHandler

    ws.Unprotect password:=""

    If Not celda.Comment Is Nothing Then
        celda.ClearComments
        MsgBox "Comentario eliminado.", vbInformation
    Else
        MsgBox "La celda no tiene comentario.", vbExclamation
    End If

    ws.Protect password:="", UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error al eliminar comentario: " & Err.Description, vbCritical
End Sub

'=================================================================================
' MACRO: CambiarSede
' Uso: Alt+F8 ? CambiarSede ? Ejecutar
'=================================================================================
Public Sub CambiarSede()

    Dim residenciaActiva As String
    Dim arrHojas As Variant
    Dim arrTodasLasPrincipales As Variant
    Dim hojaInicial As String
    Dim i As Long
    Dim ws As Worksheet

    Dim opcion As String
    Dim msg As String
    msg = "Selecciona la sede a mostrar:" & vbCrLf & vbCrLf & _
          "  1  ->  TODAS" & vbCrLf & _
          "  2  ->  GIJ�N" & vbCrLf & _
          "  3  ->  SOTO" & vbCrLf & _
          "  4  ->  OVIEDO" & vbCrLf & _
          "  5  ->  GIJ�N + SOTO"

    opcion = InputBox(msg, "Cambiar Sede", "")

    Select Case Trim(opcion)
        Case "1": residenciaActiva = "TODAS"
        Case "2": residenciaActiva = "GIJON"
        Case "3": residenciaActiva = "SOTO"
        Case "4": residenciaActiva = "OVIEDO"
        Case "5": residenciaActiva = "GIJON-SOTO"
        Case "":  Exit Sub
        Case Else
            MsgBox "Opci�n no v�lida. Introduce un n�mero del 1 al 5.", vbExclamation, "Cambiar Sede"
            Exit Sub
    End Select

    On Error Resume Next
    ThisWorkbook.Sheets("CONFIG").Range("B10").Value = residenciaActiva
    On Error GoTo 0

    Select Case residenciaActiva
        Case "OVIEDO"
            arrHojas = Array("RESIDENCIA OVIEDO", "RESUMEN OVIEDO", "Calendario OVIEDO")
            hojaInicial = "RESIDENCIA OVIEDO"
        Case "GIJON"
            arrHojas = Array("RESIDENCIA GIJ�N", "RESUMEN GIJ�N", "Calendario GIJ�N")
            hojaInicial = "RESIDENCIA GIJ�N"
        Case "SOTO"
            arrHojas = Array("RESIDENCIA SOTO", "RESUMEN SOTO", "Calendario SOTO")
            hojaInicial = "RESIDENCIA SOTO"
        Case "GIJON-SOTO"
            arrHojas = Array( _
                "RESIDENCIA GIJ�N", "RESUMEN GIJ�N", "Calendario GIJ�N", _
                "RESIDENCIA SOTO", "RESUMEN SOTO", "Calendario SOTO")
            hojaInicial = "RESIDENCIA GIJ�N"
        Case Else
            arrHojas = Array( _
                "RESIDENCIA OVIEDO", "RESUMEN OVIEDO", "Calendario OVIEDO", _
                "RESIDENCIA GIJ�N", "RESUMEN GIJ�N", "Calendario GIJ�N", _
                "RESIDENCIA SOTO", "RESUMEN SOTO", "Calendario SOTO")
            hojaInicial = "RESIDENCIA GIJ�N"
    End Select

    arrTodasLasPrincipales = Array( _
        "RESIDENCIA OVIEDO", "RESUMEN OVIEDO", "Calendario OVIEDO", _
        "RESIDENCIA GIJ�N", "RESUMEN GIJ�N", "Calendario GIJ�N", _
        "RESIDENCIA SOTO", "RESUMEN SOTO", "Calendario SOTO")

    Application.screenUpdating = False
    Application.enableEvents = False

    If SheetExists(hojaInicial) Then
        ThisWorkbook.Worksheets(hojaInicial).visible = xlSheetVisible
        ThisWorkbook.Worksheets(hojaInicial).Activate
    End If

    Dim estaEnLista As Boolean
    Dim j As Long

    For i = LBound(arrTodasLasPrincipales) To UBound(arrTodasLasPrincipales)
        estaEnLista = False
        For j = LBound(arrHojas) To UBound(arrHojas)
            If arrTodasLasPrincipales(i) = arrHojas(j) Then
                estaEnLista = True
                Exit For
            End If
        Next j
        If Not estaEnLista Then
            If SheetExists(CStr(arrTodasLasPrincipales(i))) Then
                On Error Resume Next
                ThisWorkbook.Worksheets(CStr(arrTodasLasPrincipales(i))).visible = xlSheetVeryHidden
                On Error GoTo 0
            End If
        End If
    Next i

    For i = LBound(arrHojas) To UBound(arrHojas)
        If SheetExists(CStr(arrHojas(i))) Then
            On Error Resume Next
            ThisWorkbook.Worksheets(CStr(arrHojas(i))).visible = xlSheetVisible
            On Error GoTo 0
        End If
    Next i

    If SheetExists(hojaInicial) Then
        ThisWorkbook.Worksheets(hojaInicial).Activate
    End If

    Application.enableEvents = True
    Application.screenUpdating = True

    MsgBox "Sede cambiada a: " & residenciaActiva, vbInformation, "CambiarSede"

End Sub

Private Function SheetExists(ByVal nombre As String) As Boolean
    If Trim(nombre) = "" Then Exit Function
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nombre)
    SheetExists = (Err.Number = 0 And Not ws Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

