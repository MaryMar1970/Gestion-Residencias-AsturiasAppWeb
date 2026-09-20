Attribute VB_Name = "EvaluacionSolicitudes"
Option Explicit

' === VARIABLES GLOBALES PARA CACHÉ ===
Private dictEval As Object
Private ultimaActualizacion As Date
Private Const MINUTOS_CACHE As Long = 30 ' Cache válido por 30 minutos

Public Function ObtenerResultado(finalidad As String, empleo As String, situacion As String) As String
    Dim fTrim As String, eTrim As String, sTrim As String
    fTrim = Trim(finalidad)
    eTrim = Trim(empleo)
    sTrim = Trim(situacion)
    
    If fTrim = "" Or eTrim = "" Or sTrim = "" Then
        ObtenerResultado = ""
        Exit Function
    End If
    
    If dictEval Is Nothing Or (Now - ultimaActualizacion) > TimeSerial(0, MINUTOS_CACHE, 0) Then
        Set dictEval = CrearDiccionarioEvaluacion()
        ultimaActualizacion = Now
    End If

    Dim clave As String
    clave = NormalizarClave(fTrim) & "|" & NormalizarClave(eTrim) & "|" & NormalizarClave(sTrim)

    If dictEval.Exists(clave) Then
        ObtenerResultado = dictEval(clave)
    Else
        ObtenerResultado = "No válido"
    End If
End Function

Public Sub RecargarDiccionarioEvaluacion()
    Set dictEval = CrearDiccionarioEvaluacion()
    ultimaActualizacion = Now
End Sub

Public Function NormalizarClave(ByVal txt As String) As String
    Dim s As String, i As Long, ch As String, cCode As Long
    s = LCase(Trim(txt))
    Dim res As String: res = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        cCode = AscW(ch)
        Select Case cCode
            Case 225, 224, 228, 226, 227 ' á, à, ä, â, ã
                res = res & "a"
            Case 233, 232, 235, 234       ' é, è, ë, ê
                res = res & "e"
            Case 237, 236, 239, 238       ' í, ì, ï, î
                res = res & "i"
            Case 243, 242, 246, 244, 245 ' ó, ò, ö, ô, õ
                res = res & "o"
            Case 250, 249, 252, 251       ' ú, ù, ü, û
                res = res & "u"
            Case 65533, 63                 ' Carácter de sustitución corrupto (?)
                res = res & ""
            Case Else
                res = res & ch
        End Select
    Next i
    NormalizarClave = res
End Function

Private Function ObtenerHojaEvaluacion() As Worksheet
    On Error Resume Next
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        If sh.CodeName = "Hoja15" Then
            Set ObtenerHojaEvaluacion = sh
            Exit Function
        End If
    Next sh
    Set ObtenerHojaEvaluacion = ThisWorkbook.Sheets("Evaluación")
    On Error GoTo 0
End Function

Private Function CrearDiccionarioEvaluacion() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim arrDatos As Variant
    Dim clave As String
    
    On Error GoTo ErrorHandler
    Set ws = ObtenerHojaEvaluacion()
    If ws Is Nothing Then
        Set CrearDiccionarioEvaluacion = d
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        Set CrearDiccionarioEvaluacion = d
        Exit Function
    End If
    
    arrDatos = ws.Range("A2:D" & lastRow).Value
    
    For i = 1 To UBound(arrDatos, 1)
        If Trim(CStr(arrDatos(i, 1))) <> "" And Trim(CStr(arrDatos(i, 2))) <> "" And Trim(CStr(arrDatos(i, 3))) <> "" Then
            clave = NormalizarClave(CStr(arrDatos(i, 1))) & "|" & _
                    NormalizarClave(CStr(arrDatos(i, 2))) & "|" & _
                    NormalizarClave(CStr(arrDatos(i, 3)))
            d(clave) = Trim(CStr(arrDatos(i, 4)))
        End If
    Next i
    
    Set CrearDiccionarioEvaluacion = d
    Exit Function
    
ErrorHandler:
    MsgBox "Error al crear diccionario de evaluación: " & Err.Description, vbCritical
    Set CrearDiccionarioEvaluacion = d
End Function

Public Sub ActualizarCacheEvaluacion()
    Call RecargarDiccionarioEvaluacion
    MsgBox "Caché de evaluación actualizado correctamente.", vbInformation
End Sub
