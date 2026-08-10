Attribute VB_Name = "EvaluacionSolicitudes"
Option Explicit

' === VARIABLES GLOBALES PARA CACH� ===
Private dictEval As Object
Private ultimaActualizacion As Date
Private Const MINUTOS_CACHE As Long = 30 ' Cache v�lido por 30 minutos

Function ObtenerResultado(finalidad As String, empleo As String, situacion As String) As String
    If finalidad = "" Or empleo = "" Or situacion = "" Then
        ObtenerResultado = ""
        Exit Function
    End If
    
    If dictEval Is Nothing Or (Now - ultimaActualizacion) > TimeSerial(0, MINUTOS_CACHE, 0) Then
        Set dictEval = CrearDiccionarioEvaluacion()
        ultimaActualizacion = Now
    End If

    Dim clave As String
    clave = LCase(Trim(finalidad)) & "|" & LCase(Trim(empleo)) & "|" & LCase(Trim(situacion))

    If dictEval.Exists(clave) Then
        ObtenerResultado = dictEval(clave)
    Else
        ObtenerResultado = "No v�lido"
    End If
End Function

Sub RecargarDiccionarioEvaluacion()
    Set dictEval = CrearDiccionarioEvaluacion()
    ultimaActualizacion = Now
End Sub

Private Function CrearDiccionarioEvaluacion() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim arrDatos As Variant
    Dim clave As String
    
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets("Evaluaci�n")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        Set CrearDiccionarioEvaluacion = d
        Exit Function
    End If
    
    arrDatos = ws.Range("A2:D" & lastRow).Value
    
    For i = 1 To UBound(arrDatos, 1)
        If arrDatos(i, 1) <> "" And arrDatos(i, 2) <> "" And arrDatos(i, 3) <> "" Then
            clave = LCase(Trim(CStr(arrDatos(i, 1)))) & "|" & _
                    LCase(Trim(CStr(arrDatos(i, 2)))) & "|" & _
                    LCase(Trim(CStr(arrDatos(i, 3))))
            d(clave) = CStr(arrDatos(i, 4))
        End If
    Next i
    
    Set CrearDiccionarioEvaluacion = d
    Exit Function
    
ErrorHandler:
    MsgBox "Error al crear diccionario de evaluaci�n: " & Err.Description, vbCritical
    Set CrearDiccionarioEvaluacion = d
End Function

Sub ActualizarCacheEvaluacion()
    Call RecargarDiccionarioEvaluacion
    MsgBox "Cach� de evaluaci�n actualizado correctamente.", vbInformation
End Sub
