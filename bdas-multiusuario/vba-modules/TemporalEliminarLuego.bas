Attribute VB_Name = "TemporalEliminarLuego"
Sub DiagnosticarListaHojas()
    Dim listaHojas As String
    Dim arrLista() As String
    Dim ws As Worksheet
    Dim nombreHoja As String
    Dim i As Long
    Dim encontrada As Boolean
    
    ' === TU LISTA ACTUAL ===
    listaHojas = "Habitaciones Bloqueadas OVIEDO,Festivos,"
    listaHojas = listaHojas & "FACTURACI�N OVIEDO,FACTURA OVIEDO,SOBRE OVIEDO,"
    listaHojas = listaHojas & "PLANTILLAS OVIEDO,LOG_OVIEDO,"
    listaHojas = listaHojas & "Lavander�a Oviedo,HISTORICO_LAVANDERIA_OVIEDO,"
    listaHojas = listaHojas & "Habitaciones Bloqueadas GIJ�N,RESIDENCIA ESTUDIANTES,FACTURA R. ESTUDIANTES,"
    listaHojas = listaHojas & "LIQUIDACION R. ESTUDIANTES,FACTURACI�N GIJ�N,FACTURA GIJ�N,SOBRE GIJ�N,"
    listaHojas = listaHojas & "PLANTILLAS GIJ�N,LOG_GIJ�N,"
    listaHojas = listaHojas & "Lavander�a Gij�n,HISTORICO_LAVANDERIA_GIJON,"
    listaHojas = listaHojas & "Habitaciones Bloqueadas SOTO,"
    listaHojas = listaHojas & "FACTURACI�N SOTO,FACTURA SOTO,"
    listaHojas = listaHojas & "PLANTILLAS SOTO,LOG_SOTO,"
    listaHojas = listaHojas & "Lavander�a Soto,HISTORICO_LAVANDERIA_SOTO,"
    listaHojas = listaHojas & "TARIFAS UNIDAD,Turnos-Precios,LISTA NEGRA,"
    listaHojas = listaHojas & "USUARIOS,CODIGOS POSTALES,Auxiliar,Evaluaci�n,"
    listaHojas = listaHojas & "INICIO,CONFIG"
    
    arrLista = Split(listaHojas, ",")
    
    Debug.Print "========================================="
    Debug.Print "DIAGN�STICO: Hojas en LISTA que NO EXISTEN"
    Debug.Print "========================================="
    
    For i = LBound(arrLista) To UBound(arrLista)
        nombreHoja = Trim(arrLista(i))
        If nombreHoja <> "" Then
            encontrada = False
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name = nombreHoja Then
                    encontrada = True
                    Exit For
                End If
            Next ws
            
            If Not encontrada Then
                Debug.Print "? NO EXISTE: [" & nombreHoja & "]"
            End If
        End If
    Next i
    
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "HOJAS que EXISTEN pero NO est�n en LISTA"
    Debug.Print "========================================="
    
    For Each ws In ThisWorkbook.Worksheets
        ' Excluir hojas que NO deben ocultarse
        If ws.Name <> "RESIDENCIA GIJ�N" And ws.Name <> "RESUMEN GIJ�N" And ws.Name <> "Calendario GIJ�N" And _
           ws.Name <> "RESIDENCIA SOTO" And ws.Name <> "RESUMEN SOTO" And ws.Name <> "Calendario SOTO" And _
           ws.Name <> "RESIDENCIA OVIEDO" And ws.Name <> "RESUMEN OVIEDO" And ws.Name <> "Calendario OVIEDO" Then
            
            encontrada = False
            For i = LBound(arrLista) To UBound(arrLista)
                If Trim(arrLista(i)) = ws.Name Then
                    encontrada = True
                    Exit For
                End If
            Next i
            
            If Not encontrada Then
                Debug.Print "?? FALTA EN LISTA: [" & ws.Name & "]"
            End If
        End If
    Next ws
    
    Debug.Print ""
    Debug.Print "========================================="
    Debug.Print "DIAGN�STICO COMPLETADO"
    Debug.Print "========================================="
End Sub


