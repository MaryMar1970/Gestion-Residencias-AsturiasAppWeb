VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReservasUnidad 
   Caption         =   "Generador de textos para RESIDENCIA UNIDAD"
   ClientHeight    =   10200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8130
   OleObjectBlob   =   "frmReservasUnidad.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmReservasUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' UserForm: frmReservasUnidad
' Versi�n: Optimizada y robusta
' - Centraliza l�gica de columnas por hoja
' - Normaliza comprobaciones (SOTO vs GIJ�N/OVIEDO)
' - Rellena txtPass, txtTurno, txtComision seg�n reglas
' - Programa/cancela cierre autom�tico de forma segura
' - Documentaci�n extensa para f�cil mantenimiento
' =============================================================================

Option Explicit

' -------------------------------------------------------------------------
' Variables de instancia / estado del formulario
' -------------------------------------------------------------------------
Private cerrarTimer As Double         ' instante programado para cierre autom�tico (Application.OnTime)
Private mFilaSeleccionada As Long     ' fila encontrada por N� ORDEN
Private mResolucionOriginal As String ' valor original de P (RESOLUCI�N) normalizado
Private mClavePlantilla As String     ' clave de plantilla (CONCEDIDA / DESESTIMADA / ...)
Private mUnidad As String             ' p.ej. "GIJ�N", "SOTO"
Private mHojaReservas As String       ' nombre exacto de la hoja origen (p.ej. "RESIDENCIA GIJ�N")

' -------------------------------------------------------------------------
' Constantes columna (valores por defecto para algunas hojas)
' - NOTA: para SOTO la "comisi�n" no est� en H; la comprobaci�n se hace en D.
' -------------------------------------------------------------------------
Private Const COL_N_ORDEN As String = "A"
Private Const COL_F_SOLIC As String = "B"
Private Const COL_FINALID As String = "D"   ' FINALIDAD
Private Const COL_COMISION As String = "H"  ' En GIJ�N/OVIEDO la columna COMISI�N est� en H
Private Const COL_RANGO As String = "J"
Private Const COL_NOMBRE As String = "K"
Private Const COL_F_ENTRA As String = "L"
Private Const COL_F_SALID As String = "M"
Private Const COL_RESOLUC As String = "P"

' =============================================================================
' BOTONES
' =============================================================================

Private Sub btnCerrar_Click()
    ' Cierre seguro del formulario (cancela temporizador si existiera)
    CerrarFormulario
End Sub

Private Sub btnGenerarTexto_Click()
    ' Solo generar si la plantilla corresponde a DESESTIMADA (comportamiento original)
    If mClavePlantilla = "DESESTIMADA" Then
        GenerarTexto
    End If
End Sub

' =============================================================================
' CIERRE / TEMPORIZADOR
' =============================================================================

Public Sub CerrarFormulario()
    ' Cierra el formulario y cancela cualquier temporizador programado
    On Error Resume Next
    CancelScheduledClose
    Me.Hide
    DoEvents
    Unload Me
End Sub

Private Sub CancelScheduledClose()
    ' SOLUCI�N v2.0: No intentar cancelar OnTime - simplemente limpiar variable
    ' El temporizador se ejecutar� pero encontrar� que el formulario ya est� cerrado
    cerrarTimer = 0
End Sub

' =============================================================================
' INICIALIZACI�N DEL FORMULARIO
' =============================================================================

Private Sub UserForm_Initialize()
    ' Poblaci�n inicial del combobox de motivos (DESESTIMADA)
    With cbMotivo
        .Clear
        .AddItem "Solicitud de +7 d�as de alojamiento"
        .AddItem "Solicitud efectuada con +30 d�as naturales de antelaci�n a la fecha de entrada"
        .AddItem "La solicitud excede la capacidad m�xima del alojamiento"
        .AddItem "La Residencia no dispone de camas supletorias ni habitaciones triples o cu�druple"
        .AddItem "Solicitud duplicada (id�ntico DNI, fecha de entrada/salida y n�mero de personas)"
        .AddItem "La solicitud no est� correctamente cumplimentada (revisar legibilidad, ausencia de datos y/o firma, incoherencia en fechas y/o Residencia peticionada)"
        .AddItem "Incumplimiento de la normativa relativa a beneficiarios (c�nyuge e hijos y, excepcionalmente de manera acreditada, familiares 2� grado o personas presten cuidados)"
        .AddItem "Solicitud de fechas solapadas con Turnos (Navidad, S. Santa o Verano) y, por tanto, no siendo esta Unidad competente para su adjudicaci�n"
        .AddItem "Residencia completa en las fechas solicitadas"
        .AddItem "Otros:"
    End With

    ' Ocultar controles dependientes al inicio y limpiar campos
    ShowMotivosControls False
    ShowHabitacionesControls False
    btnGenerarTexto.visible = False
    ClearAllFields
    cerrarTimer = 0
End Sub

' =============================================================================
' WRAPPER PRINCIPAL: llamada desde la hoja (cuando se hace click en AC)
' =============================================================================

Public Sub GenerarTextoDesdeOrden(ByVal nroOrden As Long, ByVal Unidad As String, ByVal nombreHoja As String)
    ' Guarda contexto de la invocaci�n y lanza procesamiento
    mUnidad = Unidad
    mHojaReservas = nombreHoja

    ' Mostrar la unidad en el formulario (texto informativo)
    On Error Resume Next
    Me.lblUnidad.Caption = "Residencia de Unidad de " & mUnidad
    On Error GoTo 0

    Me.txtNroOrden.Value = CStr(nroOrden)
    ProcesarNroOrden
End Sub

' =============================================================================
' PROCESO PRINCIPAL: validaciones y control visual
' =============================================================================

Private Sub ProcesarNroOrden()
    ' Valida n� orden, localiza fila, carga controles y decide visibilidades / temporizador.
    Dim nro As Variant
    Dim ws As Worksheet

    nro = Trim$(txtNroOrden.Value)
    If Len(nro) = 0 Then Exit Sub
    If Not IsNumeric(nro) Then
        MsgBox "Introduzca un N� ORDEN num�rico.", vbExclamation
        txtNroOrden.SetFocus
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(mHojaReservas)
    mFilaSeleccionada = FindRowByOrder(ws, CLng(nro))
    If mFilaSeleccionada = 0 Then
        MsgBox "N� ORDEN no encontrado en " & mHojaReservas & ".", vbExclamation
        Exit Sub
    End If

    ' Cargar los valores en los controles del formulario
    LoadFieldsFromRow ws, mFilaSeleccionada

    ' Guardar resoluci�n y clave de plantilla (normalizada)
    mResolucionOriginal = UCase$(Trim$(CStr(ws.Cells(mFilaSeleccionada, COL_RESOLUC).Value)))
    mClavePlantilla = NormalizeResolutionToClave(mResolucionOriginal)

    ' Mostrar/ocultar controles espec�ficos de la hoja SOTO (Turno)
    If mHojaReservas = "RESIDENCIA SOTO" Then
        lblTurno.visible = True
        txtTurno.visible = True
    Else
        lblTurno.visible = False
        txtTurno.visible = False
        txtTurno.Value = ""
    End If

    ' Mostrar/ocultar PASS si es OVIEDO (no mostrar)
    If mHojaReservas = "RESIDENCIA OVIEDO" Then
        lblPass.visible = False
        txtPass.visible = False
    Else
        lblPass.visible = True
        txtPass.visible = True
    End If

    ' Visibilidad de IMPORTES / HABITACIONES / PASS seg�n resoluci�n (SI / CONCEDIDA / REEVALUADA)
    Select Case UCase$(Trim$(mResolucionOriginal))
        Case "SI", "CONCEDIDA", "REEVALUADA"
            lblImporte.visible = True
            txtImporte.visible = True
            lblHabitaciones.visible = True
            txtHabitaciones.visible = True
            lblPass.visible = True
            txtPass.visible = True
        Case Else
            lblImporte.visible = False
            txtImporte.visible = False
            lblHabitaciones.visible = False
            txtHabitaciones.visible = False
            lblPass.visible = False
            txtPass.visible = False
    End Select

    ' Si la plantilla es DESESTIMADA, se deja abierto y se muestran motivos; si NO, se genera texto y se programa cierre
    If mClavePlantilla = "DESESTIMADA" Then
        ShowMotivosControls True
        btnGenerarTexto.visible = True
        CancelScheduledClose
        cerrarTimer = 0
    Else
        ShowMotivosControls False
        btnGenerarTexto.visible = False
        ' Generar texto (rellena txtTextoFinal y copia al portapapeles)
        GenerarTexto
        ' Programar cierre autom�tico (segundos ajustables)
        CancelScheduledClose
        cerrarTimer = Now + (1 / 86400)  ' <-- cambiar 1 por n�mero de segundos deseados
        Application.OnTime cerrarTimer, "CerrarReservasUnidadFlotante"
    End If
End Sub

' =============================================================================
' GENERAR TEXTO: busca plantilla, calcula valores y hace los REPLACE
' =============================================================================

Private Sub GenerarTexto()
    ' Construye el texto final reemplazando placeholders por valores reales.
    Dim wsP As Worksheet, wsR As Worksheet
    Dim filaPlantilla As Long, UltFila As Long
    Dim plantilla As String, texto As String
    Dim importeVal As Variant, importeStr As String
    Dim habStr As String, passStr As String, otros As String
    Dim i As Long, colHab As String, colImporte As String
    Dim finalidad As String, comi As String, turnoVal As String

    If mFilaSeleccionada = 0 Or Len(Trim$(mClavePlantilla)) = 0 Then Exit Sub
    Set wsR = ThisWorkbook.Worksheets(mHojaReservas)

    ' -------------------------
    ' Determinar 'comi': desde columna D (FINALIDAD), para todas las hojas
    ' SI = "Comisi�n" / "Comisi�n NO indem." / "Destino"
    ' NO = cualquier otro valor
    ' -------------------------
    Dim finalidadD As String
    finalidadD = Trim(CStr(wsR.Cells(mFilaSeleccionada, COL_FINALID).Value))
    
    If finalidadD = "Comisi�n" Or finalidadD = "Comisi�n NO indem." Or finalidadD = "Destino" Then
        comi = "SI"
    Else
        comi = "NO"
    End If

    ' -------------------------
    ' Obtener la plantilla correcta seg�n la hoja
    ' -------------------------
    Select Case mHojaReservas
        Case "RESIDENCIA GIJ�N": Set wsP = ThisWorkbook.Worksheets("PLANTILLAS GIJ�N")
        Case "RESIDENCIA OVIEDO": Set wsP = ThisWorkbook.Worksheets("PLANTILLAS OVIEDO")
        Case "RESIDENCIA SOTO": Set wsP = ThisWorkbook.Worksheets("PLANTILLAS SOTO")
        Case Else: Set wsP = ThisWorkbook.Worksheets("PLANTILLAS GIJ�N")
    End Select

    ' Buscar fila de plantilla
    UltFila = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    filaPlantilla = 0
    For i = 2 To UltFila
        If UCase$(Trim$(wsP.Cells(i, "A").Value)) = UCase$(Trim$(mClavePlantilla)) Then
            filaPlantilla = i
            Exit For
        End If
    Next i
    If filaPlantilla = 0 Then Exit Sub
    plantilla = CStr(wsP.Cells(filaPlantilla, "C").Value)

    ' -------------------------
    ' Preparar datos (fechas, nombre, importe, habitaciones)
    ' -------------------------
    Dim fSol As String, fEnt As String, fSal As String, nombre As String
    fSol = SafeFormatDate(wsR.Cells(mFilaSeleccionada, COL_F_SOLIC).Value)
    fEnt = SafeFormatDate(wsR.Cells(mFilaSeleccionada, COL_F_ENTRA).Value)
    fSal = SafeFormatDate(wsR.Cells(mFilaSeleccionada, COL_F_SALID).Value)
    nombre = CStr(wsR.Cells(mFilaSeleccionada, COL_NOMBRE).Value)

    GetColsForHoja mHojaReservas, colHab, colImporte
    habStr = CStr(wsR.Cells(mFilaSeleccionada, colHab).Value)
    importeVal = wsR.Cells(mFilaSeleccionada, colImporte).Value
    If IsNumeric(importeVal) Then
        importeStr = Replace(Format(importeVal, "0.00"), ".", ",")
    Else
        importeStr = ""
    End If

    ' -------------------------
    ' Obtener contrase�as (si aplica) y asignarlas a txtPass
    ' - Gij�n: ObtenerContrasenasGijon
    ' - Soto: ObtenerContrasenasSoto (si existe)
    ' -------------------------
    passStr = ""
    If Len(Trim$(habStr)) > 0 Then
        On Error Resume Next
        If mHojaReservas = "RESIDENCIA GIJ�N" Then
            passStr = ObtenerContrasenasGijon(habStr)
        ElseIf mHojaReservas = "RESIDENCIA SOTO" Then
            ' Si dispones de funci�n para SOTO, la llamamos; si no, queda cadena vac�a
            passStr = ObtenerContrasenasSoto(habStr)
        End If
        On Error GoTo 0
    End If
    ' Reflejar pass en el control del formulario (si est� visible)
    On Error Resume Next
    Me.txtPass.Value = passStr
    On Error GoTo 0

    ' -------------------------
    ' Determinar valor de [[TURNO]]:
    ' - Si txtTurno tiene valor (cargado en LoadFieldsFromRow), lo usamos.
    ' - Si no, intentar leer desde la columna de turno (GetColTurno)
    ' -------------------------
    turnoVal = Trim$(Me.txtTurno.Value)
    If Len(turnoVal) = 0 Then
        Dim colTurno As String
        colTurno = GetColTurno(mHojaReservas)
        If Len(colTurno) > 0 Then turnoVal = CStr(wsR.Cells(mFilaSeleccionada, colTurno).Value)
    End If

    ' -------------------------
    ' Motivos DESESTIMADA (construcci�n de texto adicional)
    ' -------------------------
    otros = ""
    If mClavePlantilla = "DESESTIMADA" Then
        Dim motivosSeleccionados As String
        motivosSeleccionados = ""
        For i = 0 To cbMotivo.ListCount - 1
            If cbMotivo.Selected(i) Then
                If Len(motivosSeleccionados) > 0 Then motivosSeleccionados = motivosSeleccionados & "; "
                motivosSeleccionados = motivosSeleccionados & cbMotivo.List(i)
            End If
        Next i
        If Len(motivosSeleccionados) > 0 Then otros = motivosSeleccionados
        If Len(Trim$(txtOtrosMotivos.Value)) > 0 Then
            If Len(otros) > 0 Then otros = otros & " "
            otros = otros & Trim$(txtOtrosMotivos.Value)
        End If
    End If

    ' -------------------------
    ' Reemplazos en plantilla (placeholders)
    ' -------------------------
    texto = plantilla
    texto = Replace(texto, "[[REGISTRO]]", Trim$(txtNroOrden.Value))
    texto = Replace(texto, "[[NOMBRE]]", nombre)
    texto = Replace(texto, "[[FECHA_SOLICITUD]]", fSol)
    texto = Replace(texto, "[[COMISION]]", comi)           ' comi ya contiene SI/NO o valor real
    texto = Replace(texto, "[[ENTRADA]]", fEnt)
    texto = Replace(texto, "[[SALIDA]]", fSal)
    texto = Replace(texto, "[[HABITACION]]", habStr)
    texto = Replace(texto, "[[PASWORD]]", passStr)
    texto = Replace(texto, "[[RESOLUCION]]", mResolucionOriginal)
    texto = Replace(texto, "[[IMPORTE]]", importeStr)
    texto = Replace(texto, "[[OTROS_MOTIVOS]]", otros)
    texto = Replace(texto, "[[UNIDAD]]", mUnidad)
    texto = Replace(texto, "[[TURNO]]", turnoVal)         ' Nuevo placeholder [[TURNO]]

    ' Mostrar resultado y copiar al portapapeles solo si no est� vac�o
    txtTextoFinal.Value = texto
    If Trim(txtTextoFinal.Value) <> "" Then
        CopiarTextoAlPortapapelesAPI txtTextoFinal.Value
    End If
End Sub

' =============================================================================
' HELPERS / UTILIDADES
' =============================================================================

Private Function FindRowByOrder(ws As Worksheet, ByVal nro As Long) As Long
    ' Busca en la columna A (COL_N_ORDEN) la coincidencia exacta y devuelve la fila
    Dim ult As Long, m As Variant
    ult = ws.Cells(ws.Rows.Count, COL_N_ORDEN).End(xlUp).Row
    If ult < 2 Then
        FindRowByOrder = 0
        Exit Function
    End If
    m = Application.Match(nro, ws.Range(COL_N_ORDEN & "2:" & COL_N_ORDEN & ult), 0)
    If IsError(m) Then
        FindRowByOrder = 0
    Else
        FindRowByOrder = CLng(m) + 1
    End If
End Function

Private Sub LoadFieldsFromRow(ws As Worksheet, ByVal fila As Long)
    ' Rellena los controles del formulario con los datos de la fila proporcionada.
    ' Aplica la regla especial para SOTO: txtComision = "SI"/"NO" seg�n FINALIDAD (col D).
    Dim colHab As String, colImporte As String
    Dim finalidadRaw As String

    txtNombre.Value = CStr(ws.Cells(fila, COL_NOMBRE).Value)
    txtResolucion.Value = NormalizeResolutionToClave(ws.Cells(fila, COL_RESOLUC).Value)
    txtFechaSolicitud.Value = SafeFormatDate(ws.Cells(fila, COL_F_SOLIC).Value)

   ' txtComision: misma l�gica para todas las hojas, basada en columna D
    Dim finalidadComis As String
    finalidadComis = Trim(CStr(ws.Cells(fila, COL_FINALID).Value))
    
    If finalidadComis = "Comisi�n" Or finalidadComis = "Comisi�n NO indem." Or finalidadComis = "Destino" Then
        txtComision.Value = "SI"
    Else
        txtComision.Value = "NO"
    End If

    ' --------------------------
    ' txtTurno (si corresponde): preferimos cargarlo aqu�
    ' --------------------------
    Dim colTurno As String
    colTurno = GetColTurno(ws.Name)
    If Len(colTurno) > 0 Then
        txtTurno.Value = CStr(ws.Cells(fila, colTurno).Value)
    Else
        txtTurno.Value = ""
    End If

    txtFechaEntrada.Value = SafeFormatDate(ws.Cells(fila, COL_F_ENTRA).Value)
    txtFechaSalida.Value = SafeFormatDate(ws.Cells(fila, COL_F_SALID).Value)

    ' Habitaciones / importe seg�n hoja
    GetColsForHoja ws.Name, colHab, colImporte
    txtHabitaciones.Value = CStr(ws.Cells(fila, colHab).Value)
    txtImporte.Value = CStr(ws.Cells(fila, colImporte).Value)

    ' Limpieza de control de texto final / motivos
    txtTextoFinal.Value = vbNullString
    If cbMotivo.ListCount > 0 Then cbMotivo.ListIndex = -1
    txtOtrosMotivos.Value = vbNullString
End Sub

Private Sub ClearAllFields()
    ' Limpia todos los controles del formulario (estado inicial)
    txtNroOrden.Value = vbNullString
    txtNombre.Value = vbNullString
    txtResolucion.Value = vbNullString
    txtFechaSolicitud.Value = vbNullString
    txtComision.Value = vbNullString
    txtFechaEntrada.Value = vbNullString
    txtFechaSalida.Value = vbNullString
    txtHabitaciones.Value = vbNullString
    txtOtrosMotivos.Value = vbNullString
    txtTextoFinal.Value = vbNullString
    txtTurno.Value = vbNullString
    txtPass.Value = vbNullString
    If cbMotivo.ListCount > 0 Then cbMotivo.ListIndex = -1
End Sub

Private Function SafeFormatDate(v As Variant) As String
    ' Formatea fecha a dd/mm/yyyy si v es fecha, sino devuelve cadena vac�a
    If IsDate(v) Then SafeFormatDate = Format$(CDate(v), "dd/mm/yyyy") Else SafeFormatDate = ""
End Function

Private Function NormalizeResolutionToClave(ByVal res As String) As String
    ' Normaliza diferentes textos de resoluci�n a claves conocidas
    Select Case UCase$(Trim$(res))
        Case "CONCEDIDA", "SI", "REEVALUADA": NormalizeResolutionToClave = "CONCEDIDA"
        Case "DENEGADA", "NO": NormalizeResolutionToClave = "DENEGADA"
        Case "DESESTIMADA": NormalizeResolutionToClave = "DESESTIMADA"
        Case "RENUNCIA": NormalizeResolutionToClave = "RENUNCIA"
        Case Else: NormalizeResolutionToClave = UCase$(Trim$(res))
    End Select
End Function

Private Sub ShowHabitacionesControls(ByVal visible As Boolean)
    lblHabitaciones.visible = visible
    txtHabitaciones.visible = visible
End Sub

Private Sub ShowMotivosControls(ByVal visible As Boolean)
    cbMotivo.visible = visible
    txtOtrosMotivos.visible = visible
    lblMotivo.visible = visible
End Sub

' =============================================================================
' L�GICA CENTRALIZADA DE COLUMNAS / FINALIDAD
' =============================================================================

Private Sub GetColsForHoja(ByVal nombreHoja As String, ByRef colHab As String, ByRef colImporte As String)
    ' Devuelve columnas de HABITACIONES e IMPORTE seg�n la hoja
    Select Case UCase$(Trim$(nombreHoja))
        Case "RESIDENCIA GIJ�N"
            colHab = "S": colImporte = "U"
        Case "RESIDENCIA OVIEDO"
            colHab = "T": colImporte = "V"
        Case "RESIDENCIA SOTO"
            colHab = "S": colImporte = "U"
        Case Else
            colHab = "S": colImporte = "U" ' por defecto
    End Select
End Sub

Private Function GetColComision(ByVal nombreHoja As String) As String
    ' Devuelve la columna donde se encuentra 'COMISI�N' para GIJ�N/OVIEDO.
    ' En SOTO no se usa esta columna para decidir comi; sin embargo devolvemos "H"
    ' en caso de que exista un valor relacionado.
    Select Case UCase$(Trim$(nombreHoja))
        Case "RESIDENCIA GIJ�N", "RESIDENCIA OVIEDO"
            GetColComision = "H"
        Case "RESIDENCIA SOTO"
            GetColComision = "H"  ' si en SOTO hubiera algo en H, es turno; no se usa para comi
        Case Else
            GetColComision = "H"
    End Select
End Function

Private Function GetColTurno(ByVal nombreHoja As String) As String
    ' Devuelve la columna donde est� el TURNO (solo aplicable para SOTO en tu configuraci�n)
    Select Case UCase$(Trim$(nombreHoja))
        Case "RESIDENCIA SOTO"
            GetColTurno = "H" ' en SOTO la columna H es TURNO (seg�n tus notas)
        Case Else
            GetColTurno = ""  ' no aplicable en otras hojas
    End Select
End Function

Private Function IsFinalidadComision(ws As Worksheet, ByVal fila As Long) As Boolean
    ' Comprueba si la FINALIDAD (col D) est� entre las opciones que consideramos "Comisi�n".
    ' Normaliza (trim + may�sculas + elimina acentos) antes de comparar.
    Dim s As String
    s = NormalizeText(CStr(ws.Cells(fila, COL_FINALID).Value))
    Select Case s
        Case "COMISION", "COMISION NO INDEM.", "DESTINO"
            IsFinalidadComision = True
        Case Else
            IsFinalidadComision = False
    End Select
End Function

Private Function NormalizeText(ByVal s As String) As String
    ' Normaliza texto: trim, may�sculas, elimina acentos comunes
    Dim t As String
    t = UCase$(Trim$(s))
    ' Reemplazar acentos (������ si se desea). Convertir �->N (si no quieres N puedes eliminar)
    t = Replace(t, "�", "A")
    t = Replace(t, "�", "E")
    t = Replace(t, "�", "I")
    t = Replace(t, "�", "O")
    t = Replace(t, "�", "U")
    t = Replace(t, "�", "N")
    NormalizeText = t
End Function

' =============================================================================
' FIN DEL M�DULO
' =============================================================================

' Notas adicionales:
' - Si cambias los nombres exactos de las hojas (acento, espacios), actualiza los
'   comparadores en GetColsForHoja / GetColTurno / GetColComision.
' - Si la funci�n ObtenerContrasenasSoto no existe, quita la llamada o deja que
'   devuelva cadena vac�a; el c�digo soporta que no haya passStr.
' - Para ajustar el tiempo de cierre autom�tico modifica el valor "4" en:
'       cerrarTimer = Now + (4 / 86400)
'   (ese 4 son segundos; p.ej. 8 => 8 s)
' - El procedimiento p�blico programado por OnTime es "CerrarReservasUnidadFlotante".
'   Debe existir como procedure p�blico en un m�dulo est�ndar y llamar a:
'       frmReservasUnidad.CerrarFormulario
'   (si ya lo ten�as, lo dejamos igual; si no, a�ade esa rutina en un m�dulo est�ndar).

