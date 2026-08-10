VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPanelRGPD 
   Caption         =   "Panel de Control - Archivado RGPD"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   OleObjectBlob   =   "frmPanelRGPD.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmPanelRGPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================================================
' FORMULARIO: frmPanelRGPD
' Prop�sito: Panel de control completo para gesti�n de borrado
' Fecha: 2025-11-19
' �ltima modificaci�n: 2026-02-12
' Autor: Sistema de Gesti�n Residencias Asturias
' VERSI�N: 3 BDAS separadas (GIJ�N, SOTO, OVIEDO)
'=================================================================================

'=================================================================================
' EVENTO: Inicializar formulario
' Se ejecuta autom�ticamente al abrir el formulario
'=================================================================================
Private Sub UserForm_Initialize()
    ' Cargar configuraci�n actual desde CONFIG
    Call CargarConfiguracion
End Sub

'=================================================================================
' EVENTO: Detectar tecla ESC para cerrar
'=================================================================================
Private Sub UserForm_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Si presiona ESC (c�digo 27), cerrar formulario
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

'=================================================================================
' FUNCI�N PRIVADA: Cargar configuraci�n desde hoja CONFIG
'=================================================================================
Private Sub CargarConfiguracion()
    On Error Resume Next
    
    Dim wsConfig As Worksheet
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    
    If wsConfig Is Nothing Then
        ' Si no existe CONFIG, usar valores por defecto
        optAutomatico.Value = True
        txtDiasRetencion.text = "90"
        Exit Sub
    End If
    
        ' Leer modo de borrado (celda B5)
    Dim modoAuto As String
    modoAuto = UCase(Trim(CStr(wsConfig.Range("B5").Value)))
    
    ' Configurar radio buttons seg�n valor
    If modoAuto = "SI" Or modoAuto = "S�" Or modoAuto = "TRUE" Or modoAuto = "1" Or modoAuto = "" Then
        optAutomatico.Value = True
        optManual.Value = False
    Else
        optAutomatico.Value = False
        optManual.Value = True
    End If
    
    ' Leer d�as de retenci�n (celda B4)
    Dim dias As Variant
    dias = wsConfig.Range("B4").Value
    
    If IsNumeric(dias) Then
        txtDiasRetencion.text = CStr(CLng(dias))
    Else
        txtDiasRetencion.text = "90"
    End If
    
    On Error GoTo 0
End Sub

'=================================================================================
' BOT�N: Archivar Ahora
' Ejecuta borrado manual con confirmaci�n previa
' CORREGIDO: No pasa par�metros a ArchivarDatosManual (no los necesita)
'=================================================================================
Private Sub btnArchivarAhora_Click()
    On Error GoTo ErrorHandler
    
    ' Verificar que estamos en una hoja operativa
    Dim hojaActiva As String
    hojaActiva = ""
    
    If Not ActiveSheet Is Nothing Then
        hojaActiva = ActiveSheet.Name
    End If
    
    If hojaActiva <> "RESIDENCIA GIJ�N" And _
       hojaActiva <> "RESIDENCIA SOTO" And _
       hojaActiva <> "RESIDENCIA OVIEDO" Then
        MsgBox "Debe situarse en una de las hojas operativas:" & vbCrLf & vbCrLf & _
               "  - RESIDENCIA GIJ�N" & vbCrLf & _
               "  - RESIDENCIA SOTO" & vbCrLf & _
               "  - RESIDENCIA OVIEDO" & vbCrLf & vbCrLf & _
               "Por favor, sit�ese en una de esas hojas y vuelva a ejecutar el borrado.", _
               vbExclamation, "Advertencia - Hoja no v�lida"
        Exit Sub
    End If
    
    ' Ocultar formulario para mejor visualizaci�n
    Me.Hide
    
    ' ? CORRECCI�N: Llamar SIN par�metros
    Call ModuloArchivadoDatosBDAS.ArchivarDatosManual
    
    ' Volver a mostrar formulario
    Me.Show
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al ejecutar borrado:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & _
           "N�mero: " & Err.Number, _
           vbCritical, "Error en borrado"
    Me.Show
End Sub

'=================================================================================
' FUNCI�N AUXILIAR: Ejecutar Borrado manual con confirmaci�n
'=================================================================================
Private Sub EjecutarArchivadoManualConConfirmacion()
    On Error GoTo ErrorHandler
    
    ' PASO 1: Contar cu�ntos registros se van a archivar
    Dim totalRegistros As Long
    totalRegistros = ModuloArchivadoDatosBDAS.ContarRegistrosParaArchivar()
    
    If totalRegistros = 0 Then
        Dim msg1 As String
        msg1 = "No hay registros antiguos para archivar." & vbCrLf & vbCrLf
        msg1 = msg1 & "Todos los registros tienen menos de " & txtDiasRetencion.text & " d�as" & vbCrLf
        msg1 = msg1 & "desde la fecha de salida (finalizaci�n del servicio)." & vbCrLf & vbCrLf
        msg1 = msg1 & "Los datos archivados se conservar�n en BDAS durante 3 a�os" & vbCrLf
        msg1 = msg1 & "seg�n Real Decreto 933/2021, de 26 de octubre."
        
        MsgBox msg1, vbInformation, "Borrado - Sin Acci�n Necesaria"
        Exit Sub
    End If
    
    ' PASO 2: Mostrar resumen y pedir confirmaci�n
    Dim msg2 As String
    Dim msg3 As String
    Dim msg4 As String
    Dim msgCompleto As String
    
    msg2 = "SE VAN A BORRAR " & totalRegistros & " REGISTROS" & vbCrLf & vbCrLf
    msg2 = msg2 & "Registros con m�s de " & txtDiasRetencion.text & " d�as desde fecha de salida." & vbCrLf & vbCrLf
'    msg2 = msg2 & "Cada residencia se archivar� en su BDAS:" & vbCrLf
'    msg2 = msg2 & "  - RESIDENCIA GIJ�N -> BDAS GIJ�N" & vbCrLf
'    msg2 = msg2 & "  - RESIDENCIA SOTO -> BDAS SOTO" & vbCrLf
'    msg2 = msg2 & "  - RESIDENCIA OVIEDO -> BDAS OVIEDO" & vbCrLf & vbCrLf
    
'    msg3 = "DATOS QUE SE ELIMINAR�N de hojas operativas:" & vbCrLf
'    msg3 = msg3 & "  DNI/NIF, Nombre, Tel�fono, Direcci�n, Grabaci�n" & vbCrLf & vbCrLf
'    msg3 = msg3 & "DATOS QUE SE CONSERVAN:" & vbCrLf
'    msg3 = msg3 & "  N� Orden, Fechas estancia, Importes, Estado pago" & vbCrLf & vbCrLf
'    msg3 = msg3 & "Copia completa guardada en BDAS correspondiente" & vbCrLf
'    msg3 = msg3 & "(conservaci�n: 3 a�os seg�n RD 933/2021)" & vbCrLf & vbCrLf
    
'    msg4 = "BASE LEGAL:" & vbCrLf
'    msg4 = msg4 & "Real Decreto 933/2021, de 26 de octubre:" & vbCrLf
'    msg4 = msg4 & "  Los datos del registro inform�tico deber�n conservarse" & vbCrLf
'    msg4 = msg4 & "  durante un plazo de tres a�os a contar desde la" & vbCrLf
'    msg4 = msg4 & "  finalizaci�n del servicio o prestaci�n contratada." & vbCrLf & vbCrLf
'    msg4 = msg4 & "RGPD (UE 2016/679) - Arts. 5.1. c y 5.1.e" & vbCrLf & vbCrLf
    msg4 = msg4 & "Continuar con el borrado?"
    
    msgCompleto = msg2 & msg3 & msg4
    
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox(msgCompleto, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Borrado")
    
    If respuesta = vbNo Then
        MsgBox "Operaci�n cancelada por el usuario.", vbInformation, "Borrado"
        Exit Sub
    End If
    
    ' PASO 3: Ejecutar borrado manual
    Call ModuloArchivadoDatosBDAS.ArchivarDatosManual
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en borrado manual: " & Err.Description, vbCritical
End Sub

'=================================================================================
' BOT�N: Ver BDAS
' Versi�n mejorada: Detecta qu� BDAS existen y adapta el selector
'=================================================================================
'Private Sub btnVerBDAS_Click()
'    On Error GoTo ErrorHandler
    
'    Dim password As String
'    Dim seleccion As String
'    Dim nombreBDAS As String
    
    ' Solicitar contrase�a
'    password = InputBox("Acceso a datos archivados (RGPD)" & vbCrLf & vbCrLf & _
'                       "Esta hoja contiene datos personales sensibles." & vbCrLf & _
'                       "Introduzca la contrase�a maestra:", _
'                       "Acceso BDAS - Datos Sensibles")
    
    ' Si cancela, salir
'    If password = "" Then Exit Sub
    
    ' Validar contrase�a
'    If password <> ModuloConfigSegura.ObtenerClaveMaestra() Then
'        MsgBox "Contrase�a incorrecta.", vbCritical, "Acceso Denegado"
'        Exit Sub
'    End If
    
    ' Detectar qu� hojas BDAS existen
'    Dim bdasDisponibles As Collection
'    Set bdasDisponibles = DetectarBDASDisponibles()
    
    ' Si no existe ninguna BDAS
'    If bdasDisponibles.Count = 0 Then
'        MsgBox "No existen hojas BDAS." & vbCrLf & vbCrLf & _
'               "No hay datos archivados a�n." & vbCrLf & _
'               "Las hojas BDAS se crear�n autom�ticamente cuando se archive por primera vez.", _
'               vbInformation, "Sin BDAS disponibles"
'        Exit Sub
'    End If
    
    ' Si solo existe UNA BDAS, mostrarla directamente
'    If bdasDisponibles.Count = 1 Then
'        nombreBDAS = bdasDisponibles(1)
        
'        MsgBox "Se mostrar� la hoja:" & vbCrLf & vbCrLf & _
'               nombreBDAS, vbInformation, "BDAS Disponible"
'    Else
        ' Si existen VARIAS BDAS, mostrar selector din�mico
'        Dim mensajeSelector As String
'        Dim i As Long
        
'        mensajeSelector = "Seleccione qu� BDAS desea visualizar:" & vbCrLf & vbCrLf
'        mensajeSelector = mensajeSelector & "Escriba una de las siguientes opciones:" & vbCrLf
        
'        For i = 1 To bdasDisponibles.Count
'            mensajeSelector = mensajeSelector & "  " & i & " - " & bdasDisponibles(i) & vbCrLf
'        Next i
        
'        seleccion = InputBox(mensajeSelector, "Seleccionar BDAS", "1")
        
'        If seleccion = "" Then Exit Sub
        
        ' Validar selecci�n
'        Dim numSeleccion As Long
'        If Not IsNumeric(seleccion) Then
'            MsgBox "Opci�n no v�lida.  Debe escribir un n�mero.", vbExclamation, "Selecci�n Inv�lida"
'            Exit Sub
'        End If
        
'        numSeleccion = CLng(seleccion)
        
'        If numSeleccion < 1 Or numSeleccion > bdasDisponibles.Count Then
'            MsgBox "Opci�n fuera de rango. Debe escribir un n�mero entre 1 y " & bdasDisponibles.Count, _
'                   vbExclamation, "Selecci�n Inv�lida"
'            Exit Sub
'        End If
        
        ' Obtener nombre de la BDAS seleccionada
'        nombreBDAS = bdasDisponibles(numSeleccion)
'    End If
    
    ' Verificar que existe la BDAS (doble verificaci�n)
'    Dim wsBDAS As Worksheet
'    On Error Resume Next
'    Set wsBDAS = ThisWorkbook.Worksheets(nombreBDAS)
'    On Error GoTo ErrorHandler
    
'    If wsBDAS Is Nothing Then
'        MsgBox "Error: La hoja " & nombreBDAS & " no pudo ser accedida.", vbCritical, "Error"
'        Exit Sub
'    End If
    
    ' Mostrar mensaje de advertencia
'    MsgBox "DATOS SENSIBLES - RGPD" & vbCrLf & vbCrLf & _
'           "Va a acceder a datos personales archivados de:" & vbCrLf & _
'           nombreBDAS & vbCrLf & vbCrLf & _
'           "Este acceso ser� registrado en el LOG." & vbCrLf & vbCrLf & _
'           "La hoja se ocultar� autom�ticamente al cambiar de pesta�a.", _
'           vbExclamation, "Acceso BDAS"
    
    ' Ocultar formulario y hacer visible BDAS seleccionada
'    Unload Me
'    wsBDAS.visible = xlSheetVisible
'    wsBDAS.Activate
    
    ' Registrar acceso en LOG
'    Call RegistrarAccesoBDASDesdePanel(nombreBDAS)
    
'    Exit Sub
    
'ErrorHandler:
'    MsgBox "Error al acceder a BDAS: " & Err.Description, vbCritical, "Error"
'End Sub

'=================================================================================
' FUNCI�N AUXILIAR: Detectar qu� hojas BDAS existen realmente
' Retorna: Collection con los nombres de las BDAS que existen
'=================================================================================
'Private Function DetectarBDASDisponibles() As Collection
'    Dim disponibles As Collection
'    Set disponibles = New Collection
    
'    Dim ws As Worksheet
'    Dim nombresBDAS As Variant
'    Dim i As Long
    
    ' Lista de nombres posibles de BDAS
'    nombresBDAS = Array("BDAS GIJ�N", "BDAS SOTO", "BDAS OVIEDO")
    
    ' Verificar cu�les existen
'    For i = LBound(nombresBDAS) To UBound(nombresBDAS)
'        On Error Resume Next
'        Set ws = Nothing
'        Set ws = ThisWorkbook.Worksheets(CStr(nombresBDAS(i)))
'        On Error GoTo 0
        
'        If Not ws Is Nothing Then
'            disponibles.Add CStr(nombresBDAS(i))
'        End If
'    Next i
    
'    Set DetectarBDASDisponibles = disponibles
'End Function

'=================================================================================
' FUNCI�N AUXILIAR: Registrar acceso desde Panel
' ACTUALIZADO: Incluye nombre de BDAS accedida
' CORREGIDO: Ya NO usa Application.UserName (usuario de Windows)
'=================================================================================
'Private Sub RegistrarAccesoBDASDesdePanel(nombreBDAS As String)
'    On Error GoTo ErrorHandler
    
'    Dim wsLOG As Worksheet
'    Dim ultimaFila As Long
'    Dim logPassword As String
'    Dim logSheets As Variant
'    Dim i As Long
'    Dim usuarioActual As String
    
    ' ? CORRECCI�N: Obtener usuario SOLO de la sesi�n activa
'    usuarioActual = Trim(ThisWorkbook.usuarioActual)
    
    ' ? VALIDACI�N: Si est� vac�o, no deber�a estar aqu�
'    If usuarioActual = "" Then
        ' Esto NO deber�a ocurrir nunca si el login funciona bien
'        MsgBox "Error cr�tico: No se puede identificar el usuario de sesi�n." & vbCrLf & _
'               "Por favor, cierre el libro y vuelva a iniciar sesi�n.", vbCritical, "Error de Sesi�n"
'        Debug.Print "ERROR CR�TICO: usuarioActual vac�o en RegistrarAccesoBDASDesdePanel - " & Now
'        Exit Sub
'    End If
    
'    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
'    logSheets = Array("LOG_GIJ�N", "LOG_SOTO", "LOG_OVIEDO")
    
    ' Registrar en todas las hojas LOG
'    For i = LBound(logSheets) To UBound(logSheets)
'        On Error Resume Next
'        Set wsLOG = ThisWorkbook.Worksheets(logSheets(i))
        
'        If Not wsLOG Is Nothing Then
'            wsLOG.Unprotect password:=logPassword
'            ultimaFila = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
'            wsLOG.Cells(ultimaFila, 1).Value = usuarioActual   ' ? Ahora usa el usuario correcto
'            wsLOG.Cells(ultimaFila, 2).Value = Now
'            wsLOG.Cells(ultimaFila, 3).Value = ""
'            wsLOG.Cells(ultimaFila, 4).Value = "ACCESO " & nombreBDAS
'            wsLOG.Cells(ultimaFila, 5).Value = ""
'            wsLOG.Cells(ultimaFila, 6).Value = ""
'            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
'        End If
'        Set wsLOG = Nothing
'        On Error GoTo ErrorHandler
'    Next i
    
'    Exit Sub
    
'ErrorHandler:
'    Debug.Print "Error en RegistrarAccesoBDASDesdePanel: " & Err.Number & " - " & Err.Description & " (" & Now & ")"
'    Resume Next
'End Sub

'=================================================================================
' BOT�N: Guardar Configuraci�n
' CORREGIDO: Usa B4 (d�as), B5 (modo), cierra formulario al terminar
'=================================================================================
Private Sub btnGuardar_Click()
    On Error GoTo ErrorHandler
    
    ' VALIDAR d�as de retenci�n
    If Not IsNumeric(txtDiasRetencion.text) Then
        MsgBox "El valor de 'D�as de retenci�n' debe ser un n�mero." & vbCrLf & vbCrLf & _
               "Por favor, introduce un valor num�rico.", _
               vbExclamation, "Valor Inv�lido"
        txtDiasRetencion.SetFocus
        Exit Sub
    End If
    
    Dim dias As Long
    dias = CLng(txtDiasRetencion.text)
    
    ' Validar rango
    If dias < 30 Or dias > 365 Then
        Dim msgValidacion As String
        msgValidacion = "Los d�as de retenci�n deben estar entre 30 y 365." & vbCrLf & vbCrLf
        msgValidacion = msgValidacion & "Valor actual: " & dias & vbCrLf
        msgValidacion = msgValidacion & "Rango permitido: 30 - 365 d�as" ' & vbCrLf & vbCrLf
'        msgValidacion = msgValidacion & "Nota: Los datos archivados se conservan en BDAS durante" & vbCrLf
'        msgValidacion = msgValidacion & "3 a�os seg�n Real Decreto 933/2021, de 26 de octubre."
        
        MsgBox msgValidacion, vbExclamation, "Valor Fuera de Rango"
        txtDiasRetencion.SetFocus
        Exit Sub
    End If
    
    ' ? Verificar si los d�as han cambiado
    Dim wsConfig As Worksheet
    Dim diasAnteriores As Long
    Dim cambianDias As Boolean
    
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    diasAnteriores = CLng(wsConfig.Range("B4").Value)  ' ? B4 no B51
    cambianDias = (dias <> diasAnteriores)
    
    ' ? PASO 1: PEDIR CONFIRMACI�N (con advertencias si cambian d�as)
    Dim modoTexto As String
    If optAutomatico.Value Then
        modoTexto = "AUTOM�TICO al abrir"
    Else
        modoTexto = "SOLO MANUAL"
    End If
    
    Dim respuesta As VbMsgBoxResult
    Dim mensajeConfirmacion As String
    
    mensajeConfirmacion = "�Guardar cambios en la configuraci�n?" & vbCrLf & vbCrLf & _
                          "Modo: " & modoTexto & vbCrLf & _
                          "D�as de retenci�n: " & dias & " d�as"
    
    If cambianDias Then
        mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & _
                              "ATENCI�N - CAMBIO DE CONFIGURACI�N CR�TICA:" & vbCrLf & _
                              "  Valor actual: " & diasAnteriores & " d�as" & vbCrLf & _
                              "  Valor nuevo: " & dias & " d�as" & vbCrLf & vbCrLf & _
                              "Este cambio afectar� al borrado autom�tico RGPD" & vbCrLf & _
                              "y ser� registrado en el LOG de auditor�a."
    Else
        mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & _
                              "Los cambios se aplicar�n en la pr�xima sesi�n."
    End If
    
    respuesta = MsgBox(mensajeConfirmacion, vbQuestion + vbYesNo, "Confirmar Guardado")
    
    If respuesta = vbNo Then
        MsgBox "Cambios descartados.", vbInformation, "Configuraci�n"
        Exit Sub
    End If
    
    ' ? PASO 2: Si cambian d�as, AHORA solicitar contrase�a como VALIDACI�N FINAL
    If cambianDias Then
        Dim usuarioActual As String
        Dim passwordIngresada As String
        Dim validacionExitosa As Boolean
        
        usuarioActual = Trim(ThisWorkbook.usuarioActual)
        If usuarioActual = "" Then
            MsgBox "Error: No se puede identificar el usuario de sesi�n.", vbCritical
            Exit Sub
        End If
        
        ' Mostrar di�logo de validaci�n FINAL
        passwordIngresada = InputBox( _
            "VALIDACI�N FINAL DE SEGURIDAD" & vbCrLf & vbCrLf & _
            "Para aplicar el cambio de d�as de retenci�n:" & vbCrLf & _
            "  " & diasAnteriores & " d�as  a  " & dias & " d�as" & vbCrLf & vbCrLf & _
            "Usuario: " & usuarioActual & vbCrLf & vbCrLf & _
            "Introduzca su contrase�a para confirmar:", _
            "Validaci�n de Identidad")
        
        If passwordIngresada = "" Then
            MsgBox "Operaci�n cancelada." & vbCrLf & vbCrLf & _
                   "No se han guardado cambios.", vbInformation, "Cancelado"
            Exit Sub
        End If
        
        ' ? VALIDAR CONTRASE�A
        validacionExitosa = ValidarPasswordUsuarioActual(usuarioActual, passwordIngresada)
        
        If Not validacionExitosa Then
            MsgBox "Contrase�a incorrecta." & vbCrLf & vbCrLf & _
                   "El cambio NO se ha realizado." & vbCrLf & _
                   "La configuraci�n no ha sido modificada.", _
                   vbCritical, "Acceso Denegado"
            Exit Sub
        End If
    End If
    
    ' ? PASO 3: GUARDAR en CONFIG (solo si pas� todas las validaciones)
    Call GuardarConfiguracion
    
    ' ? PASO 4: Si cambiaron los d�as, registrar en LOG
    If cambianDias Then
        Call RegistrarCambioDiasRetencion(ThisWorkbook.usuarioActual, diasAnteriores, dias)
    End If
    
    ' ? PASO 5: Mensaje de �xito
    Dim mensajeExito As String
    mensajeExito = "Configuraci�n guardada correctamente." & vbCrLf & vbCrLf & _
                   "Modo: " & modoTexto & vbCrLf & _
                   "D�as de retenci�n: " & dias & " d�as"
    
    If cambianDias Then
        mensajeExito = mensajeExito & vbCrLf & vbCrLf & _
                       "CAMBIO REGISTRADO EN LOG:" & vbCrLf & _
                       "  Usuario: " & ThisWorkbook.usuarioActual & vbCrLf & _
                       "  Cambio: " & diasAnteriores & " a " & dias & " d�as" & vbCrLf & _
                       "  Fecha: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & vbCrLf & _
                       "El cambio se aplicar� al reiniciar Excel."
    Else
        mensajeExito = mensajeExito & vbCrLf & vbCrLf & _
                       "Los cambios se aplicar�n al reiniciar Excel."
    End If
    
    MsgBox mensajeExito, vbInformation, "Configuraci�n Guardada"
    
    ' ? PASO 6: CERRAR EL FORMULARIO
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al guardar configuraci�n:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
End Sub

'=================================================================================
' FUNCI�N PRIVADA: Guardar configuraci�n en hoja CONFIG
' CORREGIDO: B4 (d�as), B5 (modo autom�tico)
'=================================================================================
Private Sub GuardarConfiguracion()
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim passwordConfig As String
    Dim estabaProtegida As Boolean
    Dim diasGuardar As Long
    
    ' ? 1. Obtener hoja CONFIG
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    
    If wsConfig Is Nothing Then
        MsgBox "ERROR: No existe la hoja CONFIG.", vbCritical
        Exit Sub
    End If
    
    ' ? 2. Obtener contrase�a y desproteger si es necesario
    passwordConfig = ModuloConfigSegura.ObtenerPasswordConfig()
    estabaProtegida = wsConfig.ProtectContents
    
    If estabaProtegida Then
        wsConfig.Unprotect password:=passwordConfig
    End If
    
    ' ? 3.  GUARDAR MODO AUTOM�TICO en B5 (no B50)
    If optAutomatico.Value Then
        wsConfig.Range("B5").Value = "SI"
    Else
        wsConfig.Range("B5").Value = "NO"
    End If
    
    ' ? 4.  GUARDAR D�AS DE RETENCI�N en B4 (no B51)
    diasGuardar = CLng(txtDiasRetencion.text)
    wsConfig.Range("B4").Value = diasGuardar
    
    ' ? 5.  Agregar etiquetas si no existen
    If wsConfig.Range("A4").Value = "" Then
        wsConfig.Range("A4").Value = "D�as de retenci�n RGPD:"
    End If
    If wsConfig.Range("A5").Value = "" Then
        wsConfig.Range("A5").Value = "Borrado autom�tico RGPD:"
    End If
    
    ' ? 6. Reproteger si estaba protegida
    If estabaProtegida Then
        wsConfig.Protect password:=passwordConfig, UserInterfaceOnly:=True
    End If
    
    ' ? 7.  DEBUG
    Debug.Print "CONFIG guardado - B4=" & wsConfig.Range("B4").Value & " B5=" & wsConfig.Range("B5").Value
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al guardar configuraci�n:" & vbCrLf & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    If estabaProtegida And Not wsConfig Is Nothing Then
        wsConfig.Protect password:=passwordConfig, UserInterfaceOnly:=True
    End If
End Sub

'=================================================================================
' FUNCI�N: Validar contrase�a del usuario actual
'=================================================================================
Private Function ValidarPasswordUsuarioActual(usuario As String, password As String) As Boolean
    On Error Resume Next
    
    Dim wsUsuarios As Worksheet
    Dim i As Long
    Dim saltUsuario As String
    Dim hashUsuario As String
    
    ValidarPasswordUsuarioActual = False
    
    Set wsUsuarios = ThisWorkbook.Sheets("USUARIOS")
    If wsUsuarios Is Nothing Then Exit Function
    
    For i = 2 To wsUsuarios.Cells(wsUsuarios.Rows.Count, "A").End(xlUp).Row
        If Trim(wsUsuarios.Cells(i, 1).Value) = usuario Then
            saltUsuario = Trim(wsUsuarios.Cells(i, 3).Value)
            hashUsuario = Trim(wsUsuarios.Cells(i, 4).Value)
            
            If ModuloConfigSegura.ValidarPasswordHash(password, saltUsuario, hashUsuario) Then
                ValidarPasswordUsuarioActual = True
                Exit Function
            End If
        End If
    Next i
    
    On Error GoTo 0
End Function

'=================================================================================
' FUNCI�N: Registrar cambio en LOGs
'=================================================================================
Private Sub RegistrarCambioDiasRetencion(usuario As String, valorAnterior As Long, valorNuevo As Long)
    On Error Resume Next
    
    Dim wsLOG As Worksheet
    Dim ultimaFila As Long
    Dim logPassword As String
    Dim logSheets As Variant
    Dim i As Long
    Dim screenStatus As Boolean
    
    screenStatus = Application.screenUpdating
    Application.screenUpdating = False
    
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    logSheets = Array("LOG_GIJ�N", "LOG_SOTO", "LOG_OVIEDO")
    
    For i = LBound(logSheets) To UBound(logSheets)
        Set wsLOG = ThisWorkbook.Worksheets(logSheets(i))
        
        If Not wsLOG Is Nothing Then
            wsLOG.Unprotect password:=logPassword
            ultimaFila = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            
            With wsLOG
                .Cells(ultimaFila, 1).Value = usuario
                .Cells(ultimaFila, 2).Value = Now
                .Cells(ultimaFila, 3).Value = "CONFIG"
                .Cells(ultimaFila, 4).Value = "RETENCI�N DATOS"
                .Cells(ultimaFila, 5).Value = CStr(valorAnterior) & " d�as"
                .Cells(ultimaFila, 6).Value = CStr(valorNuevo) & " d�as"
            End With
            
            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
        End If
        Set wsLOG = Nothing
    Next i
    
    Application.screenUpdating = screenStatus
    
    On Error GoTo 0
End Sub

'=================================================================================
' BOT�N: Cerrar
' Cierra el formulario sin guardar cambios
'=================================================================================
Private Sub btnCerrar_Click()
    Unload Me
End Sub

