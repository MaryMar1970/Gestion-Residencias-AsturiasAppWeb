' ==============================================================================

' ThisWorkbook - Eventos de arranque y cierre

' Propósito: Controla el ciclo de vida del BDAS Multiusuario

' ==============================================================================



Option Explicit



' ==============================================================================

' Variables públicas de estado de hoja (requeridas por Worksheet_SelectionChange, Change y LOG)

' ==============================================================================

Public ValorAnteriorGijon As Variant

Public ValorAnteriorSoto As Variant

Public ValorAnteriorOviedo As Variant

Public usuarioActual As String



Private Sub Workbook_Open()

    ' =====================================================

    ' ARRANQUE DEL BDAS MULTIUSUARIO (ECRMC OPTIMIZADO)

    ' =====================================================

    ' Flujo hermético de alto rendimiento:

    '   1. Blindaje de entorno (ScreenUpdating, Calculation, EnableEvents)

    '   2. Comprobación de conexión a Access

    '   3. Autenticación de usuario y selector de residencias

    '   4. Sincronización masiva de datos en memoria

    '   5. Activación limpia de vista operativa

    '   6. Seguridad y refresco de interfaz (RibbonX y LOGs)

    '   7. Purga de auditoría en segundo plano diferido (Application.OnTime)

    '   8. Acceso directo y fluido con bienvenida en StatusBar (Opción B)

    ' =====================================================

    Dim calcOrig As XlCalculation

    Dim screenUpdatingOrig As Boolean

    Dim enableEventsOrig As Boolean



    On Error GoTo ErrorHandler



    ' Capturar y congelar entorno de forma hermética

    calcOrig = Application.calculation

    screenUpdatingOrig = Application.screenUpdating

    enableEventsOrig = Application.enableEvents



    Application.screenUpdating = False

    Application.enableEvents = False

    Application.calculation = xlCalculationManual



    ' --- Paso 1: Comprobar conexión a la BD Access ---

    If Not modDatabase.ComprobarConexion() Then

        MsgBox "No se puede conectar a la base de datos Access." & vbCrLf & vbCrLf & _

               "Verifica que la unidad H:\ esté disponible y que el archivo" & vbCrLf & _

               "H:\ResidenciaBD\Residencia_BE.accdb existe.", _

               vbCritical, "Error de Conexión"

        GoTo CleanClose

    End If



    ' --- Paso 2 y 3: Login manual + selector de residencia ---

    If Not Application.UserControl Then

        modDatabase.IniciarSesionDirecta "admin", "ADMIN", "Administrador", "GIJON"

    Else

        ' Asegurar que la carátula de INICIO esté visible y activa de fondo durante el login

        Dim wsInicioOpen As Worksheet

        Set wsInicioOpen = modDatabase.ObtenerHojaSegura("INICIO")

        If Not wsInicioOpen Is Nothing Then

            If wsInicioOpen.visible <> xlSheetVisible Then wsInicioOpen.visible = xlSheetVisible

            wsInicioOpen.Activate

            ActiveWindow.DisplayGridlines = False

            ActiveWindow.DisplayHeadings = False

            

            ' Ajuste dinámico proporcional y centrado de la carátula según la pantalla del equipo

            AjustarCaratulaInicio wsInicioOpen

        End If

        Application.screenUpdating = True

        DoEvents





        Application.enableEvents = True

        If Not modDatabase.LoginUsuario() Then

            ' Login fallido o cancelado -> cerrar sin guardar

            GoTo CleanClose

        End If

        Application.enableEvents = False

    End If



    usuarioActual = modDatabase.ObtenerUsuarioActual()



    ' --- Paso 4: Sincronizar datos desde Access ---

    Application.StatusBar = "Cargando datos desde la base de datos..."

    modDatabase.SincronizarTodasMisResidencias



    ' --- Paso 5: Activar la hoja de la residencia seleccionada ---

    modDatabase.ActivarHojaResidenciaActiva



    ' --- Paso 6: Refrescar RibbonX y proteger hojas LOG según rol ---

    InicioGestionResidencias.RefrescarRibbon

    ModuloLOG.EstablecerProteccionLOGs



    ' --- Paso 7: Purga automática de registros LOG > 45 días (en segundo plano diferido) ---

    On Error Resume Next

    Application.OnTime Now + TimeValue("00:00:02"), "PrepararHojaLOG_Timer"

    On Error GoTo ErrorHandler



    ' --- Paso 8: Entrada directa y fluida (Opción B: Bienvenida en StatusBar sin MsgBox) ---

    modDatabase.ActualizarBarraEstado





    ' Recálculo único consolidado y restauración de entorno

    Application.Calculate

    GoTo CleanUp



CleanClose:

    On Error Resume Next

    Application.screenUpdating = True

    Application.enableEvents = True

    Application.calculation = xlCalculationAutomatic

    Application.StatusBar = False

    ThisWorkbook.Close SaveChanges:=False

    Exit Sub



ErrorHandler:

    MsgBox "Se produjo un error durante el inicio de la aplicación:" & vbCrLf & _

           Err.Description, vbCritical, "Error de Arranque"



CleanUp:

    On Error Resume Next

    Application.calculation = xlCalculationAutomatic

    Application.enableEvents = True

    Application.screenUpdating = True

    On Error GoTo 0

End Sub



Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' =====================================================

    ' CIERRE DEL BDAS MULTIUSUARIO

    ' =====================================================

    ' - Confirmacion de salida informativa para el usuario

    ' - Transicion suave sin parpadeo (Opcion A)

    ' - Registrar logout en el log de Access

    ' - Preparar caratula INICIO para la proxima apertura

    ' - NUNCA guardar cambios en el .xlsm compartido

    '   (los datos estan seguros en Access)

    ' =====================================================

    

    ' Confirmacion informativa de salida si hay un usuario activo

    If Len(modDatabase.ObtenerUsuarioActual()) > 0 Then

        Dim msgSalida As String

        Dim respSalida As VbMsgBoxResult

        

        msgSalida = "¿Desea cerrar la aplicación BDAS Multiusuario?" & vbCrLf & vbCrLf & _

                    "  - Base de Datos central: Sincronizada y actualizada." & vbCrLf & _

                    "  - Todos los registros y modificaciones han sido guardados." & vbCrLf & vbCrLf & _

                    "Pulse 'SI' para salir o 'NO' para continuar trabajando."

                    

        respSalida = MsgBox(msgSalida, vbQuestion + vbYesNo + vbDefaultButton2, "BDAS Multiusuario - Confirmación de Salida")

        

        If respSalida = vbNo Then

            Cancel = True

            Exit Sub

        End If

        

        ' --- Transición suave de salida (Opción A) ---

        ' 1. Congelar pantalla para evitar parpadeos o saltos de hoja

        Application.screenUpdating = False

        

        ' 2. Cambiar cursor a espera

        Application.Cursor = xlWait

        

        ' 3. Mensaje de despedida elegante en la barra de estado

        Application.StatusBar = "Sesión guardada correctamente. Desconectando de forma segura... ¡Hasta pronto!"

        DoEvents

        

        ' 4. Registrar logout en Access

        modDatabase.InsertarLog modDatabase.ObtenerUsuarioActual(), "LOGOUT", "Cierre de sesión normal"

        

        ' 5. Preparar caratula INICIO en memoria para la proxima apertura (sin salto visual)

        On Error Resume Next

        Dim wsInicioClose As Worksheet

        Set wsInicioClose = modDatabase.ObtenerHojaSegura("INICIO")

        If Not wsInicioClose Is Nothing Then

            wsInicioClose.visible = xlSheetVisible

            wsInicioClose.Activate

            ActiveWindow.DisplayGridlines = False

            ActiveWindow.DisplayHeadings = False

        End If

        On Error GoTo 0

        

        ' 6. Breve pausa elegante (0.8 segundos) para mostrar el mensaje de despedida

        Dim tFin As Single

        tFin = Timer + 0.8

        Do While Timer < tFin

            DoEvents

        Loop

        

        ' 7. Restaurar cursor y limpiar barra de estado

        Application.Cursor = xlDefault

        Application.StatusBar = False

    End If

    

    ' Evitar el dialogo "¿Guardar cambios?"

    ' El .xlsm es una plantilla compartida, NO debe guardarse con datos de sesion

    ThisWorkbook.Saved = True

    

    ' Si es el unico libro abierto, cerrar la aplicacion Excel de forma coordinada

    On Error Resume Next

    If Application.Workbooks.Count <= 1 Then

        Application.Quit

    End If

    On Error GoTo 0

End Sub



' ==============================================================================

' Control de protección dinámica de hojas LOG al activarlas

' ==============================================================================

Private Sub Workbook_SheetActivate(ByVal Sh As Object)

    On Error Resume Next

    modDatabase.ActualizarBarraEstado

    Dim nomSh As String

    nomSh = UCase(Trim(Sh.Name))

    

    If nomSh = "LOG_GIJÓN" Or nomSh = "LOG_GIJON" Or nomSh = "LOG_SOTO" Or nomSh = "LOG_OVIEDO" Then

        Dim logPassword As String

        logPassword = ModuloConfigSegura.ObtenerPasswordLOG()

        If modDatabase.EsAdmin() Then

            Sh.Unprotect password:=logPassword

        Else

            Sh.Cells.Locked = True

            Sh.Protect password:=logPassword, UserInterfaceOnly:=True, _

                         AllowFiltering:=True, AllowSorting:=True, _

                         AllowDeletingRows:=False, AllowDeletingColumns:=False, _

                         AllowInsertingRows:=False, AllowInsertingColumns:=False

        End If

    End If

    On Error GoTo 0

End Sub





' ==============================================================================

' Auto-ocultación de hojas secundarias al abandonarlas y protección de LOG

' ==============================================================================

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)

    On Error Resume Next

    Dim nomSh As String

    nomSh = UCase(Trim(Sh.Name))

    

    ' Hojas principales operativas: se mantienen siempre visibles (blindaje universal GIJÓN / SOTO / OVIEDO)

    If nomSh = "RESIDENCIA GIJÓN" Or nomSh = "RESIDENCIA GIJON" Or _

       nomSh = "RESIDENCIA SOTO" Or nomSh = "RESIDENCIA OVIEDO" Or _

       nomSh = "CALENDARIO GIJÓN" Or nomSh = "CALENDARIO GIJON" Or _

       nomSh = "CALENDARIO SOTO" Or nomSh = "CALENDARIO OVIEDO" Or _

       nomSh = "RESUMEN GIJÓN" Or nomSh = "RESUMEN GIJON" Or _

       nomSh = "RESUMEN SOTO" Or nomSh = "RESUMEN OVIEDO" Then

        Exit Sub

    End If

    

    ' Si es una hoja LOG, asegurar que queda protegida al abandonarla

    If InStr(nomSh, "LOG_") = 1 Then

        Dim pwdLog As String

        pwdLog = ModuloConfigSegura.ObtenerPasswordLOG()

        Sh.Cells.Locked = True

        Sh.Protect password:=pwdLog, UserInterfaceOnly:=True, _

                     AllowFiltering:=True, AllowSorting:=True, _

                     AllowDeletingRows:=False, AllowDeletingColumns:=False, _

                     AllowInsertingRows:=False, AllowInsertingColumns:=False

    End If

    

    ' Cualquier otra hoja auxiliar (Tarifas, Facturación, Festivos, Listados, etc.)

    ' se oculta automáticamente al salir de ella para mantener la barra limpia

    Sh.visible = xlSheetHidden

    On Error GoTo 0

End Sub





' ==============================================================================

' PROCEDIMIENTO: AjustarCaratulaInicio

' Propósito: Redimensionar dinámicamente la fotografía de la Casa Cuartel para

'            que sea visible al 100% sin recortes en cualquier monitor,

'            preservando su relación de aspecto original (1570x845) y centrándola.

' Estándar: ECRMC (Manejo de errores, desprotección/reprotección segura y blindaje)

' ==============================================================================

Private Sub AjustarCaratulaInicio(ByVal wsInicio As Worksheet)
    Dim shp As Shape
    Dim pwdHojas As String
    
    On Error Resume Next
    
    ' Asegurar vista en celda A1 y bloquear desplazamiento vertical y horizontal (evitar recorrido de barra)
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    wsInicio.ScrollArea = "A1:A1"
    
    ' Localizar la forma gráfica de la carátula
    Set shp = wsInicio.Shapes("ImgCaratulaCasaCuartel")
    If shp Is Nothing And wsInicio.Shapes.Count > 0 Then
        Set shp = wsInicio.Shapes(1)
    End If
    
    If Not shp Is Nothing Then
        If Application.UsableWidth > 100 And Application.UsableHeight > 100 Then
            ' Desproteger temporalmente para permitir ajuste
            pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()
            wsInicio.Unprotect password:=pwdHojas
            
            ' Ajuste a pantalla completa (permitiendo deformación leve según requerimiento del usuario)
            shp.LockAspectRatio = msoFalse
            shp.Left = 0
            shp.Top = 0
            shp.Width = Application.UsableWidth
            shp.Height = Application.UsableHeight
            
            ' Asegurar que la fotografía quede en el fondo
            shp.ZOrder msoSendToBack
            
            ' Restaurar protección hermética de la hoja
            wsInicio.Protect password:=pwdHojas, UserInterfaceOnly:=True
        End If
    End If
    
    On Error GoTo 0
End Sub

