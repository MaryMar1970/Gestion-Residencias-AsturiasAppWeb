' ==============================================================================

' ThisWorkbook — Eventos de arranque y cierre

' Propósito: Controla el ciclo de vida del BDAS Multiusuario

'

' INSTRUCCIONES:

' 1. Abrir el Editor VBA (ALT+F11)

' 2. Hacer doble clic en "ThisWorkbook" (en el panel izquierdo)

' 3. Pegar este código COMPLETO en la ventana de código

'    (reemplazando cualquier contenido existente)

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

    ' ARRANQUE DEL BDAS MULTIUSUARIO

    ' =====================================================

    ' Flujo:

    '   1. Comprobar conexión a Access

    '   2. Mostrar formulario de login manual

    '   3. Selector de residencia (si tiene varias)

    '   4. Sincronizar datos desde Access a las hojas

    '   5. Activar la hoja de la residencia seleccionada

    ' =====================================================

    

    Application.screenUpdating = False

    

    ' --- Paso 1: Comprobar conexión a la BD Access ---

    If Not modDatabase.ComprobarConexion() Then

        MsgBox "No se puede conectar a la base de datos Access." & vbCrLf & vbCrLf & _

               "Verifica que la unidad H:\ está disponible y que el archivo" & vbCrLf & _

               "H:\ResidenciaBD\Residencia_BE.accdb existe.", _

               vbCritical, "Error de Conexión"

        Application.screenUpdating = True

        ThisWorkbook.Close SaveChanges:=False

        Exit Sub

    End If

    

    ' --- Paso 2 y 3: Login manual + selector de residencia ---

    If Not modDatabase.LoginUsuario() Then

        ' Login fallido o cancelado -> cerrar sin guardar

        Application.screenUpdating = True

        ThisWorkbook.Close SaveChanges:=False

        Exit Sub

    End If

    

    usuarioActual = modDatabase.ObtenerUsuarioActual()

    

    ' --- Paso 4: Sincronizar datos desde Access ---

    Application.StatusBar = "Cargando datos desde la base de datos..."

    modDatabase.SincronizarTodasMisResidencias

    

    ' --- Paso 5: Activar la hoja de la residencia seleccionada ---

    modDatabase.ActivarHojaResidenciaActiva

    

    ' --- Listo ---

    Application.StatusBar = "Conectado como " & modDatabase.ObtenerNombreCompleto() & _

                           " | Residencia: " & modDatabase.ObtenerResidenciaActiva()

    Application.screenUpdating = True

    

    ' Mostrar bienvenida (opcional, se puede quitar)

    MsgBox "Bienvenido/a, " & modDatabase.ObtenerNombreCompleto() & "." & vbCrLf & vbCrLf & _

           "Residencia activa: " & modDatabase.ObtenerResidenciaActiva() & vbCrLf & _

           "Rol: " & modDatabase.ObtenerRolUsuario(), _

           vbInformation, "BDAS - Sesion Iniciada"

End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' =====================================================
    ' CIERRE DEL BDAS MULTIUSUARIO
    ' =====================================================
    ' - Confirmación de salida informativa para el usuario
    ' - Transición suave sin parpadeo (Opción A)
    ' - Registrar logout en el log de Access
    ' - Preparar carátula INICIO para la próxima apertura
    ' - NUNCA guardar cambios en el .xlsm compartido
    '   (los datos están seguros en Access)
    ' =====================================================
    
    ' Confirmación informativa de salida si hay un usuario activo
    If Len(modDatabase.ObtenerUsuarioActual()) > 0 Then
        Dim msgSalida As String
        Dim respSalida As VbMsgBoxResult
        
        msgSalida = "¿Desea cerrar la aplicación BDAS Multiusuario?" & vbCrLf & vbCrLf & _
                    "  - Base de Datos central: Sincronizada y actualizada." & vbCrLf & _
                    "  - Todos los registros y modificaciones han sido guardados." & vbCrLf & vbCrLf & _
                    "Pulse 'Sí' para salir o 'No' para continuar trabajando."
                    
        respSalida = MsgBox(msgSalida, vbQuestion + vbYesNo + vbDefaultButton2, "BDAS Multiusuario - Confirmación de Salida")
        
        If respSalida = vbNo Then
            Cancel = True
            Exit Sub
        End If
        
        ' --- Transición suave de salida (Opción A) ---
        ' 1. Congelar pantalla para evitar parpadeos o saltos de hoja
        Application.ScreenUpdating = False
        
        ' 2. Cambiar cursor a espera
        Application.Cursor = xlWait
        
        ' 3. Mensaje de despedida elegante en la barra de estado
        Application.StatusBar = "Sesión guardada correctamente. Desconectando de forma segura... ¡Hasta pronto!"
        DoEvents
        
        ' 4. Registrar logout en Access
        modDatabase.InsertarLog modDatabase.ObtenerUsuarioActual(), "LOGOUT", "Cierre de sesión normal"
        
        ' 5. Preparar carátula INICIO en memoria para la próxima apertura (sin salto visual)
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
    
    ' Evitar el diálogo "¿Guardar cambios?"
    ' El .xlsm es una plantilla compartida, NO debe guardarse con datos de sesión
    ThisWorkbook.Saved = True
    
    ' Si es el único libro abierto, cerrar la aplicación Excel de forma coordinada
    On Error Resume Next
    If Application.Workbooks.Count <= 1 Then
        Application.Quit
    End If
    On Error GoTo 0
End Sub

' ==============================================================================

' Auto-ocultación de hojas secundarias al abandonarlas

' ==============================================================================

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)

    On Error Resume Next

    Select Case Sh.Name

        Case "RESIDENCIA GIJÓN", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO", _

             "Calendario GIJÓN", "Calendario SOTO", "Calendario OVIEDO", _

             "RESUMEN GIJÓN", "RESUMEN SOTO", "RESUMEN OVIEDO"

            ' Hojas principales operativas: se mantienen siempre visibles

            

        Case Else

            ' Cualquier otra hoja auxiliar (Tarifas, Facturación, Festivos, Listados, etc.)

            ' se oculta automáticamente al salir de ella para mantener la barra limpia

            Sh.visible = xlSheetHidden

    End Select

    On Error GoTo 0

End Sub

