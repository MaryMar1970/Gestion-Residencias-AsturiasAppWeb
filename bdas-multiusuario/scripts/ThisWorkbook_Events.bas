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
    ' - Registrar logout en el log de Access
    ' - NUNCA guardar cambios en el .xlsm compartido
    '   (los datos están seguros en Access)
    ' =====================================================
    
    ' Registrar logout si hubo login exitoso
    If Len(modDatabase.ObtenerUsuarioActual()) > 0 Then
        modDatabase.InsertarLog modDatabase.ObtenerUsuarioActual(), "LOGOUT", _
                               "Cierre de sesión normal"
    End If
    
    ' Evitar el diálogo "¿Guardar cambiosí"
    ' El .xlsm es una plantilla compartida, NO debe guardarse con datos de sesión
    ThisWorkbook.Saved = True
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

