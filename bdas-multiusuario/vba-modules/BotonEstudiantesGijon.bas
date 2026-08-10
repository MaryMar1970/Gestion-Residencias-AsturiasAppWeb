Attribute VB_Name = "BotonEstudiantesGijon"
Option Explicit

'==============================================================================
' M�DULO OPTIMIZADO: ModuloNavegacionGijon
'
' Ventajas sobre c�digo original:
' 1. Manejo de errores integrado (no crashea si hoja no existe)
' 2. C�digo m�s mantenible (l�gica centralizada)
' 3. Consistencia en comportamiento
' 4. F�cil agregar nuevas hojas
'==============================================================================

'---------------------------------------------------------------
' MOSTRAR FACTURA ESTUDIANTES - GIJ�N
' Llamado desde bot�n en la cinta de opciones
'---------------------------------------------------------------
Sub MostrarFacturaEstudianteGijon(control As IRibbonControl)
    MostrarHojaSegura "FACTURA R. ESTUDIANTES"
End Sub

'---------------------------------------------------------------
' MOSTRAR RESIDENCIA ESTUDIANTES - GIJ�N
' Llamado desde bot�n en la cinta de opciones
'---------------------------------------------------------------
Sub ResidenciaEstudiantesGijon(control As IRibbonControl)
    MostrarHojaSegura "RESIDENCIA ESTUDIANTES"
End Sub

'---------------------------------------------------------------
' MOSTRAR LIQUIDACI�N ESTUDIANTES - GIJ�N
' Llamado desde bot�n en la cinta de opciones
'---------------------------------------------------------------
Sub LiquidacionEstudiantesGijon(control As IRibbonControl)
    MostrarHojaSegura "LIQUIDACION R. ESTUDIANTES"
End Sub

'---------------------------------------------------------------
' PROCEDIMIENTO PRIVADO REUTILIZABLE
' (Manejador central para mostrar cualquier hoja)
'
' PAR�METROS:
'   nombreHoja: Nombre exacto de la hoja a mostrar
'
' CARACTER�STICAS:
' - Manejo de errores silencioso (no muestra mensajes al usuario)
' - Log de errores en ventana Inmediato (Ctrl+G) para depuraci�n
' - Validaci�n b�sica de par�metros
'---------------------------------------------------------------
Private Sub MostrarHojaSegura(ByVal nombreHoja As String)
    
    ' VALIDACI�N: Si nombre est� vac�o, salir silenciosamente
    If Trim(nombreHoja) = "" Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' 1. OBTENER REFERENCIA A LA HOJA
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    
    ' 2. HACER VISIBLE (por si est� oculta)
    ws.visible = xlSheetVisible
    
    ' 3. ACTIVAR (traer al frente)
    ws.Activate
    
    ' Salida exitosa
    Exit Sub

ErrorHandler:
    ' MANEJO DE ERRORES SILENCIOSO:
    ' - No muestra mensajes al usuario (evita interrupciones)
    ' - Registra error para depuraci�n (ventana Inmediato, Ctrl+G)
    Debug.Print "[" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "] " & _
                "Error mostrando hoja '" & nombreHoja & "': " & _
                Err.Description & " (Error #" & Err.Number & ")"
    
    ' Opcional: Si quieres mostrar mensaje al usuario, descomenta:
    ' MsgBox "No se pudo mostrar la hoja: " & nombreHoja, vbInformation, "Aviso"
    
End Sub

