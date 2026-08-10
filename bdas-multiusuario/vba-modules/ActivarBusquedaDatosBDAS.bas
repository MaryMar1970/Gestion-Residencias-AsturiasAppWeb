Attribute VB_Name = "ActivarBusquedaDatosBDAS"
'Attribute VB_Name = "ActivarBusquedaDatosBDAS"

' =============================================================================
' M�DULO: ActivarBusquedaDatosBDAS
' PROP�SITO: Interfaz para activar/desactivar b�squeda de datos en hojas BDAS
' =============================================================================
'
' CONFIGURACI�N EN HOJA CONFIG:
' -----------------------------------------------------------------------------
' | Celda | Prop�sito                          | Valores    | Este m�dulo |
' |-------|------------------------------------|-----------:|:-----------:|
' | B5    | Archivado autom�tico RGPD          | SI/NO      | NO usa      |
' | B6    | Activar b�squeda en hojas BDAS     | SI/NO      | S� usa      |
' -----------------------------------------------------------------------------
'
' NOTA: B5 controla el archivado autom�tico (ModuloArchivadoDatos, frmPanelRGPD)
'       B6 controla si al introducir un DNI se busca tambi�n en hojas BDAS
'
' RELACI�N CON OTROS M�DULOS:
'   - Este m�dulo CONFIGURA el valor de B6
'   - BusquedaDNIResidencias LEE el valor de B6 para decidir si buscar en BDAS
'
' =============================================================================
' =============================================================================
' M�DULO: ActivarBusquedaDatosBDAS
' PROP�SITO: Permitir activar/desactivar la b�squeda de datos en hojas BDAS
'            al introducir un DNI en las hojas RESIDENCIA
' =============================================================================
' =============================================================================
' PROCEDIMIENTO: ConfigurarBusquedaBDAS
' PROP�SITO: Macro principal para activar/desactivar b�squeda en BDAS
' USO: Ejecutar manualmente desde Macros (Alt+F8)
' =============================================================================
Public Sub ConfigurarBusquedaBDAS()
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim estadoActual As String
    Dim respuesta As VbMsgBoxResult
    Dim mensaje As String
    Dim estabaProtegida As Boolean
    Dim passwordConfig As String
    
    ' Verificar que existe la hoja CONFIG
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    On Error GoTo ErrorHandler
    
    If wsConfig Is Nothing Then
        MsgBox "Error: No se encuentra la hoja CONFIG en el libro." & vbCrLf & vbCrLf & _
               "La configuraci�n no puede realizarse.", vbCritical, "Error de Configuraci�n"
        Exit Sub
    End If
    
    ' Verificar si la hoja est� protegida
    estabaProtegida = wsConfig.ProtectContents
    
    ' Si est� protegida, intentar desprotegerla (sin password o con password vac�o)
    If estabaProtegida Then
        On Error Resume Next
        wsConfig.Unprotect password:=""
        If Err.Number <> 0 Then
            ' La hoja tiene contrase�a, solicitar al usuario
            passwordConfig = InputBox("La hoja CONFIG est� protegida con contrase�a." & vbCrLf & vbCrLf & _
                                     "Ingrese la contrase�a para continuar:", _
                                     "Contrase�a requerida")
            If passwordConfig = "" Then
                MsgBox "Operaci�n cancelada.  No se ingres� contrase�a.", vbExclamation, "Cancelado"
                Exit Sub
            End If
            
            Err.Clear
            wsConfig.Unprotect password:=passwordConfig
            If Err.Number <> 0 Then
                MsgBox "Contrase�a incorrecta. No se puede modificar la configuraci�n.", vbCritical, "Error"
                Exit Sub
            End If
        End If
        On Error GoTo ErrorHandler
    End If
    
    ' Leer estado actual
    estadoActual = UCase(Trim(wsConfig.Range("B6").Value))
    
    ' Construir mensaje informativo
    mensaje = "-------------------------------------------" & vbCrLf & _
              "CONFIGURACI�N DE B�SQUEDA EN BDAS" & vbCrLf & _
              "-------------------------------------------" & vbCrLf & vbCrLf & _
              "Estado actual: " & IIf(estadoActual = "S�" Or estadoActual = "SI", "ACTIVADO ?", "DESACTIVADO ?") & vbCrLf & vbCrLf & _
              "�Qu� hace esta configuraci�n?" & vbCrLf & _
              "--------------------------------" & vbCrLf & _
              "Al introducir un DNI en la columna I de las hojas" & vbCrLf & _
              "RESIDENCIA (GIJ�N, SOTO u OVIEDO):" & vbCrLf & vbCrLf & _
              "� SI est� ACTIVADO:" & vbCrLf & _
              "  Buscar� el DNI primero en la misma hoja," & vbCrLf & _
              "  y si no lo encuentra, buscar� en las hojas BDAS" & vbCrLf & _
              "  para autocompletar los datos." & vbCrLf & vbCrLf & _
              "� SI est� DESACTIVADO:" & vbCrLf & _
              "  Solo buscar� el DNI en la misma hoja," & vbCrLf & _
              "  sin consultar las hojas BDAS." & vbCrLf & vbCrLf & _
              "-------------------------------------------" & vbCrLf & vbCrLf & _
              "�Desea ACTIVAR la b�squeda en hojas BDAS?" & vbCrLf & vbCrLf & _
              "Presione S� para ACTIVAR" & vbCrLf & _
              "Presione NO para DESACTIVAR"
    
    ' Mostrar cuadro de di�logo
    respuesta = MsgBox(mensaje, vbYesNo + vbQuestion, "Configurar B�squeda en BDAS")
    
    ' Actualizar configuraci�n seg�n respuesta
    If respuesta = vbYes Then
        wsConfig.Range("B6").Value = "S�"
        
        ' Re-proteger si estaba protegida
        If estabaProtegida Then
            On Error Resume Next
            wsConfig.Protect password:=passwordConfig, DrawingObjects:=True, Contents:=True, Scenarios:=True
            On Error GoTo ErrorHandler
        End If
        
        MsgBox "B�squeda en hojas BDAS ACTIVADA" & vbCrLf & vbCrLf & _
               "Al introducir un DNI en las hojas RESIDENCIA," & vbCrLf & _
               "se buscar� autom�ticamente en las hojas BDAS" & vbCrLf & _
               "si no se encuentra en la misma hoja.", _
               vbInformation, "Configuraci�n Guardada"
    Else
        wsConfig.Range("B6").Value = "NO"
        
        ' Re-proteger si estaba protegida
        If estabaProtegida Then
            On Error Resume Next
            wsConfig.Protect password:=passwordConfig, DrawingObjects:=True, Contents:=True, Scenarios:=True
            On Error GoTo ErrorHandler
        End If
        
        MsgBox "B�squeda en hojas BDAS DESACTIVADA" & vbCrLf & vbCrLf & _
               "Al introducir un DNI en las hojas RESIDENCIA," & vbCrLf & _
               "solo se buscar� en la misma hoja," & vbCrLf & _
               "sin consultar las hojas BDAS.", _
               vbInformation, "Configuraci�n Guardada"
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Re-proteger la hoja si hubo error
    If estabaProtegida And Not wsConfig Is Nothing Then
        On Error Resume Next
        wsConfig.Protect password:=passwordConfig, DrawingObjects:=True, Contents:=True, Scenarios:=True
        On Error GoTo 0
    End If
    
    MsgBox "Error al configurar b�squeda en BDAS:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
End Sub

' =============================================================================
' PROCEDIMIENTO: VerEstadoBusquedaBDAS
' PROP�SITO: Ver el estado actual de la configuraci�n sin modificarlo
' USO: Ejecutar manualmente desde Macros (Alt+F8)
' =============================================================================
Public Sub VerEstadoBusquedaBDAS()
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim estadoActual As String
    Dim mensaje As String
    
    ' Verificar que existe la hoja CONFIG
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    On Error GoTo ErrorHandler
    
    If wsConfig Is Nothing Then
        MsgBox "Error: No se encuentra la hoja CONFIG.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Leer estado actual
    estadoActual = UCase(Trim(wsConfig.Range("B6").Value))
    
    ' Mostrar estado
    mensaje = "ESTADO ACTUAL DE B�SQUEDA EN BDAS" & vbCrLf & _
              "--------------------------------" & vbCrLf & vbCrLf
    
    If estadoActual = "S�" Or estadoActual = "SI" Then
        mensaje = mensaje & "ACTIVADO" & vbCrLf & vbCrLf & _
                  "Las b�squedas de DNI incluyen las hojas BDAS."
    Else
        mensaje = mensaje & "DESACTIVADO" & vbCrLf & vbCrLf & _
                  "Las b�squedas de DNI solo se realizan en la misma hoja."
    End If
    
    MsgBox mensaje, vbInformation, "Estado de Configuraci�n"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al consultar estado: " & Err.Description, vbCritical, "Error"
End Sub

