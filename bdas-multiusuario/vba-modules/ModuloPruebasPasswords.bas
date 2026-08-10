Attribute VB_Name = "ModuloPruebasPasswords"
'Attribute VB_Name = "ModuloPruebasPasswords"
Option Explicit

'=================================================================================
' M�dulo: Pruebas y Utilidades del Sistema de Contrase�as
' Descripci�n: Herramientas de verificaci�n y gesti�n para administradores
' Versi�n: 1.1 - Con sincronizaci�n de protecciones
' Fecha: 2025-11-22
' Usuario: Bustiello2
'=================================================================================

'=================================================================================
' FUNCI�N PRINCIPAL: Abrir formulario de gesti�n
'=================================================================================

Sub AbrirGestionPasswords()
    '=================================================================
    ' Abrir formulario visual de gesti�n de contrase�as
    ' USO: Ejecutar esta macro para administrar contrase�as
    '=================================================================
    frmGestionPasswords.Show
End Sub

'=================================================================================
' DIAGN�STICO Y VERIFICACI�N
'=================================================================================

Sub DIAGNOSTICO_Sistema_Passwords()
    '=================================================================
    ' Diagn�stico completo del sistema de contrase�as
    ' Verifica que todo funcione correctamente
    '=================================================================
    
    Dim resultado As String
Dim ws As Worksheet
Dim pwd As String
Dim totalPasswords As Long
Dim i As Long
Dim nombrePwd As String
Dim usuarioActual As String

' ? Obtener usuario de sesi�n
usuarioActual = Trim(ThisWorkbook.usuarioActual)
If usuarioActual = "" Then usuarioActual = "NO_IDENTIFICADO"

resultado = "+---------------------------------------------------+" & vbCrLf
resultado = resultado & "�   DIAGN�STICO DEL SISTEMA DE CONTRASE�AS        �" & vbCrLf
resultado = resultado & "�   Usuario: " & usuarioActual & String(35 - Len(usuarioActual), " ") & "�" & vbCrLf
resultado = resultado & "�   Fecha: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & "                  �" & vbCrLf
resultado = resultado & "+---------------------------------------------------+" & vbCrLf & vbCrLf

' Verificar CONFIG_PWD
On Error Resume Next
Set ws = ThisWorkbook.Sheets("CONFIG_PWD")
On Error GoTo 0
    If ws Is Nothing Then
        resultado = resultado & "? HOJA CONFIG_PWD NO EXISTE" & vbCrLf
        resultado = resultado & "   El sistema no est� inicializado." & vbCrLf
        MsgBox resultado, vbCritical, "Error del Sistema"
        Exit Sub
    End If
    
    totalPasswords = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
    
    resultado = resultado & "? Sistema inicializado correctamente" & vbCrLf
    resultado = resultado & "   Total de contrase�as: " & totalPasswords & vbCrLf & vbCrLf
    
    ' Listar y verificar cada contrase�a
    resultado = resultado & "CONTRASE�AS DEL SISTEMA:" & vbCrLf
    resultado = resultado & String(50, "-") & vbCrLf
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        nombrePwd = ws.Cells(i, 1).Value
        
        On Error Resume Next
        pwd = ModuloConfigSegura.ObtenerPassword(nombrePwd)
        
        If Err.Number = 0 And pwd <> "" Then
            resultado = resultado & "? " & nombrePwd & vbCrLf
            resultado = resultado & "   +- Longitud: " & Len(pwd) & " caracteres" & vbCrLf
            resultado = resultado & "   +- Modificada: " & ws.Cells(i, 3).Value & vbCrLf
            resultado = resultado & "   +- Por: " & ws.Cells(i, 4).Value & vbCrLf
            resultado = resultado & "   +- Descripci�n: " & ws.Cells(i, 5).Value & vbCrLf
        Else
            resultado = resultado & "? " & nombrePwd & " - ERROR al recuperar" & vbCrLf
        End If
        Err.Clear
        On Error GoTo 0
        
        resultado = resultado & vbCrLf
    Next i
    
    ' Verificar funciones principales
    resultado = resultado & "FUNCIONES PRINCIPALES:" & vbCrLf
    resultado = resultado & String(50, "-") & vbCrLf
    
    On Error Resume Next
    
    pwd = ModuloConfigSegura.ObtenerClaveMaestra()
    resultado = resultado & IIf(pwd <> "", "?", "?") & " ObtenerClaveMaestra()" & vbCrLf
    
    pwd = ModuloConfigSegura.ObtenerPasswordHojasYEstructura()
    resultado = resultado & IIf(pwd <> "", "?", "?") & " ObtenerPasswordHojasYEstructura()" & vbCrLf
        
    On Error GoTo 0
    
    resultado = resultado & vbCrLf & String(50, "-") & vbCrLf
    resultado = resultado & "SISTEMA OPERATIVO CORRECTAMENTE" & vbCrLf
    
    MsgBox resultado, vbInformation, "Diagn�stico del Sistema"
    Debug.Print resultado
End Sub

Sub VER_Contrase�as_Actuales()
    '=================================================================
    ' Mostrar las contrase�as actuales del sistema
    ' ADVERTENCIA: �salo solo en entorno seguro
    ' ACTUALIZADO: 2025-11-22 01:54 UTC - Sistema de 2 contrase�as
    '=================================================================
    
    Dim resultado As String
    Dim respuesta As Integer
    
    respuesta = MsgBox("ADVERTENCIA DE SEGURIDAD" & vbCrLf & vbCrLf & _
                       "Vas a ver todas las contrase�as en texto plano." & vbCrLf & _
                       "Asegurate de que nadie mas esta mirando tu pantalla." & vbCrLf & vbCrLf & _
                       "�Deseas continuar?", _
                       vbYesNo + vbExclamation, "Advertencia de Seguridad")
    
    If respuesta = vbNo Then Exit Sub
    
    resultado = "CONTRASE�AS DEL SISTEMA" & vbCrLf
    resultado = resultado & String(60, "=") & vbCrLf & vbCrLf
    
    resultado = resultado & "1. CLAVE_MAESTRA" & vbCrLf
    resultado = resultado & "   Contrase�a: " & ModuloConfigSegura.ObtenerClaveMaestra() & vbCrLf
    resultado = resultado & "   Uso: Login del sistema + Acceso BDAS" & vbCrLf & vbCrLf
    
    resultado = resultado & "2. PASSWORD_HOJAS_Y_ESTRUCTURA" & vbCrLf
    resultado = resultado & "   Contrase�a: " & ModuloConfigSegura.ObtenerPasswordHojasYEstructura() & vbCrLf
    resultado = resultado & "   Uso: Hojas protegidas + Estructura" & vbCrLf & vbCrLf
    
    resultado = resultado & String(60, "-") & vbCrLf
    resultado = resultado & "PROTECCION DE BDAS:" & vbCrLf
    resultado = resultado & "La hoja BDAS esta protegida con: BDAS2025" & vbCrLf
    resultado = resultado & "Acceso validado con: CLAVE_MAESTRA" & vbCrLf & vbCrLf
    
    resultado = resultado & String(60, "=") & vbCrLf
    resultado = resultado & "NO COMPARTAS ESTA INFORMACION" & vbCrLf
    
    MsgBox resultado, vbInformation, "Contrase�as del Sistema"
    Debug.Print resultado
End Sub

Sub VER_Hoja_CONFIG_PWD()
    '=================================================================
    ' Mostrar temporalmente la hoja CONFIG_PWD para inspecci�n
    '=================================================================
    
    Dim ws As Worksheet
    Dim passwordProteccion As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CONFIG_PWD")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La hoja CONFIG_PWD no existe.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Desproteger
    passwordProteccion = ModuloConfigSegura.ObtenerPasswordHojasYEstructura()
    On Error Resume Next
    ws.Unprotect password:=passwordProteccion
    On Error GoTo 0
    
    ' Hacer visible
    ws.visible = xlSheetVisible
    ws.Activate
    ws.Columns("A:E").AutoFit
    
    MsgBox "Hoja CONFIG_PWD ahora visible." & vbCrLf & vbCrLf & _
           "ESTRUCTURA:" & vbCrLf & _
           "---------------------------------" & vbCrLf & _
           "Columna A: Nombre de la contrase�a" & vbCrLf & _
           "Columna B: Valor ofuscado (ROT1)" & vbCrLf & _
           "Columna C: Fecha de modificaci�n" & vbCrLf & _
           "Columna D: Usuario que modific�" & vbCrLf & _
           "Columna E: Descripci�n" & vbCrLf & vbCrLf & _
           "?? NO EDITES MANUALMENTE" & vbCrLf & _
           "Usa el formulario de gesti�n." & vbCrLf & vbCrLf & _
           "Cuando cierres este mensaje, la hoja se ocultar�.", _
           vbInformation, "Hoja CONFIG_PWD"
    
    ' Volver a ocultar y proteger
    ws.visible = xlSheetVeryHidden
    On Error Resume Next
    ws.Protect password:=passwordProteccion, UserInterfaceOnly:=True
    On Error GoTo 0
    
    MsgBox "Hoja CONFIG_PWD ocultada de nuevo.", vbInformation
End Sub

Sub SINCRONIZAR_Todas_Protecciones()
    '=================================================================
    ' Sincroniza TODAS las protecciones de hojas con las contrase�as
    ' actuales del sistema CONFIG_PWD
    ' ACTUALIZADO: 2025-11-22 02:23 UTC - Sistema de 2 contrase�as
    ' Usuario: Bustiello2
    '=================================================================
    
    Dim resultado As Integer
    Dim ws As Worksheet
    Dim passwordHojasYEstructura As String
    Dim passwordClaveMaestra As String
    Dim passwordAntiguaEstructura As String
    Dim contadorActualizados As Integer
    
    resultado = MsgBox("SINCRONIZACION DE PROTECCIONES" & vbCrLf & vbCrLf & _
                       "Esta macro actualizara las protecciones de:" & vbCrLf & vbCrLf & _
                       "- Hojas LOG (Gijon, Soto, Oviedo)" & vbCrLf & _
                       "- Hoja USUARIOS" & vbCrLf & _
                       "- Hoja CONFIG" & vbCrLf & _
                       "- Hoja CONFIG_PWD" & vbCrLf & _
                       "- Estructura del libro" & vbCrLf & vbCrLf & _
                       "NOTA: BDAS mantiene su proteccion independiente (BDAS2025)" & vbCrLf & vbCrLf & _
                       "�Deseas continuar?", _
                       vbYesNo + vbQuestion, "Sincronizacion")
    
    If resultado = vbNo Then Exit Sub
    
    ' Obtener contrase�as actuales del sistema
    passwordHojasYEstructura = ModuloConfigSegura.ObtenerPasswordHojasYEstructura()
    passwordClaveMaestra = ModuloConfigSegura.ObtenerClaveMaestra()
    
    ' Solicitar contrase�a antigua
    passwordAntiguaEstructura = InputBox("Ingresa la contrase�a ANTIGUA de hojas/estructura:" & vbCrLf & vbCrLf & _
                                        "(Si no la cambiaste, era: Log$Residencias2025)", _
                                        "Contrase�a Antigua Estructura", "Log$Residencias2025")
    
    If passwordAntiguaEstructura = "" Then Exit Sub
    
    contadorActualizados = 0
    
    On Error GoTo ErrorHandler
    
    ' ===== ACTUALIZAR HOJAS LOG =====
    Dim hojasLOG As Variant
    hojasLOG = Array("LOG_GIJ�N", "LOG_SOTO", "LOG_OVIEDO")
    
    Dim i As Long
    For i = LBound(hojasLOG) To UBound(hojasLOG)
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(hojasLOG(i))
        
        
        
        
        If Not ws Is Nothing Then
            ws.Unprotect password:=passwordAntiguaEstructura
            If Err.Number = 0 Then
                ws.Protect password:=passwordHojasYEstructura, UserInterfaceOnly:=True
                Debug.Print "OK " & hojasLOG(i) & " actualizado"
                contadorActualizados = contadorActualizados + 1
            Else
                Debug.Print "AVISO " & hojasLOG(i) & " - Error al desproteger"
            End If
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next i
    
    ' ===== ACTUALIZAR HOJA USUARIOS =====
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("USUARIOS")
    If Not ws Is Nothing Then
        ws.Unprotect password:=passwordAntiguaEstructura
        If Err.Number = 0 Then
            ws.Protect password:=passwordHojasYEstructura, UserInterfaceOnly:=True
            Debug.Print "OK USUARIOS actualizado"
            contadorActualizados = contadorActualizados + 1
        End If
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' ===== ACTUALIZAR HOJA CONFIG =====
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CONFIG")
    If Not ws Is Nothing Then
        ws.Unprotect password:=passwordAntiguaEstructura
        If Err.Number = 0 Then
            ws.Protect password:=passwordHojasYEstructura, UserInterfaceOnly:=True
            Debug.Print "OK CONFIG actualizado"
            contadorActualizados = contadorActualizados + 1
        End If
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' ===== ACTUALIZAR HOJA CONFIG_PWD =====
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CONFIG_PWD")
    If Not ws Is Nothing Then
        ws.Unprotect password:=passwordAntiguaEstructura
        If Err.Number = 0 Then
            ws.Protect password:=passwordHojasYEstructura, UserInterfaceOnly:=True
            Debug.Print "OK CONFIG_PWD actualizado"
            contadorActualizados = contadorActualizados + 1
        End If
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' ===== ACTUALIZAR ESTRUCTURA DEL LIBRO =====
    On Error Resume Next
    ThisWorkbook.Unprotect password:=passwordAntiguaEstructura
    If Err.Number = 0 Then
        ThisWorkbook.Protect password:=passwordHojasYEstructura, Structure:=True, Windows:=False
        Debug.Print "OK Estructura del libro actualizada"
        contadorActualizados = contadorActualizados + 1
    Else
        Debug.Print "AVISO Estructura del libro - Error al desproteger"
    End If
    Err.Clear
    On Error GoTo ErrorHandler
    
    ' ===== MENSAJE FINAL =====
    MsgBox "SINCRONIZACION COMPLETADA" & vbCrLf & vbCrLf & _
           "Elementos actualizados: " & contadorActualizados & vbCrLf & vbCrLf & _
           "CONTRASE�AS APLICADAS:" & vbCrLf & _
           String(50, "=") & vbCrLf & _
           "- Hojas/Estructura: " & passwordHojasYEstructura & vbCrLf & vbCrLf & _
           "NOTA: BDAS mantiene proteccion independiente (BDAS2025)." & vbCrLf & _
           "      No requiere sincronizacion." & vbCrLf & vbCrLf & _
           "Revisa la Ventana Inmediato (Ctrl+G) para detalles.", _
           vbInformation, "Sincronizacion Exitosa"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "ERROR durante la sincronizacion:" & vbCrLf & vbCrLf & _
           "Numero: " & Err.Number & vbCrLf & _
           "Descripcion: " & Err.Description & vbCrLf & vbCrLf & _
           "Elementos actualizados hasta el error: " & contadorActualizados & vbCrLf & vbCrLf & _
           "Verifica las contrase�as antiguas y vuelve a intentar.", _
           vbCritical, "Error de Sincronizacion"
End Sub '=================================================================================
' UTILIDADES PARA DESARROLLADORES
'=================================================================================

Sub GENERAR_Codigo_Ofuscado()
    '=================================================================
    ' Genera c�digo ofuscado para usar en VBA
    ' �til al a�adir nuevas contrase�as hardcoded
    '=================================================================
    
    Dim password As String
    
    password = InputBox("Ingresa la contrase�a para generar c�digo ofuscado:" & vbCrLf & vbCrLf & _
                        "Esto genera el c�digo necesario para" & vbCrLf & _
                        "a�adir contrase�as hardcoded en el c�digo VBA.", _
                        "Generador de C�digo Ofuscado", "")
    
    If password <> "" Then
        Call ModuloConfigSegura.GenerarCodigoOfuscado(password)
    End If
End Sub

Sub EXPORTAR_Configuracion_Passwords()
    '=================================================================
    ' Exporta la configuraci�n actual a un archivo de texto
    ' �til para backup o documentaci�n
    '=================================================================
    
    Dim ws As Worksheet
    Dim fso As Object
    Dim archivo As Object
    Dim rutaArchivo As String
    Dim i As Long
    Dim contenido As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CONFIG_PWD")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "La hoja CONFIG_PWD no existe.", vbExclamation
        Exit Sub
    End If
    
    ' Crear archivo de texto
    rutaArchivo = ThisWorkbook.Path & "\CONFIG_PASSWORDS_" & Format(Now, "yyyymmdd_hhmmss") & ".txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set archivo = fso.CreateTextFile(rutaArchivo, True)
    
Dim usuarioActual As String

' ? Obtener usuario de sesi�n
usuarioActual = Trim(ThisWorkbook.usuarioActual)
If usuarioActual = "" Then usuarioActual = "NO_IDENTIFICADO"

' Escribir contenido
contenido = "CONFIGURACI�N DE CONTRASE�AS DEL SISTEMA" & vbCrLf
contenido = contenido & "Fecha exportaci�n: " & Now & vbCrLf
contenido = contenido & "Usuario: " & usuarioActual & vbCrLf
contenido = contenido & String(70, "=") & vbCrLf & vbCrLf
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        contenido = contenido & "Nombre: " & ws.Cells(i, 1).Value & vbCrLf
        contenido = contenido & "Descripci�n: " & ws.Cells(i, 5).Value & vbCrLf
        contenido = contenido & "�ltima modificaci�n: " & ws.Cells(i, 3).Value & vbCrLf
        contenido = contenido & "Modificada por: " & ws.Cells(i, 4).Value & vbCrLf
        contenido = contenido & String(70, "-") & vbCrLf
    Next i
    
    contenido = contenido & vbCrLf & "NOTA: Las contrase�as NO est�n incluidas por seguridad."
    
    archivo.WriteLine contenido
    archivo.Close
    
    MsgBox "Configuraci�n exportada exitosamente:" & vbCrLf & vbCrLf & _
           rutaArchivo, vbInformation, "Exportaci�n Completa"
End Sub

'=================================================================================
' NOTAS PARA EL DESARROLLADOR:
'
' FUNCIONES DISPONIBLES:
' ----------------------
' - AbrirGestionPasswords: Abre el formulario visual de gesti�n
' - DIAGNOSTICO_Sistema_Passwords: Diagn�stico completo del sistema
' - VER_Contrase�as_Actuales: Muestra las 3 contrase�as en texto plano
' - VER_Hoja_CONFIG_PWD: Muestra temporalmente la hoja oculta
' - SINCRONIZAR_Todas_Protecciones: Actualiza protecciones de hojas (NUEVO)
' - GENERAR_Codigo_Ofuscado: Genera c�digo ofuscado para VBA
' - EXPORTAR_Configuracion_Passwords: Exporta configuraci�n a archivo
'
' ESTRUCTURA DEL SISTEMA:
' -----------------------
' 3 Contrase�as principales:
' 1. CLAVE_MAESTRA ? Login
' 2. PASSWORD_HOJAS_Y_ESTRUCTURA ? Hojas + Estructura
' 3. PASSWORD_BDAS ? Base de Datos
'
' SINCRONIZACI�N:
' ---------------
' Cuando cambies una contrase�a en el formulario, ejecuta:
' SINCRONIZAR_Todas_Protecciones para actualizar las protecciones de hojas.
'
' ACCESO A CONTRASE�AS EN C�DIGO:
' --------------------------------
' pwd = ModuloConfigSegura.ObtenerClaveMaestra()
' pwd = ModuloConfigSegura.ObtenerPasswordHojasYEstructura()

' O usar la funci�n universal:
' pwd = ModuloConfigSegura.ObtenerPassword("NOMBRE_PASSWORD")
'
'=================================================================================

