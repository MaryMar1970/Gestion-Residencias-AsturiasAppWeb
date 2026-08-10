Attribute VB_Name = "GestionarAccesoUsuarios"
'========================================================================================
' M�dulo: Gesti�n de Usuarios para Excel
' Descripci�n:
'   Este m�dulo permite gestionar usuarios en una hoja llamada "USUARIOS" mediante:
'     - Alta o modificaci�n de usuario y contrase�a
'     - Eliminaci�n de usuario
'     - Consulta de lista de usuarios
'   Todas las acciones quedan registradas en una hoja llamada "LOG" (si existe).
'
' Uso:
'   Se invoca desde una macro o bot�n Ribbon (por ejemplo, desde el bot�n "Gesti�n de Usuarios").
'
' Correcci�n aplicada:
'   Se ha a�adido una verificaci�n para evitar mostrar "Opci�n no v�lida" si el usuario cancela
'   el primer InputBox (men� de opciones), saliendo silenciosamente del procedimiento.
'
' Requisitos:
'   - La hoja "USUARIOS" debe existir y tener al menos las columnas:
'         A: Nombre de usuario
'         B: Contrase�a
'   - Opcional: La hoja "LOG" para registrar acciones.
'
'========================================================================================

Sub GestionarUsuarios(control As IRibbonControl)
    ' === Declaraci�n de variables ===
    Dim ws As Worksheet, wsLOG As Worksheet
    Dim usuario As String, clave As String
    Dim fila As Long, encontrado As Boolean
    Dim i As Long, lastRow As Long, filaLog As Long
    Dim respuesta As VbMsgBoxResult
    Dim accion As String

    ' === Referencia a la hoja USUARIOS ===
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("USUARIOS")
    If ws Is Nothing Then
        MsgBox "No se encuentra la hoja 'USUARIOS'.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' === Men� de opciones principal ===
    accion = InputBox("�Qu� deseas hacer?" & vbCrLf & _
                      "1 - A�adir o modificar usuario" & vbCrLf & _
                      "2 - Eliminar usuario" & vbCrLf & _
                      "3 - Ver lista de usuarios", "Gesti�n de usuarios")
    If accion = "" Then Exit Sub ' <-- Corregido: salir si se cancela o deja vac�o

    Select Case Trim(accion)
    Case "1"
        ' === Alta o modificaci�n de usuario ===
        usuario = InputBox("Introduce el nombre de usuario:", "Gesti�n de usuarios")
        usuario = Trim(usuario)
        If usuario = "" Then MsgBox "El nombre de usuario no puede estar vac�o.": Exit Sub
        If InStr(usuario, " ") > 0 Then MsgBox "El nombre de usuario no puede contener espacios.": Exit Sub

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        encontrado = False
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = usuario Then
                encontrado = True
                fila = i
                Exit For
            End If
        Next i

        clave = InputBox("Introduce la contrase�a para el usuario '" & usuario & "':", "Gesti�n de usuarios")
        clave = Trim(clave)
        If clave = "" Then MsgBox "La contrase�a no puede estar vac�a.": Exit Sub
        
        ' -> VALIDAR FORTALEZA DE CONTRASE�A
        If Not ModuloConfigSegura.ValidarFortalezaPassword(clave) Then
            Exit Sub
        End If
        
        ' -> GENERAR SALT Y HASH
        Dim resultado As Variant
        Dim nuevoSalt As String
        Dim nuevoHash As String
        resultado = ModuloConfigSegura.CrearPasswordHash(clave)
        nuevoSalt = resultado(0)
        nuevoHash = resultado(1)
        
        ' -> DESPROTEGER HOJA
        Dim passwordUsuarios As String
        passwordUsuarios = ModuloConfigSegura.ObtenerPasswordHojaUsuarios()
        ws.Unprotect password:=passwordUsuarios

        If encontrado Then
            respuesta = MsgBox("El usuario ya existe. �Deseas modificar su contrase�a?", vbYesNo + vbQuestion)
            If respuesta = vbYes Then
                ' ? Actualizar con el formato correcto
                ws.Cells(fila, 2).Value = "NO"  ' �Debe Cambiar Password?
                ws.Cells(fila, 3).NumberFormat = "@"
                ws.Cells(fila, 3).Value = nuevoSalt  ' Salt
                ws.Cells(fila, 4).NumberFormat = "@"
                ws.Cells(fila, 4).Value = "'" & nuevoHash  ' PasswordHash
                
                MsgBox "Contrase�a actualizada correctamente.", vbInformation
                RegistrarLOG "MODIFICACI�N USUARIO: " & usuario
            Else
                MsgBox "No se realizaron cambios.", vbInformation
            End If
        Else
            ' -> A�adir nuevo usuario con el formato correcto
            fila = lastRow + 1
            ws.Cells(fila, 1).Value = usuario  ' Usuario
            ws.Cells(fila, 2).Value = "NO"  ' �Debe Cambiar Password?
            ws.Cells(fila, 3).NumberFormat = "@"
            ws.Cells(fila, 3).Value = nuevoSalt  ' Salt
            ws.Cells(fila, 4).NumberFormat = "@"
            ws.Cells(fila, 4).Value = "'" & nuevoHash  ' PasswordHash
            
            MsgBox "Usuario a�adido correctamente.", vbInformation
            RegistrarLOG "ALTA USUARIO: " & usuario
        End If
        
        ' -> PROTEGER HOJA DE NUEVO
        ws.Protect password:=passwordUsuarios, UserInterfaceOnly:=True

    Case "2"
        ' === Eliminaci�n de usuario ===
        usuario = InputBox("Introduce el nombre de usuario a eliminar:", "Eliminar usuario")
        usuario = Trim(usuario)
        If usuario = "" Then Exit Sub

        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        encontrado = False
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value = usuario Then
                encontrado = True
                fila = i
                Exit For
            End If
        Next i
                If encontrado Then
            respuesta = MsgBox("�Est�s seguro de que deseas eliminar al usuario '" & usuario & "'?", vbYesNo + vbExclamation)
            If respuesta = vbYes Then
                ' -> DESPROTEGER HOJA ANTES DE ELIMINAR
                passwordUsuarios = ModuloConfigSegura.ObtenerPasswordHojaUsuarios()
                ws.Unprotect password:=passwordUsuarios
                
                ' Eliminar fila
                ws.Rows(fila).Delete
                
                ' -> PROTEGER HOJA DE NUEVO
                ws.Protect password:=passwordUsuarios, UserInterfaceOnly:=True
                
                MsgBox "Usuario eliminado correctamente.", vbInformation
                RegistrarLOG "ELIMINACI�N USUARIO: " & usuario
            Else
                MsgBox "No se realiz� ninguna eliminaci�n.", vbInformation
            End If
        Else
            MsgBox "El usuario no existe.", vbExclamation
        End If

    Case "3"
        ' === Mostrar lista de usuarios ===
        Dim lista As String
        lista = "Usuarios registrados:" & vbCrLf & vbCrLf
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> "" Then
                lista = lista & "� " & ws.Cells(i, 1).Value & vbCrLf
            End If
        Next i
        MsgBox lista, vbInformation, "Lista de usuarios"

    Case Else
        MsgBox "Opci�n no v�lida.", vbExclamation
    End Select
End Sub

'----------------------------------------------------------------------------------------
' Procedimiento auxiliar para registrar acciones en la hoja LOG
'----------------------------------------------------------------------------------------
Private Sub RegistrarLOG(ByVal descripcion As String)
    ' Registra la acci�n en la hoja LOG, si existe, con:
    '   - Usuario actual (usar propiedad personalizada o adaptar seg�n sea necesario)
    '   - Fecha y hora de la acci�n
    '   - Descripci�n de la acci�n realizada

    Dim wsLOG As Worksheet
    Dim filaLog As Long

    On Error Resume Next
    Set wsLOG = ThisWorkbook.Sheets("LOG")
    If Not wsLOG Is Nothing Then
        filaLog = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
        wsLOG.Cells(filaLog, 1).Value = ThisWorkbook.usuarioActual ' <- adaptar seg�n implementaci�n
        wsLOG.Cells(filaLog, 2).Value = Now
        wsLOG.Cells(filaLog, 3).Value = descripcion
        wsLOG.Cells(filaLog, 4).Value = ""
        wsLOG.Cells(filaLog, 5).Value = ""
    End If
    On Error GoTo 0
End Sub

'========================================================================================
' Documentaci�n adicional:
'
' - El m�dulo est� dise�ado para usarse con hojas de c�lculo orientadas a la gesti�n de usuarios.
' - Todas las entradas y salidas se gestionan por InputBox y MsgBox para m�xima compatibilidad.
' - El control de errores b�sico est� presente para evitar interrupciones si no existen las hojas requeridas.
' - Para registrar el usuario actual, aseg�rate de tener una propiedad o m�todo para obtener el nombre del usuario.
'   Si no existe 'UsuarioActual' en ThisWorkbook, reemplaza por Application.Username u otro m�todo.
' - Si la hoja LOG no est� presente, las acciones no se registrar�n pero el flujo principal no se interrumpe.
'
'========================================================================================

