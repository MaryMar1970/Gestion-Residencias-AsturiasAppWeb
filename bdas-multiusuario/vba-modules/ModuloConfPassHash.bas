Attribute VB_Name = "ModuloConfPassHash"
Option Explicit

'=================================================================================
' M�dulo: Migraci�n de Contrase�as a Hash
' Descripci�n: Script para migrar contrase�as de texto plano a sistema hash+salt
' Uso: Ejecutar una sola vez, luego puede eliminarse este m�dulo
' Fecha: 2025-11-16
' PASO 5: Sistema de Hash + Salt
'=================================================================================

Sub MigrarPasswordsAHash()
    Dim wsUsuarios As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim passwordActual As String
    Dim salt As String
    Dim hash As String
    Dim resultado As Variant
    Dim contador As Long
    Dim mensaje As String
    
    ' Confirmaci�n de seguridad
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("ATENCI�N: Esta operaci�n convertir� todas las contrase�as actuales a hash." & vbCrLf & vbCrLf & _
                       "Las contrase�as en texto plano NO se eliminar�n autom�ticamente." & vbCrLf & _
                       "Podr�s revisarlas antes de eliminarlas." & vbCrLf & vbCrLf & _
                       "�Deseas continuar?", _
                       vbQuestion + vbYesNo, "Confirmar Migraci�n")
    
    If respuesta = vbNo Then
        MsgBox "Operaci�n cancelada.", vbInformation
        Exit Sub
    End If
    
    ' Acceder a hoja USUARIOS
    On Error GoTo ErrorHandler
    Set wsUsuarios = ThisWorkbook.Sheets("USUARIOS")
    lastRow = wsUsuarios.Cells(wsUsuarios.Rows.Count, "A").End(xlUp).Row
    
    mensaje = "MIGRACI�N DE CONTRASE�AS A HASH" & vbCrLf & vbCrLf
    contador = 0
    
    ' Desproteger hoja USUARIOS si est� protegida
    Dim passwordUsuarios As String
    passwordUsuarios = ModuloConfigSegura.ObtenerPasswordHojaUsuarios()
    On Error Resume Next
    wsUsuarios.Unprotect password:=passwordUsuarios
    On Error GoTo ErrorHandler
    
    ' Procesar cada usuario
    For i = 2 To lastRow
        ' Obtener contrase�a actual (columna B)
        passwordActual = Trim(wsUsuarios.Cells(i, 2).Value)
        
        ' Solo procesar si hay contrase�a y no tiene hash ya
        If passwordActual <> "" And Trim(wsUsuarios.Cells(i, 5).Value) = "" Then
            ' Crear salt y hash
            resultado = ModuloConfigSegura.CrearPasswordHash(passwordActual)
            salt = resultado(0)
            hash = resultado(1)
            
            ' Guardar en columnas D (Salt) y E (PasswordHash)
            wsUsuarios.Cells(i, 4).Value = salt
            wsUsuarios.Cells(i, 5).Value = hash
            
            ' Contador y log
            contador = contador + 1
            mensaje = mensaje & "? " & wsUsuarios.Cells(i, 1).Value & vbCrLf
            mensaje = mensaje & "   Salt: " & salt & vbCrLf
            mensaje = mensaje & "   Hash: " & hash & vbCrLf & vbCrLf
        ElseIf Trim(wsUsuarios.Cells(i, 5).Value) <> "" Then
            mensaje = mensaje & "?? " & wsUsuarios.Cells(i, 1).Value & " - Ya tiene hash (omitido)" & vbCrLf & vbCrLf
        End If
    Next i
    
    mensaje = mensaje & "????????????????????????" & vbCrLf
    mensaje = mensaje & "Total migrados: " & contador & " usuarios" & vbCrLf & vbCrLf
    mensaje = mensaje & "?? IMPORTANTE: Las contrase�as en texto plano (columna B) a�n est�n visibles." & vbCrLf
    mensaje = mensaje & "Una vez que pruebes el sistema, podr�s eliminar esa columna."
    
    MsgBox mensaje, vbInformation, "Migraci�n Completada"
    
    ' Proteger hoja USUARIOS de nuevo
    wsUsuarios.Protect password:=passwordUsuarios, UserInterfaceOnly:=True
    
    ' Guardar cambios
    ThisWorkbook.Save
    
    Exit Sub

ErrorHandler:
    MsgBox "Error en la migraci�n: " & Err.Description, vbCritical, "Error"
End Sub

' =====================================================================================
' Sub para probar que el hash funciona correctamente ANTES de actualizar el login
' =====================================================================================
Sub ProbarValidacionHash()
    Dim wsUsuarios As Worksheet
    Dim usuario As String
    Dim passwordPrueba As String
    Dim salt As String
    Dim hash As String
    Dim esValida As Boolean
    Dim mensaje As String
    
    ' Pedir datos de prueba
    usuario = InputBox("Ingresa el nombre de usuario para probar:", "Prueba de Hash")
    If usuario = "" Then Exit Sub
    
    passwordPrueba = InputBox("Ingresa la contrase�a para validar:", "Prueba de Hash")
    If passwordPrueba = "" Then Exit Sub
    
    ' Buscar usuario
    On Error GoTo ErrorHandler
    Set wsUsuarios = ThisWorkbook.Sheets("USUARIOS")
    Dim i As Long, lastRow As Long
    lastRow = wsUsuarios.Cells(wsUsuarios.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If wsUsuarios.Cells(i, 1).Value = usuario Then
            salt = wsUsuarios.Cells(i, 4).Value
            hash = wsUsuarios.Cells(i, 5).Value
            
            If salt = "" Or hash = "" Then
                MsgBox "? Este usuario no tiene hash generado todav�a." & vbCrLf & vbCrLf & _
                       "Ejecuta primero: MigrarPasswordsAHash", vbExclamation, "Sin Hash"
                Exit Sub
            End If
            
            ' Validar
            esValida = ModuloConfigSegura.ValidarPasswordHash(passwordPrueba, salt, hash)
            
            If esValida Then
                mensaje = "? CONTRASE�A CORRECTA" & vbCrLf & vbCrLf
                mensaje = mensaje & "Usuario: " & usuario & vbCrLf
                mensaje = mensaje & "Salt: " & salt & vbCrLf
                mensaje = mensaje & "Hash: " & hash & vbCrLf & vbCrLf
                mensaje = mensaje & "?? La validaci�n de hash funciona correctamente."
                MsgBox mensaje, vbInformation, "Validaci�n Exitosa"
            Else
                mensaje = "? CONTRASE�A INCORRECTA" & vbCrLf & vbCrLf
                mensaje = mensaje & "El hash no coincide con la contrase�a ingresada."
                MsgBox mensaje, vbCritical, "Validaci�n Fallida"
            End If
            
            Exit Sub
        End If
    Next i
    
    MsgBox "? Usuario '" & usuario & "' no encontrado.", vbExclamation, "Usuario No Encontrado"
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

