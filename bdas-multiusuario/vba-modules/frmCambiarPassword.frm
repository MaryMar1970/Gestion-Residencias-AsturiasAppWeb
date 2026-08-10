VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCambiarPassword 
   Caption         =   "CAMBIO DE CONTRASE�A OBLIGATORIO"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4920
   OleObjectBlob   =   "frmCambiarPassword.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmCambiarPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private usuarioActual As String

' Propiedad para establecer el usuario desde el login
Public Property Let usuario(valor As String)
    usuarioActual = valor
End Property

' Bot�n Cambiar Contrase�a
Private Sub btnCambiar_Click()
    Dim nuevaPassword As String
    Dim confirmarPassword As String
    Dim wsUsuarios As Worksheet
    Dim i As Long, lastRow As Long
    Dim logPassword As String
    
    ' Obtener valores
    nuevaPassword = Trim(Me.txtNuevaPassword.Value)
    confirmarPassword = Trim(Me.txtConfirmarPassword.Value)
    
    ' Validaci�n: campos no vac�os
    If nuevaPassword = "" Or confirmarPassword = "" Then
        MsgBox "Debes completar ambos campos.", vbExclamation, "Campos Vac�os"
        Exit Sub
    End If
    
    ' Validaci�n: contrase�as coinciden
    If nuevaPassword <> confirmarPassword Then
        MsgBox "Las contrase�as no coinciden. Int�ntalo de nuevo.", vbExclamation, "Error"
        Me.txtNuevaPassword.Value = ""
        Me.txtConfirmarPassword.Value = ""
        Me.txtNuevaPassword.SetFocus
        Exit Sub
    End If
    
    ' Validaci�n: fortaleza de contrase�a
    If Not ModuloConfigSegura.ValidarFortalezaPassword(nuevaPassword) Then
        Me.txtNuevaPassword.Value = ""
        Me.txtConfirmarPassword.Value = ""
        Me.txtNuevaPassword.SetFocus
        Exit Sub
    End If
    
    ' Actualizar contrase�a en la hoja USUARIOS
    Set wsUsuarios = ThisWorkbook.Sheets("USUARIOS")
    lastRow = wsUsuarios.Cells(wsUsuarios.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        If wsUsuarios.Cells(i, 1).Value = usuarioActual Then
            Dim resultado As Variant
            Dim nuevoSalt As String
            Dim nuevoHash As String
        
            ' Generar nuevo salt y hash para la contrase�a
            resultado = ModuloConfigSegura.CrearPasswordHash(nuevaPassword)
            nuevoSalt = resultado(0)
            nuevoHash = resultado(1)
        
            ' Desproteger hoja para actualizar
            Dim passwordUsuarios As String
            passwordUsuarios = ModuloConfigSegura.ObtenerPasswordHojaUsuarios()
            wsUsuarios.Unprotect password:=passwordUsuarios

            ' Actualizar Salt y Hash (columnas C y D)
            wsUsuarios.Cells(i, 3).NumberFormat = "@"
            wsUsuarios.Cells(i, 3).Value = nuevoSalt
            wsUsuarios.Cells(i, 4).NumberFormat = "@"
            wsUsuarios.Cells(i, 4).Value = "'" & nuevoHash

            ' Marcar que ya no necesita cambiar contrase�a
            wsUsuarios.Cells(i, 2).Value = "NO"

            ' Proteger hoja de nuevo
            wsUsuarios.Protect password:=passwordUsuarios, UserInterfaceOnly:=True
            Exit For
        End If
    Next i
    
    ' Registrar en LOG
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    Call RegistrarCambioPasswordEnLog(usuarioActual, logPassword)
    
    ' Mensaje de confirmaci�n
    MsgBox "Contrase�a cambiada exitosamente." & vbCrLf & vbCrLf & _
           "Por favor, recuerda tu nueva contrase�a.", _
           vbInformation, "Cambio Exitoso"
    
    ' Cerrar formulario
    Unload Me
End Sub

' Bot�n Cancelar/Salir
Private Sub btnCancelar_Click()
    Dim respuesta As VbMsgBoxResult
    
    respuesta = MsgBox("Si no cambias tu contrase�a ahora, se cerrar� la aplicaci�n." & vbCrLf & vbCrLf & _
                       "�Est�s seguro de que deseas salir sin cambiar la contrase�a?", _
                       vbExclamation + vbYesNo, "Confirmar Salida")
    
    If respuesta = vbYes Then
        Unload Me
        ' Cerrar el libro
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

' Funci�n para registrar el cambio en el LOG
Private Sub RegistrarCambioPasswordEnLog(usuario As String, logPass As String)
    Dim wsLOG As Worksheet
    Dim logSheets As Variant
    Dim i As Long, filaLog As Long
    
    logSheets = Array("LOG_GIJ�N", "LOG_SOTO", "LOG_OVIEDO")
    
    For i = LBound(logSheets) To UBound(logSheets)
        On Error Resume Next
        Set wsLOG = ThisWorkbook.Sheets(logSheets(i))
        If Not wsLOG Is Nothing Then
            wsLOG.Unprotect password:=logPass
            filaLog = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            wsLOG.Cells(filaLog, 1).Value = usuario
            wsLOG.Cells(filaLog, 2).Value = Now
            wsLOG.Cells(filaLog, 3).Value = "CAMBIO DE CONTRASE�A (Obligatorio por seguridad)"
            wsLOG.Cells(filaLog, 4).Value = ""
            wsLOG.Cells(filaLog, 5).Value = ""
            wsLOG.Protect password:=logPass, UserInterfaceOnly:=True
        End If
        On Error GoTo 0
    Next i
End Sub

' Al cargar el formulario
Private Sub UserForm_Initialize()
    Me.txtNuevaPassword.SetFocus
End Sub

