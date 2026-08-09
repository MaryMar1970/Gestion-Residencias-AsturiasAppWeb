VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin
   Caption         =   "BDAS — Inicio de Sesión"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================================
' UserForm: frmLogin
' Propósito: Formulario de login manual para el BDAS Multiusuario.
'            Solicita usuario y contraseña, valida contra tabla Usuarios de Access.
'
' INSTRUCCIONES PARA CREAR ESTE FORMULARIO EN VBA:
' 1. Abrir el Editor VBA (ALT+F11)
' 2. Insertar > UserForm
' 3. Renombrar a "frmLogin" (en la ventana de Propiedades)
' 4. Añadir los siguientes controles:
'
'    Control          | Nombre          | Propiedades clave
'    -----------------|-----------------|------------------------------------------
'    Label            | lblTitulo       | Caption="BDAS — Residencias", Font=14pt Bold
'    Label            | lblSubtitulo    | Caption="Inicio de Sesión"
'    Label            | lblUsuario      | Caption="Usuario:"
'    TextBox          | txtUsuario      | TabIndex=0
'    Label            | lblClave        | Caption="Contraseña:"
'    TextBox          | txtClave        | PasswordChar="*", TabIndex=1
'    Label            | lblError        | Caption="", ForeColor=&H000000FF (rojo), Visible=False
'    CommandButton    | btnEntrar       | Caption="Entrar", Default=True, TabIndex=2
'    CommandButton    | btnSalir        | Caption="Salir", Cancel=True, TabIndex=3
'
' 5. Pegar este código en el módulo del formulario
' ==============================================================================

Option Explicit

Private m_Cancelado As Boolean

Private Sub UserForm_Initialize()
    ' Configuración inicial del formulario
    Me.Caption = "BDAS — Inicio de Sesión"
    Me.StartUpPosition = 1  ' CenterOwner
    
    ' Limpiar campos
    txtUsuario.Text = ""
    txtClave.Text = ""
    lblError.Visible = False
    m_Cancelado = False
    
    ' Foco en el campo de usuario
    txtUsuario.SetFocus
End Sub

Private Sub btnEntrar_Click()
    Dim usuario As String
    Dim clave As String
    
    usuario = Trim(txtUsuario.Text)
    clave = txtClave.Text
    
    ' Validar campos vacíos
    If Len(usuario) = 0 Then
        MostrarError "Introduce tu nombre de usuario."
        txtUsuario.SetFocus
        Exit Sub
    End If
    
    If Len(clave) = 0 Then
        MostrarError "Introduce tu contraseña."
        txtClave.SetFocus
        Exit Sub
    End If
    
    ' Deshabilitar botón mientras valida
    btnEntrar.Enabled = False
    btnEntrar.Caption = "Validando..."
    DoEvents
    
    ' Validar credenciales contra Access
    If modDatabase.ValidarCredenciales(usuario, clave) Then
        ' Login exitoso — cerrar formulario
        Me.Hide
    Else
        ' Credenciales incorrectas
        MostrarError "Usuario o contraseña incorrectos."
        txtClave.Text = ""
        txtClave.SetFocus
    End If
    
    ' Rehabilitar botón
    btnEntrar.Enabled = True
    btnEntrar.Caption = "Entrar"
End Sub

Private Sub btnSalir_Click()
    m_Cancelado = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Si cierran con la X, tratar como "Salir"
    If CloseMode = vbFormControlMenu Then
        m_Cancelado = True
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub MostrarError(ByVal mensaje As String)
    lblError.Caption = mensaje
    lblError.Visible = True
End Sub

''' Indica si el usuario canceló el login.
Public Property Get Cancelado() As Boolean
    Cancelado = m_Cancelado
End Property

Private Sub txtUsuario_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Ocultar error al empezar a escribir
    lblError.Visible = False
End Sub

Private Sub txtClave_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Ocultar error al empezar a escribir
    lblError.Visible = False
End Sub
