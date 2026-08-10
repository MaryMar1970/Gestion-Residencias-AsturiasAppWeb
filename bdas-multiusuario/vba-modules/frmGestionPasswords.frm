VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGestionPasswords 
   Caption         =   "Gesti�n de Contrase�as - Sistema Din�mico"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360
   OleObjectBlob   =   "frmGestionPasswords.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmGestionPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private modoEdicion As Boolean ' True si estamos editando una existente

Private Sub UserForm_Initialize()
    ' Se ejecuta al abrir el formulario
    Call CargarListaPasswords
    Call LimpiarCampos
    Me.txtNombre.Enabled = False
    modoEdicion = False
    
    ' Ocultar bot�n Generar C�digo (solo para desarrolladores)
    btnGenerarCodigo.visible = False
End Sub

Private Sub CargarListaPasswords()
    ' Cargar lista de contrase�as en el ComboBox
    Dim lista As Variant
    Dim i As Long
    
    cboPasswordSeleccionada.Clear
    cboPasswordSeleccionada.AddItem "-- Seleccionar --"
    
    On Error Resume Next
    lista = ModuloConfigSegura.ListarPasswords()
    On Error GoTo 0
    
    If IsArray(lista) Then
        For i = LBound(lista) To UBound(lista)
            cboPasswordSeleccionada.AddItem lista(i)
        Next i
    End If
    
    cboPasswordSeleccionada.ListIndex = 0
End Sub

Private Sub cboPasswordSeleccionada_Change()
    ' Al seleccionar una contrase�a de la lista
    If cboPasswordSeleccionada.ListIndex <= 0 Then
        Call LimpiarCampos
        txtNombre.Enabled = False
        modoEdicion = False
        Exit Sub
    End If
    
    Dim nombrePassword As String
    nombrePassword = cboPasswordSeleccionada.Value
    
    ' MODO EDICI�N: Mostrar campo de contrase�a actual
    modoEdicion = True
    
    ' Cargar datos de la contrase�a seleccionada
    txtNombre.Value = nombrePassword
    txtNombre.Enabled = False ' No se puede cambiar el nombre al editar
    txtPasswordActual.Value = ""
    txtNuevaPassword.Value = ""
    txtConfirmarPassword.Value = ""
    txtDescripcion.Value = ObtenerDescripcionPassword(nombrePassword)
    
    ' Mostrar campo de contrase�a actual
    lblPasswordActual.visible = True
    txtPasswordActual.visible = True
    
    lblFortaleza.Caption = "Fortaleza: Ingresa la contrase�a actual para modificar"
    lblFortaleza.ForeColor = RGB(0, 0, 255) ' Azul
    
    btnGuardar.Caption = "Actualizar"
End Sub

Private Function ObtenerDescripcionPassword(nombrePassword As String) As String
    ' Buscar descripci�n en CONFIG_PWD
    Dim ws As Worksheet
    Dim i As Long
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CONFIG_PWD")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ObtenerDescripcionPassword = ""
        Exit Function
    End If
    
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(nombrePassword)) Then
            ObtenerDescripcionPassword = ws.Cells(i, 5).Value
            Exit Function
        End If
    Next i
    
    ObtenerDescripcionPassword = ""
End Function

Private Sub txtPasswordActual_Change()
    ' Cuando escribe la contrase�a actual, validarla
    If modoEdicion And txtPasswordActual.Value <> "" Then
        Call ValidarPasswordActual
    End If
End Sub

Private Sub ValidarPasswordActual()
    ' Verificar que la contrase�a actual sea correcta
    Dim nombrePwd As String
    Dim passwordActualIngresada As String
    Dim passwordRealActual As String
    
    If Not modoEdicion Then Exit Sub
    
    nombrePwd = txtNombre.Value
    passwordActualIngresada = txtPasswordActual.Value
    
    ' Obtener la contrase�a actual del sistema
    passwordRealActual = ModuloConfigSegura.ObtenerPassword(nombrePwd)
    
    If passwordActualIngresada = passwordRealActual Then
        lblFortaleza.Caption = "? Contrase�a actual correcta. Ahora ingresa la nueva contrase�a"
        lblFortaleza.ForeColor = RGB(0, 150, 0) ' Verde
        txtNuevaPassword.Enabled = True
        txtConfirmarPassword.Enabled = True
    Else
        lblFortaleza.Caption = "? Contrase�a actual incorrecta"
        lblFortaleza.ForeColor = RGB(255, 0, 0) ' Rojo
        txtNuevaPassword.Enabled = False
        txtConfirmarPassword.Enabled = False
        txtNuevaPassword.Value = ""
        txtConfirmarPassword.Value = ""
    End If
End Sub

Private Sub txtNuevaPassword_Change()
    ' Validar fortaleza en tiempo real
    Call ValidarFortalezaEnTiempoReal
End Sub

Private Sub txtConfirmarPassword_Change()
    ' Validar que coincidan
    Call ValidarFortalezaEnTiempoReal
End Sub

Private Sub ValidarFortalezaEnTiempoReal()
    Dim pwd As String
    Dim fortaleza As String
    Dim color As Long
    
    pwd = txtNuevaPassword.Value
    
    ' Si estamos en modo edici�n y la contrase�a actual no es correcta, no evaluar
    If modoEdicion Then
        Dim passwordActual As String
        Dim passwordReal As String
        passwordActual = txtPasswordActual.Value
        passwordReal = ModuloConfigSegura.ObtenerPassword(txtNombre.Value)
        
        If passwordActual <> passwordReal Then
            Exit Sub ' Ya se mostr� el error en ValidarPasswordActual
        End If
    End If
    
    If Len(pwd) = 0 Then
        If modoEdicion Then
            lblFortaleza.Caption = "Fortaleza: Ingresa nueva contrase�a para cambiar"
        Else
            lblFortaleza.Caption = "Fortaleza: ---"
        End If
        lblFortaleza.ForeColor = RGB(128, 128, 128)
        Exit Sub
    End If
    
    ' Evaluar fortaleza
    Dim puntos As Integer
    puntos = 0
    
    If Len(pwd) >= 8 Then puntos = puntos + 1
    If Len(pwd) >= 12 Then puntos = puntos + 1
    If pwd Like "*[a-z]*" Then puntos = puntos + 1
    If pwd Like "*[A-Z]*" Then puntos = puntos + 1
    If pwd Like "*[0-9]*" Then puntos = puntos + 1
    If pwd Like "*[!@#$%^&*()_+=-]*" Then puntos = puntos + 1
    
    ' Determinar fortaleza
    Select Case puntos
        Case 0 To 2
            fortaleza = "MUY D�BIL ?"
            color = RGB(255, 0, 0) ' Rojo
        Case 3
            fortaleza = "D�BIL ??"
            color = RGB(255, 128, 0) ' Naranja
        Case 4
            fortaleza = "MEDIA ??"
            color = RGB(255, 200, 0) ' Amarillo
        Case 5
            fortaleza = "FUERTE ?"
            color = RGB(0, 200, 0) ' Verde
        Case Else
            fortaleza = "MUY FUERTE ??"
            color = RGB(0, 128, 0) ' Verde oscuro
    End Select
    
    ' Verificar que coincidan
    If txtConfirmarPassword.Value <> "" And txtNuevaPassword.Value <> txtConfirmarPassword.Value Then
        fortaleza = fortaleza & " | ?? Las contrase�as NO coinciden"
        color = RGB(255, 0, 0)
    ElseIf txtConfirmarPassword.Value <> "" And txtNuevaPassword.Value = txtConfirmarPassword.Value Then
        fortaleza = fortaleza & " | ? Coinciden"
    End If
    
    lblFortaleza.Caption = "Fortaleza: " & fortaleza
    lblFortaleza.ForeColor = color
End Sub

Private Sub chkMostrar_Click()
    ' Mostrar/ocultar contrase�as
    If chkMostrar.Value Then
        txtPasswordActual.PasswordChar = ""
        txtNuevaPassword.PasswordChar = ""
        txtConfirmarPassword.PasswordChar = ""
    Else
        txtPasswordActual.PasswordChar = "*"
        txtNuevaPassword.PasswordChar = "*"
        txtConfirmarPassword.PasswordChar = "*"
    End If
End Sub

Private Sub btnRefrescar_Click()
    ' Refrescar lista
    Call CargarListaPasswords
    MsgBox "Lista actualizada.", vbInformation, "Refrescar"
End Sub

Private Sub btnNueva_Click()
    ' Crear nueva contrase�a
    Call LimpiarCampos
    txtNombre.Enabled = True
    txtNombre.SetFocus
    cboPasswordSeleccionada.ListIndex = 0
    modoEdicion = False
    
    ' Ocultar campo de contrase�a actual
    lblPasswordActual.visible = False
    txtPasswordActual.visible = False
    
    ' Habilitar campos de nueva contrase�a
    txtNuevaPassword.Enabled = True
    txtConfirmarPassword.Enabled = True
    
    btnGuardar.Caption = "Guardar"
End Sub

Private Sub btnGuardar_Click()
    ' Guardar o actualizar contrase�a
    Dim nombrePwd As String
    Dim valorPwd As String
    Dim descripcion As String
    
    ' Validaciones
    nombrePwd = Trim(txtNombre.Value)
    valorPwd = txtNuevaPassword.Value
    descripcion = Trim(txtDescripcion.Value)
    
    If nombrePwd = "" Then
        MsgBox "Debes ingresar un nombre para la contrase�a.", vbExclamation, "Campo requerido"
        txtNombre.SetFocus
        Exit Sub
    End If
    
    ' Si estamos en modo edici�n, verificar contrase�a actual
    If modoEdicion Then
        Dim passwordActual As String
        Dim passwordReal As String
        
        passwordActual = txtPasswordActual.Value
        passwordReal = ModuloConfigSegura.ObtenerPassword(nombrePwd)
        
        If passwordActual = "" Then
            MsgBox "Debes ingresar la contrase�a actual para poder modificarla.", vbExclamation, "Verificaci�n requerida"
            txtPasswordActual.SetFocus
            Exit Sub
        End If
        
        If passwordActual <> passwordReal Then
            MsgBox "La contrase�a actual es incorrecta." & vbCrLf & vbCrLf & _
                   "No se puede modificar la contrase�a sin verificar la actual.", _
                   vbCritical, "Error de verificaci�n"
            txtPasswordActual.SetFocus
            Exit Sub
        End If
    End If
    
    If valorPwd = "" Then
        MsgBox "Debes ingresar una contrase�a.", vbExclamation, "Campo requerido"
        txtNuevaPassword.SetFocus
        Exit Sub
    End If
    
    If valorPwd <> txtConfirmarPassword.Value Then
        MsgBox "Las contrase�as no coinciden.", vbExclamation, "Error"
        txtConfirmarPassword.SetFocus
        Exit Sub
    End If
    
    ' Confirmaci�n adicional si estamos modificando una contrase�a cr�tica
    If modoEdicion And EsPasswordCritica(nombrePwd) Then
        Dim respuesta As Integer
        respuesta = MsgBox("ADVERTENCIA: Est�s modificando una contrase�a cr�tica del sistema:" & vbCrLf & vbCrLf & _
                          nombrePwd & vbCrLf & vbCrLf & _
                          "Esto puede afectar el funcionamiento del sistema." & vbCrLf & _
                          "�Est�s seguro de continuar?", _
                          vbYesNo + vbExclamation, "Contrase�a Cr�tica")
        
        If respuesta = vbNo Then Exit Sub
    End If
    
    ' Guardar usando el sistema din�mico
    Call ModuloConfigSegura.EstablecerPassword(nombrePwd, valorPwd, descripcion)
    
    ' Refrescar lista y limpiar
    Call CargarListaPasswords
    Call LimpiarCampos
    modoEdicion = False
    
    ' Seleccionar la contrase�a reci�n guardada
    Dim i As Long
    For i = 0 To cboPasswordSeleccionada.ListCount - 1
        If cboPasswordSeleccionada.List(i) = nombrePwd Then
            cboPasswordSeleccionada.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Function EsPasswordCritica(nombrePassword As String) As Boolean
    ' Solo 2 contrase�as cr�ticas en el sistema simplificado
    Dim criticas As Variant
    Dim i As Long
    
    criticas = Array("CLAVE_MAESTRA", "PASSWORD_HOJAS_Y_ESTRUCTURA")
    
    For i = LBound(criticas) To UBound(criticas)
        If UCase(Trim(nombrePassword)) = UCase(Trim(criticas(i))) Then
            EsPasswordCritica = True
            Exit Function
        End If
    Next i
    
    EsPasswordCritica = False
End Function

Private Sub btnEliminar_Click()
    ' Eliminar contrase�a seleccionada
    If cboPasswordSeleccionada.ListIndex <= 0 Then
        MsgBox "Selecciona una contrase�a para eliminar.", vbExclamation, "Selecci�n requerida"
        Exit Sub
    End If
    
    Dim nombrePwd As String
    Dim respuesta As Integer
    
    nombrePwd = cboPasswordSeleccionada.Value
    
    ' Verificar contrase�a actual antes de eliminar
    Dim passwordActual As String
    passwordActual = InputBox("Por seguridad, ingresa la contrase�a actual de '" & nombrePwd & "' para confirmar la eliminaci�n:", _
                             "Verificaci�n de Seguridad", "")
    
    If passwordActual = "" Then Exit Sub
    
    Dim passwordReal As String
    passwordReal = ModuloConfigSegura.ObtenerPassword(nombrePwd)
    
    If passwordActual <> passwordReal Then
        MsgBox "Contrase�a incorrecta. No se puede eliminar.", vbCritical, "Error de verificaci�n"
        Exit Sub
    End If
    
    respuesta = MsgBox("�Est�s seguro de eliminar la contrase�a '" & nombrePwd & "'?" & vbCrLf & vbCrLf & _
                       "Esta acci�n no se puede deshacer.", _
                       vbYesNo + vbExclamation, "Confirmar eliminaci�n")
    
    If respuesta = vbYes Then
        Call ModuloConfigSegura.EliminarPassword(nombrePwd)
        Call CargarListaPasswords
        Call LimpiarCampos
        modoEdicion = False
    End If
End Sub

Private Sub btnGenerarCodigo_Click()
    ' Generar c�digo ofuscado
    Dim pwd As String
    
    pwd = InputBox("Ingresa la contrase�a para generar el c�digo ofuscado:" & vbCrLf & vbCrLf & _
                   "Esto es �til si necesitas a�adir contrase�as hardcoded al c�digo.", _
                   "Generador de C�digo Ofuscado", "")
    
    If pwd <> "" Then
        Call ModuloConfigSegura.GenerarCodigoOfuscado(pwd)
    End If
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

Private Sub LimpiarCampos()
    txtNombre.Value = ""
    txtPasswordActual.Value = ""
    txtNuevaPassword.Value = ""
    txtConfirmarPassword.Value = ""
    txtDescripcion.Value = ""
    chkMostrar.Value = False
    lblFortaleza.Caption = "Fortaleza: ---"
    lblFortaleza.ForeColor = RGB(128, 128, 128)
    
    ' Ocultar campo de contrase�a actual
    lblPasswordActual.visible = False
    txtPasswordActual.visible = False
    
    ' Habilitar campos de nueva contrase�a
    txtNuevaPassword.Enabled = True
    txtConfirmarPassword.Enabled = True
    
    btnGuardar.Caption = "Guardar"
End Sub

