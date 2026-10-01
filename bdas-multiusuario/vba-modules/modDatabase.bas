Attribute VB_Name = "modDatabase"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ==============================================================================
' Módulo: modDatabase.bas
' Propósito: Capa de acceso a datos ADO entre Excel Front-End y MS Access
'            Back-End (H:\ResidenciaBD\Residencia_BE.accdb)
' ==============================================================================

Private Const DB_PATH As String = "H:\ResidenciaBD\Residencia_BE.accdb"
Private Const PROVIDER_STRING As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="

' Constantes de estado de conexión ADO
Private Const adStateClosed As Long = 0
Private Const adStateOpen As Long = 1

' ===========================
' ESTADO DEL USUARIO (SESIÓN)
' ===========================
' Variables de sesión establecidas en LoginUsuario().
' Persisten mientras el libro esté abierto en memoria.
Private m_NombreUsuario As String             ' Usuario introducido en el login manual
Private m_NombreCompleto As String            ' Nombre legible de la tabla Usuarios
Private m_Rol As String                       ' "OPERADOR" o "ADMIN"
Private m_ResidenciasAsignadas As Collection  ' Ej: {"GIJON", "SOTO"}
Private m_ResidenciaActiva As String          ' La residencia seleccionada actualmente
Private m_LoginCompletado As Boolean          ' Flag para evitar doble login

Private Const MAX_INTENTOS_LOGIN As Integer = 3  ' Intentos máximos antes de cerrar

''' Realiza una pausa de N milisegundos cediendo control con DoEvents.
Public Sub PausaMS(ByVal milisegundos As Long)
    On Error Resume Next
    If milisegundos > 0 Then
        Sleep milisegundos
    End If
    DoEvents
End Sub

' ===========================
' HASH SHA-256 (para contraseñas)
' ===========================

''' Calcula el hash SHA-256 de un texto usando .NET a través de COM.
''' Retorna el hash en hexadecimal (64 caracteres, minúsculas).
''' Compatible con la misma función Get-SHA256Hash del script PowerShell.
Public Function HashSHA256(ByVal texto As String) As String
    Dim objSHA256 As Object
    Dim bytesTexto() As Byte
    Dim bytesHash() As Byte
    Dim i As Long
    Dim resultado As String
    
    On Error GoTo ErrorHandler
    
    ' Crear instancia de SHA256 de .NET
    Set objSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    ' Convertir texto a bytes UTF-8
    bytesTexto = StrConv(texto, vbFromUnicode)
    
    ' Calcular hash
    bytesHash = objSHA256.ComputeHash_2(bytesTexto)
    
    ' Convertir hash a hexadecimal
    resultado = ""
    For i = LBound(bytesHash) To UBound(bytesHash)
        resultado = resultado & Right("0" & Hex(bytesHash(i)), 2)
    Next i
    
    HashSHA256 = LCase(resultado)
    Exit Function

ErrorHandler:
    ' Fallback: si .NET no está disponible, usar hash simple
    ' (menos seguro, pero funcional)
    HashSHA256 = HashSimple(texto)
End Function

''' Hash simple como fallback si SHA-256 no está disponible.
''' NOTA: Menos seguro, usar solo como respaldo.
Private Function HashSimple(ByVal texto As String) As String
    Dim i As Long
    Dim hash As Long
    hash = 5381
    For i = 1 To Len(texto)
        hash = ((hash * 33) Xor Asc(Mid(texto, i, 1))) And &H7FFFFFFF
    Next i
    HashSimple = "FALLBACK_" & Hex(hash)
End Function

' ===========================
' LOGIN MANUAL
' ===========================

''' Inicia el proceso de login manual mostrando el formulario frmLogin.
''' El formulario llama a ValidarCredenciales() para autenticar.
''' Después del login, si el usuario tiene varias residencias, muestra frmSelectorResidencia.
'''
''' Retorna True si el login fue exitoso, False en caso contrario.
Public Function LoginUsuario() As Boolean
    Dim intentos As Integer
    
    On Error GoTo ErrorHandler
    
    ' Evitar doble login si ya se ejecutó
    If m_LoginCompletado Then
        LoginUsuario = True
        Exit Function
    End If
    
    ' Mostrar formulario de login (modal)
    ' El formulario llama a ValidarCredenciales() y establece m_NombreUsuario
    intentos = 0
    Do
        frmLogin.Show vbModal
        
        ' Si el usuario cerró el formulario sin autenticarse
        If Not m_LoginCompletado Then
            intentos = intentos + 1
            If Len(m_NombreUsuario) = 0 Then
                ' El usuario pulsó Salir/cerró el formulario
                LoginUsuario = False
                Exit Function
            End If
            
            If intentos >= MAX_INTENTOS_LOGIN Then
                MsgBox "Se han agotado los intentos de login (" & MAX_INTENTOS_LOGIN & ")." & vbCrLf & _
                       "El libro se cerrará.", vbCritical, "Acceso Denegado"
                LoginUsuario = False
                Exit Function
            End If
        End If
        Loop Until m_LoginCompletado
    
    ' Descargar completamente el formulario de login para que no reaparezca
    On Error Resume Next
    Unload frmLogin
    On Error GoTo ErrorHandler
    
    ' Login exitoso - ahora seleccionar residencia
    If m_ResidenciasAsignadas.Count = 1 Then
        ' Solo tiene una residencia -> activar directamente
        m_ResidenciaActiva = m_ResidenciasAsignadas(1)
    Else
        ' Tiene varias -> mostrar selector
        frmSelectorResidencia.Show vbModal
        
        ' Descargar completamente el selector de residencia
        On Error Resume Next
        Unload frmSelectorResidencia
        On Error GoTo ErrorHandler
        
        ' Verificar que se seleccionó una
        If Len(m_ResidenciaActiva) = 0 Then
            MsgBox "No se seleccionó ninguna residencia. El libro se cerrará.", _
                   vbExclamation, "Sin Selección"
            m_LoginCompletado = False
            LoginUsuario = False
            Exit Function
        End If
    End If
    
    ' Registrar el login en el log
    InsertarLog m_NombreUsuario, "LOGIN", _
                "Residencia activa: " & m_ResidenciaActiva & _
                " | Rol: " & m_Rol
    
    LoginUsuario = True
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error durante el login:" & vbCrLf & Err.Description, _
           vbCritical, "Error de Login"
    LoginUsuario = False

CleanUp:
    On Error Resume Next
    Unload frmLogin
    Unload frmSelectorResidencia
    On Error GoTo 0
End Function


''' Valida usuario y contraseña contra la tabla Usuarios de Access.
''' Se llama desde el formulario frmLogin cuando el usuario pulsa "Entrar".
'''
''' Retorna True si las credenciales son válidas y el usuario está activo.
Public Function ValidarCredenciales(ByVal usuario As String, ByVal clave As String) As Boolean
    Dim cn As Object
    Dim rs As Object
    Dim sql As String
    Dim arrResidencias() As String
    Dim i As Long
    Dim claveHash As String
    
    On Error GoTo ErrorHandler
    
    ' Calcular hash de la contraseña introducida
    claveHash = HashSHA256(clave)
    
    ' Consultar la tabla Usuarios en Access
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5
    cn.Open GetConnectionString()
    
    sql = "SELECT NombreCompleto, Residencias, Rol, Activo FROM Usuarios " & _
          "WHERE NombreUsuario = '" & EscaparSQL(usuario) & "' " & _
          "AND Clave = '" & EscaparSQL(claveHash) & "'"
    
    Set rs = cn.Execute(sql)
    
    If rs.EOF Then
        rs.Close
        cn.Close
        ValidarCredenciales = False
        GoTo CleanUp
    End If
    
    ' Verificar que está activo
    If Not CBool(rs("Activo").Value) Then
        rs.Close
        cn.Close
        MsgBox "El usuario '" & usuario & "' está desactivado." & vbCrLf & _
               "Contacta con el administrador.", _
               vbCritical, "Usuario Desactivado"
        ValidarCredenciales = False
        GoTo CleanUp
    End If
    
    ' Credenciales válidas - cargar datos del usuario
    m_NombreUsuario = usuario
    m_NombreCompleto = Nz(rs("NombreCompleto").Value, usuario)
    m_Rol = Nz(rs("Rol").Value, "OPERADOR")
    
    ' Parsear residencias asignadas (CSV: "GIJON,SOTO,OVIEDO")
    Set m_ResidenciasAsignadas = New Collection
    arrResidencias = Split(CStr(rs("Residencias").Value), ",")
    For i = LBound(arrResidencias) To UBound(arrResidencias)
        Dim resLimpia As String
        resLimpia = UCase(Trim(arrResidencias(i)))
        If Len(resLimpia) > 0 Then
            m_ResidenciasAsignadas.Add resLimpia
        End If
    Next i
    
    rs.Close
    cn.Close
    
    If m_ResidenciasAsignadas.Count = 0 Then
        MsgBox "El usuario '" & usuario & "' no tiene residencias asignadas." & vbCrLf & _
               "Contacta con el administrador.", _
               vbCritical, "Sin Residencias"
        ValidarCredenciales = False
        GoTo CleanUp
    End If
    
    m_LoginCompletado = True
    ValidarCredenciales = True
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error al validar credenciales:" & vbCrLf & Err.Description, _
           vbCritical, "Error de Login"
    ValidarCredenciales = False

CleanUp:
    If Not rs Is Nothing Then
        On Error Resume Next
        If rs.State = adStateOpen Then rs.Close
        On Error GoTo 0
        Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        On Error Resume Next
        If cn.State = adStateOpen Then cn.Close
        On Error GoTo 0
        Set cn = Nothing
    End If
End Function

''' Permite a un administrador cambiar la contraseña de un usuario.
Public Function CambiarClave(ByVal nombreUsuario As String, ByVal nuevaClave As String) As Boolean
    If m_Rol <> "ADMIN" And m_NombreUsuario <> nombreUsuario Then
        MsgBox "Solo puedes cambiar tu propia contraseña o ser administrador.", vbExclamation
        CambiarClave = False
        Exit Function
    End If
    
    Dim claveHash As String
    claveHash = HashSHA256(nuevaClave)
    
    CambiarClave = ExecuteNonQuery( _
        "UPDATE Usuarios SET Clave = '" & EscaparSQL(claveHash) & "' " & _
        "WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'" _
    )
End Function

''' Retorna el nombre de usuario autenticado.
Public Function ObtenerUsuarioActual() As String
    ObtenerUsuarioActual = m_NombreUsuario
End Function

''' Retorna el nombre completo del usuario logueado.
Public Function ObtenerNombreCompleto() As String
    ObtenerNombreCompleto = m_NombreCompleto
End Function

''' Retorna el rol del usuario ("OPERADOR" o "ADMIN").
Public Function ObtenerRolUsuario() As String
    ObtenerRolUsuario = m_Rol
End Function

' ===========================
' GESTIÓN DE RESIDENCIAS
' ===========================

''' Retorna la colección de residencias asignadas al usuario.
Public Function ObtenerResidenciasAsignadas() As Collection
    Set ObtenerResidenciasAsignadas = m_ResidenciasAsignadas
End Function

''' Retorna la residencia actualmente seleccionada.
Public Function ObtenerResidenciaActiva() As String
    ObtenerResidenciaActiva = m_ResidenciaActiva
End Function

''' Cambia la residencia activa. Solo permite cambiar a una residencia asignada.
''' Retorna True si el cambio fue exitoso.
Public Function CambiarResidenciaActiva(ByVal residencia As String) As Boolean
    residencia = UCase(Trim(residencia))
    
    If TieneAccesoAResidencia(residencia) Then
        m_ResidenciaActiva = residencia
        CambiarResidenciaActiva = True
    Else
        MsgBox "No tienes permiso para acceder a la residencia '" & residencia & "'.", _
               vbExclamation, "Acceso Denegado"
        CambiarResidenciaActiva = False
    End If
End Function

''' Establece la residencia activa directamente (usado por frmSelectorResidencia).
''' No verifica permisos: el formulario solo muestra residencias permitidas.
Public Sub EstablecerResidenciaActiva(ByVal residencia As String)
    m_ResidenciaActiva = UCase(Trim(residencia))
    MostrarHojasResidencia m_ResidenciaActiva
End Sub

''' Verifica si el usuario tiene acceso a una residencia especÃ­fica.
Public Function TieneAccesoAResidencia(ByVal residencia As String) As Boolean
    residencia = UCase(Trim(residencia))
    
    If m_ResidenciasAsignadas Is Nothing Then
        TieneAccesoAResidencia = False
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To m_ResidenciasAsignadas.Count
        If m_ResidenciasAsignadas(i) = residencia Then
            TieneAccesoAResidencia = True
            Exit Function
        End If
    Next i
    
    TieneAccesoAResidencia = False
End Function

''' Retorna el nombre de la hoja Excel correspondiente a una residencia.
Public Function ObtenerNombreHoja(ByVal residencia As String) As String
    Select Case UCase(Trim(residencia))
        Case "GIJON": ObtenerNombreHoja = "BDAS GIJÓN"
        Case "SOTO": ObtenerNombreHoja = "BDAS SOTO"
        Case "OVIEDO": ObtenerNombreHoja = "BDAS OVIEDO"
        Case Else: ObtenerNombreHoja = ""
    End Select
End Function

''' Busca y retorna una hoja de forma segura por nombre, sin lanzar nunca Error 9.
Public Function ObtenerHojaSegura(ByVal nombreHoja As String) As Worksheet
    Dim ws As Worksheet
    Dim buscado As String
    Dim actual As String
    
    Set ObtenerHojaSegura = Nothing
    If Len(Trim(nombreHoja)) = 0 Then Exit Function
    
    ' 1. Búsqueda por coincidencia de texto directa
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nombreHoja, vbTextCompare) = 0 Then
            Set ObtenerHojaSegura = ws
            Exit Function
        End If
    Next ws
    
    ' 2. Búsqueda normalizada (tolerante a tildes o codificación)
    buscado = UCase(Trim(nombreHoja))
    buscado = Replace(buscado, "Ó", "O")
    buscado = Replace(buscado, "Ã", "I")
    buscado = Replace(buscado, "Ã", "A")
    buscado = Replace(buscado, "Ã‰", "E")
    buscado = Replace(buscado, "Ãš", "U")
    buscado = Replace(buscado, "Ã‘", "N")
    buscado = Replace(buscado, " ", "")
    
    For Each ws In ThisWorkbook.Worksheets
        actual = UCase(Trim(ws.Name))
        actual = Replace(actual, "Ó", "O")
        actual = Replace(actual, "Ã", "I")
        actual = Replace(actual, "Ã", "A")
        actual = Replace(actual, "Ã‰", "E")
        actual = Replace(actual, "Ãš", "U")
        actual = Replace(actual, "Ã‘", "N")
        actual = Replace(actual, " ", "")
        
        If actual = buscado Then
            Set ObtenerHojaSegura = ws
            Exit Function
        End If
    Next ws
    
    ' 3. Intento directo por compatibilidad
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    Set ObtenerHojaSegura = ws
End Function

' ===========================
' FUNCIONES BASE
' ===========================

''' Retorna la cadena de conexión ADO OLEDB.
Public Function GetConnectionString() As String
    GetConnectionString = PROVIDER_STRING & DB_PATH & ";"
End Function

''' Sanitiza un texto para uso seguro en sentencias SQL.
''' Escapa comillas simples para prevenir inyección SQL.
Public Function EscaparSQL(ByVal texto As String) As String
    If Len(texto) = 0 Then
        EscaparSQL = ""
    Else
        EscaparSQL = Replace(texto, "'", "''")
    End If
End Function

''' Formatea una fecha para sintaxis Access SQL (#yyyy-mm-dd#).
Public Function FormatearFechaSQL(ByVal fecha As Date) As String
    FormatearFechaSQL = "#" & Format(fecha, "yyyy-mm-dd") & "#"
End Function

''' Formatea una fecha con hora para sintaxis Access SQL.
Public Function FormatearFechaHoraSQL(ByVal fecha As Date) As String
    FormatearFechaHoraSQL = "#" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & "#"
End Function

''' Comprueba si la base de datos Access es accesible.
''' Retorna True si la conexión se abre correctamente, False en caso contrario.
Public Function ComprobarConexion() As Boolean
    Dim cn As Object
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5  ' Timeout de 5 segundos
    cn.Open GetConnectionString()
    cn.Close
    ComprobarConexion = True
    GoTo CleanUp

ErrorHandler:
    ComprobarConexion = False

CleanUp:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

' ===========================
' OPERACIONES GENÃ‰RICAS
' ===========================

''' Ejecuta una consulta SQL de tipo INSERT, UPDATE o DELETE.
''' Patrón de conexión de vida corta: abre, ejecuta y cierra de inmediato.
''' Retorna True si la operación fue exitosa.
Public Function ExecuteNonQuery(ByVal sql As String) As Boolean
    Dim cn As Object
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    cn.Execute sql
    cn.Close
    
    ExecuteNonQuery = True
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Error al ejecutar operación en la base de datos:" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & "Consulta: " & Left(sql, 200), _
           vbCritical, "Error de Base de Datos"
    ExecuteNonQuery = False

CleanUp:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

''' Ejecuta una consulta SELECT que retorna un solo valor (primera columna, primera fila).
''' Ãštil para COUNT(*), MAX(...), SELECT @@IDENTITY, etc.
''' Retorna Null si hay error o no hay resultados.
Public Function ExecuteScalar(ByVal sql As String) As Variant
    Dim cn As Object
    Dim rs As Object
    
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    
    Set rs = cn.Execute(sql)
    If Not rs.EOF Then
        ExecuteScalar = rs(0).Value
    Else
        ExecuteScalar = Null
    End If
    rs.Close
    cn.Close
    GoTo CleanUp
    
ErrorHandler:
    ExecuteScalar = Null

CleanUp:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

''' Ejecuta una consulta SELECT y retorna un Recordset desconectado en memoria.
''' El Recordset se desconecta de la BD para liberar la conexión inmediatamente.
Public Function GetRecordset(ByVal sql As String) As Object
    Dim cn As Object
    Dim rs As Object
    
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    cn.Open GetConnectionString()
    rs.CursorLocation = 3 ' adUseClient
    rs.Open sql, cn, 3, 1 ' adOpenStatic, adLockReadOnly
    
    ' Desconectar el Recordset para liberar la conexión con Access
    Set rs.ActiveConnection = Nothing
    cn.Close
    
    Set GetRecordset = rs
    GoTo CleanUp
    
ErrorHandler:
    MsgBox "Error al consultar la base de datos:" & vbCrLf & _
           Err.Description, vbCritical, "Error de Consulta"
    Set GetRecordset = Nothing

CleanUp:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

' ===========================
' GESTIÓN DE ÓRDENES/SOLICITUDES
' ===========================

''' Inserta un registro de Orden/Reserva de forma atómica y retorna el N ORDEN asignado.
''' El autonumérico de Access garantiza unicidad incluso con 15 usuarios simultáneos.
''' Retorna 0 si hay error.
Public Function InsertarOrdenAtomica( _
    ByVal residencia As String, _
    ByVal fechaPeticion As Date, _
    ByVal dni As String, _
    ByVal nombre As String, _
    ByVal apellidos As String, _
    ByVal tipoHuesped As String, _
    ByVal numHabInd As Integer, _
    ByVal numHabDob As Integer, _
    ByVal fechaEntrada As Date, _
    ByVal fechaSalida As Date, _
    ByVal observaciones As String, _
    Optional ByVal telefono As String = "", _
    Optional ByVal email As String = "", _
    Optional ByVal direccion As String = "", _
    Optional ByVal codigoPostal As String = "", _
    Optional ByVal poblacion As String = "", _
    Optional ByVal provincia As String = "" _
) As Long

    Dim cn As Object
    Dim rs As Object
    Dim sql As String
    Dim nuevoNumOrden As Long
    Dim intentos As Integer
    Dim exito As Boolean
    
    intentos = 5
    exito = False
    
    Do While intentos > 0 And Not exito
        On Error Resume Next
        Err.Clear
        
        Set cn = CreateObject("ADODB.Connection")
        cn.Open GetConnectionString()
        
        If Err.Number = 0 Then
            cn.BeginTrans
            
            Set rs = cn.Execute("SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '" & EscaparSQL(residencia) & "'")
            If Not rs.EOF And Not IsNull(rs(0).Value) Then
                nuevoNumOrden = CLng(rs(0).Value) + 1
            Else
                nuevoNumOrden = 1
            End If
            rs.Close
            
            sql = "INSERT INTO Ordenes (" & _
                  "NumOrden, Residencia, FechaPeticion, DNI, Nombre, Apellidos, TipoHuesped, " & _
                  "NumHabIndividuales, NumHabDobles, FechaEntrada, FechaSalida, " & _
                  "Resolucion, EstadoPago, Observaciones, " & _
                  "Telefono, Email, Direccion, CodigoPostal, Poblacion, Provincia, " & _
                  "Solapamiento, ConsentimientoRGPD, " & _
                  "FechaCreacion, UsuarioCreacion" & _
                  ") VALUES (" & nuevoNumOrden & ", "
                  
            sql = sql & "'" & EscaparSQL(residencia) & "', " & _
                  FormatearFechaSQL(fechaPeticion) & ", " & _
                  "'" & EscaparSQL(dni) & "', " & _
                  "'" & EscaparSQL(nombre) & "', " & _
                  "'" & EscaparSQL(apellidos) & "', " & _
                  "'" & EscaparSQL(tipoHuesped) & "', " & _
                  numHabInd & ", " & numHabDob & ", " & _
                  FormatearFechaSQL(fechaEntrada) & ", " & _
                  FormatearFechaSQL(fechaSalida) & ", "
                  
            sql = sql & "'', '', " & _
                  "'" & EscaparSQL(observaciones) & "', " & _
                  "'" & EscaparSQL(telefono) & "', " & _
                  "'" & EscaparSQL(email) & "', " & _
                  "'" & EscaparSQL(direccion) & "', " & _
                  "'" & EscaparSQL(codigoPostal) & "', " & _
                  "'" & EscaparSQL(poblacion) & "', " & _
                  "'" & EscaparSQL(provincia) & "', " & _
                  "0, 0, " & _
                  FormatearFechaHoraSQL(Now) & ", " & _
                  "'" & EscaparSQL(m_NombreUsuario) & "'" & _
                  ")"
                  
            cn.Execute sql
            
            If Err.Number = 0 Then
                cn.CommitTrans
                exito = True
            Else
                cn.RollbackTrans
            End If
            
            cn.Close
        End If
        
        If Not exito Then
            intentos = intentos - 1
            PausaMS 50
        End If
    Loop
    
    If exito Then
        InsertarOrdenAtomica = nuevoNumOrden
    Else
        MsgBox "No se pudo asignar un Nº ORDEN atómico tras varios intentos concurrentes.", vbCritical, "Error Concurrencia"
        InsertarOrdenAtomica = 0
    End If
    
    GoTo CleanUp

CleanUp:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

''' Actualiza un campo especÃ­fico de una orden existente.
''' Ejemplo: ActualizarOrden 1045, "Resolucion", "CONCEDIDA"
''' Ejemplo: ActualizarOrden 1045, "HabitacionesAsignadas", "Hab.3,Hab.5"
Public Function ActualizarOrden( _
    ByVal numOrden As Long, _
    ByVal campo As String, _
    ByVal valor As String _
) As Boolean
    Dim sql As String
    
    ' Campos de fecha necesitan formato especial
    If LCase(campo) = "fechaentrada" Or LCase(campo) = "fechasalida" Or _
       LCase(campo) = "fechapeticion" Then
        If IsDate(valor) Then
            sql = "UPDATE Ordenes SET " & campo & " = " & FormatearFechaSQL(CDate(valor)) & _
                  " WHERE NumOrden = " & numOrden
        Else
            ActualizarOrden = False
            Exit Function
        End If
    ' Campos numéricos
    ElseIf LCase(campo) = "numfactura" Or LCase(campo) = "numhabindividuales" Or _
           LCase(campo) = "numhabdobles" Then
        sql = "UPDATE Ordenes SET " & campo & " = " & CLng(valor) & _
              " WHERE NumOrden = " & numOrden
    ' Campos booleanos
    ElseIf LCase(campo) = "solapamiento" Or LCase(campo) = "consentimientorgpd" Then
        sql = "UPDATE Ordenes SET " & campo & " = " & IIf(CBool(valor), "True", "False") & _
              " WHERE NumOrden = " & numOrden
    ' Campos de texto
    Else
        sql = "UPDATE Ordenes SET " & campo & " = '" & EscaparSQL(valor) & "'" & _
              " WHERE NumOrden = " & numOrden
    End If
    
    ActualizarOrden = ExecuteNonQuery(sql)
End Function

''' Actualiza múltiples campos de una orden en una sola operación.
''' camposYValores debe ser un diccionario (Scripting.Dictionary) con pares campo->valor.
Public Function ActualizarOrdenMultiple( _
    ByVal numOrden As Long, _
    ByVal camposYValores As Object _
) As Boolean
    Dim sql As String
    Dim setParts As String
    Dim k As Variant
    
    setParts = ""
    For Each k In camposYValores.Keys
        If Len(setParts) > 0 Then setParts = setParts & ", "
        setParts = setParts & CStr(k) & " = '" & EscaparSQL(CStr(camposYValores(k))) & "'"
    Next k
    
    sql = "UPDATE Ordenes SET " & setParts & " WHERE NumOrden = " & numOrden
    ActualizarOrdenMultiple = ExecuteNonQuery(sql)
End Function

' ===========================
' FACTURACIÓN
' ===========================

''' Obtiene el siguiente número de factura de forma atómica para una residencia y ejercicio.
''' Incrementa el contador en Access dentro de una transacción, garantizando unicidad.
''' Retorna 0 si hay error.
Public Function ObtenerSiguienteNumFactura( _
    ByVal residencia As String, _
    ByVal ejercicio As Integer _
) As Long
    Dim cn As Object
    Dim rs As Object
    Dim nuevoNum As Long
    
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    cn.BeginTrans
    
    ' Incrementar el contador atómicamente
    cn.Execute "UPDATE ContadorFacturas SET UltimoNumero = UltimoNumero + 1 " & _
               "WHERE Residencia = '" & EscaparSQL(residencia) & "' AND Ejercicio = " & ejercicio
    
    ' Leer el nuevo valor
    Set rs = cn.Execute("SELECT UltimoNumero FROM ContadorFacturas " & _
                        "WHERE Residencia = '" & EscaparSQL(residencia) & "' AND Ejercicio = " & ejercicio)
    
    If Not rs.EOF Then
        nuevoNum = CLng(rs(0).Value)
    End If
    rs.Close
    
    cn.CommitTrans
    cn.Close
    
    ObtenerSiguienteNumFactura = nuevoNum
    GoTo CleanUp

ErrorHandler:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then
            On Error Resume Next
            cn.RollbackTrans
            On Error GoTo 0
        End If
    End If
    MsgBox "Error al obtener número de factura:" & vbCrLf & Err.Description, _
           vbCritical, "Error de Facturación"
    ObtenerSiguienteNumFactura = 0

CleanUp:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

' ===========================
' BÚSQUEDAS
' ===========================

''' Busca en la tabla Ordenes y en Ordenes_Historico unificando resultados.
''' criterio: "DNI", "NOMBRE", "NUMORDEN", "NUMFACTURA"
''' valor: texto a buscar
''' residencia: opcional, filtrar por residencia ("GIJON", "SOTO", "OVIEDO", "" = todas)
''' incluirHistorico: opcional (True por defecto), busca tanto en activas como en histórico
''' Retorna un Recordset desconectado con los resultados priorizando la solicitud más reciente.
Public Function BuscarEnOrdenes( _
    ByVal criterio As String, _
    ByVal valor As String, _
    Optional ByVal residencia As String = "", _
    Optional ByVal incluirHistorico As Boolean = True _
) As Object
    Dim sql As String
    Dim whereClause As String
    
    Select Case UCase(criterio)
        Case "DNI"
            whereClause = "DNI LIKE '%" & EscaparSQL(Trim(valor)) & "%'"
        Case "NOMBRE"
            whereClause = "(Nombre LIKE '%" & EscaparSQL(Trim(valor)) & "%' OR Apellidos LIKE '%" & EscaparSQL(Trim(valor)) & "%')"
        Case "NUMORDEN"
            whereClause = "NumOrden = " & CLng(valor)
        Case "NUMFACTURA"
            whereClause = "NumFactura = " & CLng(valor)
        Case Else
            whereClause = "1=1"
    End Select
    
    If Len(residencia) > 0 Then
        whereClause = whereClause & " AND Residencia = '" & EscaparSQL(residencia) & "'"
    End If
    
    ' Consulta combinada (Activas + Histórico) ordenando por fecha y orden más reciente
    If incluirHistorico Then
        sql = "SELECT * FROM Ordenes WHERE " & whereClause & " " & _
              "UNION ALL " & _
              "SELECT * FROM Ordenes_Historico WHERE " & whereClause & " " & _
              "ORDER BY FechaPeticion DESC, NumOrden DESC"
    Else
        sql = "SELECT * FROM Ordenes WHERE " & whereClause & " ORDER BY FechaPeticion DESC, NumOrden DESC"
    End If
    
    Set BuscarEnOrdenes = GetRecordset(sql)
End Function


''' Comprueba si ya existe una solicitud duplicada (mismo DNI, misma residencia, mismas fechas).
''' Retorna True si existe duplicado.
Public Function ExisteDuplicado( _
    ByVal dni As String, _
    ByVal residencia As String, _
    ByVal fechaEntrada As Date, _
    ByVal fechaSalida As Date _
) As Boolean
    Dim resultado As Variant
    
    resultado = ExecuteScalar( _
        "SELECT COUNT(*) FROM Ordenes WHERE " & _
        "DNI = '" & EscaparSQL(dni) & "' AND " & _
        "Residencia = '" & EscaparSQL(residencia) & "' AND " & _
        "FechaEntrada = " & FormatearFechaSQL(fechaEntrada) & " AND " & _
        "FechaSalida = " & FormatearFechaSQL(fechaSalida) & " AND " & _
        "Resolucion NOT IN ('DENEGADA', 'RENUNCIA')" _
    )
    
    ExisteDuplicado = (Not IsNull(resultado) And CLng(Nz(resultado, 0)) > 0)
End Function

' ===========================
' LISTA NEGRA
' ===========================

''' Comprueba si un DNI está en la lista negra.
''' Retorna True si el DNI está vetado (activo).
''' Opcionalmente rellena outMotivo con el motivo registrado en Access.
Public Function ComprobarListaNegra(ByVal dni As String, Optional ByRef outMotivo As String = "") As Boolean
    Dim rs As Object
    Dim sql As String
    
    outMotivo = ""
    sql = "SELECT Motivo FROM ListaNegra WHERE DNI = '" & EscaparSQL(dni) & "' AND Activo = True"
    
    Set rs = GetRecordset(sql)
    If Not rs Is Nothing Then
        If Not rs.EOF Then
            outMotivo = CStr(Nz(rs("Motivo").Value, ""))
            ComprobarListaNegra = True
        Else
            ComprobarListaNegra = False
        End If
        rs.Close
        Set rs = Nothing
    Else
        ComprobarListaNegra = False
    End If
End Function

''' Añade un DNI a la lista negra.
Public Function AnadirAListaNegra( _
    ByVal dni As String, _
    ByVal nombre As String, _
    ByVal motivo As String _
) As Boolean
    AnadirAListaNegra = ExecuteNonQuery( _
        "INSERT INTO ListaNegra (DNI, Nombre, Motivo, FechaAlta, Activo) VALUES (" & _
        "'" & EscaparSQL(dni) & "', " & _
        "'" & EscaparSQL(nombre) & "', " & _
        "'" & EscaparSQL(motivo) & "', " & _
        FormatearFechaHoraSQL(Now) & ", True)" _
    )
End Function

' ===========================
' LOG DE ACTIVIDAD
' ===========================

''' Registra una acción en la tabla LogActividad de Access.
''' Se llama desde cualquier punto donde antes se escribÃ­a en la hoja LOG.
Public Function InsertarLog( _
    ByVal usuario As String, _
    ByVal accion As String, _
    Optional ByVal detalle As String = "" _
) As Boolean
    InsertarLog = ExecuteNonQuery( _
        "INSERT INTO LogActividad (FechaHora, Usuario, Accion, Detalle) VALUES (" & _
        FormatearFechaHoraSQL(Now) & ", " & _
        "'" & EscaparSQL(usuario) & "', " & _
        "'" & EscaparSQL(accion) & "', " & _
        "'" & EscaparSQL(detalle) & "')" _
    )
End Function

' ===========================
' SINCRONIZACIÓN ACCESS -> EXCEL (CACHÉ VISUAL)
' ===========================

Private Function ObtenerValorCampo(ByRef rs As Object, ByVal nombreCampo As String, Optional ByVal defaultVal As Variant = Null) As Variant
    Dim fld As Object
    If rs Is Nothing Then
        ObtenerValorCampo = defaultVal
        Exit Function
    End If
    For Each fld In rs.Fields
        If StrComp(fld.Name, nombreCampo, vbTextCompare) = 0 Then
            If IsNull(fld.Value) Then
                ObtenerValorCampo = defaultVal
            Else
                ObtenerValorCampo = fld.Value
            End If
            Exit Function
        End If
    Next fld
    ObtenerValorCampo = defaultVal
End Function

''' Descarga todos los registros de una residencia desde Access y los vuelca
''' en la hoja Excel correspondiente, sobreescribiendo los datos existentes.
'''
''' Las celdas funcionan como CACHÉ VISUAL EN MEMORIA para que las macros
''' de calendario, solapamientos e impresión encuentren datos en las celdas.
''' Estos datos NO se guardan al .xlsm: Access es la fuente de verdad.
'''
''' residencia: "GIJON", "SOTO" o "OVIEDO"
''' hoja: referencia a la hoja destino (ej. Sheets("RESIDENCIA GIJÓN"))
''' filaInicio: fila donde empiezan los datos (normalmente 2, siendo la 1 cabeceras)
Public Sub SincronizarHojaDesdeAccess( _
    ByVal residencia As String, _
    ByVal hoja As Worksheet, _
    Optional ByVal filaInicio As Long = 2 _
)
    Dim rs As Object
    Dim sql As String
    Dim ultimaFila As Long
    
    On Error GoTo ErrorHandler
    
    Application.screenUpdating = False
    Application.enableEvents = False  ' Evitar que Worksheet_Change se dispare durante la carga
    
    Dim estabaProtegida As Boolean
    estabaProtegida = hoja.ProtectContents
    If estabaProtegida Then
        On Error Resume Next
        Dim pwdHojas As String
        pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()
        If Len(pwdHojas) > 0 Then
            hoja.Unprotect password:=pwdHojas
        End If
        If hoja.ProtectContents Then
            hoja.Unprotect password:=""
        End If
        If hoja.ProtectContents Then
            hoja.Unprotect
        End If
        On Error GoTo ErrorHandler
    End If

    sql = "SELECT * FROM Ordenes WHERE Residencia = '" & EscaparSQL(residencia) & "' ORDER BY NumOrden ASC"
    Set rs = GetRecordset(sql)
    
    If rs Is Nothing Then
        GoTo CleanUp
    End If
    
    If rs.EOF Then
        GoTo CleanUp
    End If
    
    ' Limpiar datos antiguos de la hoja (preservar cabeceras en fila 1)
    ultimaFila = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
    If ultimaFila >= filaInicio Then
        hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(ultimaFila, 32)).ClearContents
    End If
    
    Dim isOviedo As Boolean
    Dim vVal As Variant
    isOviedo = (UCase(Trim(residencia)) = "OVIEDO")
    
    Dim maxCols As Long
    If isOviedo Then maxCols = 32 Else maxCols = 28
    
    ' Obtener total de registros
    Dim totalRows As Long
    rs.MoveLast
    totalRows = rs.RecordCount
    rs.MoveFirst
    
    If totalRows > 0 Then
        Dim arrData() As Variant
        ReDim arrData(1 To totalRows, 1 To maxCols)
        
        Dim r As Long
        For r = 1 To totalRows
            arrData(r, 1) = ObtenerValorCampo(rs, "NumOrden")                   ' Col A (Nº ORDEN)
            arrData(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")               ' Col B (FECHA PETICION)
            arrData(r, 3) = ObtenerValorCampo(rs, "NumFactura")                  ' Col C (NÚM FACT)
            arrData(r, 4) = ObtenerValorCampo(rs, "Finalidad")                   ' Col D (FINALIDAD)
            arrData(r, 5) = ObtenerValorCampo(rs, "Empleo")                      ' Col E (EMPLEO)
            arrData(r, 6) = ObtenerValorCampo(rs, "Situacion")                   ' Col F (SITUACION)
            arrData(r, 7) = ObtenerValorCampo(rs, "Evaluacion")                  ' Col G (EVALUACIÓN)
            
            vVal = ObtenerValorCampo(rs, "Comision")
            If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "Turno")
            arrData(r, 8) = vVal                                                 ' Col H (COMISIÓN / TURNO)
            
            arrData(r, 9) = ObtenerValorCampo(rs, "DNI")                         ' Col I (DNI)
            arrData(r, 10) = ObtenerValorCampo(rs, "Rango")                      ' Col J (RANGO)
            arrData(r, 11) = ObtenerValorCampo(rs, "Nombre")                     ' Col K (Nombre)
            arrData(r, 12) = ObtenerValorCampo(rs, "FechaEntrada")               ' Col L (ENTRADA)
            arrData(r, 13) = ObtenerValorCampo(rs, "FechaSalida")                ' Col M (SALIDA)
            arrData(r, 14) = ObtenerValorCampo(rs, "DiasUso")                    ' Col N (DIAS USO)
            arrData(r, 15) = ObtenerValorCampo(rs, "PAX")                        ' Col O (PAX)
            arrData(r, 16) = ObtenerValorCampo(rs, "Resolucion")                 ' Col P (RESOLUCION)
            arrData(r, 17) = ObtenerValorCampo(rs, "NumHabIndividuales")         ' Col Q (HAB. IND. / PRECIO APTO)
            arrData(r, 18) = ObtenerValorCampo(rs, "NumHabDobles")               ' Col R (HAB. DOBLE / SUPLE OCUPAN)
            
            If isOviedo Then
                arrData(r, 19) = ObtenerValorCampo(rs, "CamaSuple")              ' Col S (CAMA SUPLE.)
                arrData(r, 20) = ObtenerValorCampo(rs, "HabitacionesAsignadas")  ' Col T (NÚM HAB.)
                arrData(r, 21) = ObtenerValorCampo(rs, "DtoFamNum")              ' Col U (DTO. FAM. NUM.)
                arrData(r, 22) = ObtenerValorCampo(rs, "Importe")                ' Col V (IMPORTE)
                arrData(r, 23) = ObtenerValorCampo(rs, "Telefono")               ' Col W (TELEFONO)
                arrData(r, 24) = ObtenerValorCampo(rs, "Direccion")              ' Col X (DIRECCIÓN)
                
                vVal = ObtenerValorCampo(rs, "CP")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "CodigoPostal")
                arrData(r, 25) = vVal                                             ' Col Y (CP)
                
                arrData(r, 26) = ObtenerValorCampo(rs, "Poblacion")              ' Col Z (POBLACIÓN)
                arrData(r, 27) = ObtenerValorCampo(rs, "Provincia")              ' Col AA (PROVINCIA)
                arrData(r, 28) = ObtenerValorCampo(rs, "EstadoPago")             ' Col AB (PAGADO)
                
                vVal = ObtenerValorCampo(rs, "FechaGrabacion")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "FechaCreacion")
                arrData(r, 32) = vVal                                             ' Col AF (GRABACIÓN SOLICITUD)
            Else
                arrData(r, 19) = ObtenerValorCampo(rs, "HabitacionesAsignadas")  ' Col S (NÚM HAB. / APTO)
                arrData(r, 20) = ObtenerValorCampo(rs, "DtoFamNum")              ' Col T (DTO. FAM. NUM.)
                arrData(r, 21) = ObtenerValorCampo(rs, "Importe")                ' Col U (IMPORTE)
                arrData(r, 22) = ObtenerValorCampo(rs, "Telefono")               ' Col V (TELEFONO)
                arrData(r, 23) = ObtenerValorCampo(rs, "Direccion")              ' Col W (DIRECCIÓN)
                
                vVal = ObtenerValorCampo(rs, "CP")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "CodigoPostal")
                arrData(r, 24) = vVal                                             ' Col X (CP)
                
                arrData(r, 25) = ObtenerValorCampo(rs, "Poblacion")              ' Col Y (POBLACIÓN)
                arrData(r, 26) = ObtenerValorCampo(rs, "Provincia")              ' Col Z (PROVINCIA)
                arrData(r, 27) = ObtenerValorCampo(rs, "EstadoPago")             ' Col AA (PAGADO)
                
                vVal = ObtenerValorCampo(rs, "FechaGrabacion")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "FechaCreacion")
                arrData(r, 28) = vVal                                             ' Col AB (GRABACIÓN SOLICITUD)
            End If
            
            rs.MoveNext
        Next r
        
        ' Volcado ultra-rápido en una sola llamada de bloque COM
        hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(filaInicio + totalRows - 1, maxCols)).Value = arrData
    End If
    
    rs.Close
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error al sincronizar datos desde Access:" & vbCrLf & _
           Err.Description, vbCritical, "Error de Sincronización"

CleanUp:
    If estabaProtegida Then
        On Error Resume Next
        hoja.Protect password:=ModuloConfigSegura.ObtenerPasswordHojas(), UserInterfaceOnly:=True
        On Error GoTo 0
    End If
    Application.enableEvents = True
    Application.screenUpdating = True
    If Not rs Is Nothing Then Set rs = Nothing
End Sub

''' Sincroniza la hoja de la residencia actualmente seleccionada.
''' Llama a SincronizarHojaDesdeAccess con la residencia activa.
Public Sub SincronizarResidenciaActiva()
    If Len(m_ResidenciaActiva) = 0 Then
        MsgBox "No se ha establecido la residencia activa.", vbExclamation
        Exit Sub
    End If
    
    Dim nombreHoja As String
    nombreHoja = ObtenerNombreHoja(m_ResidenciaActiva)
    If Len(nombreHoja) = 0 Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ObtenerHojaSegura(nombreHoja)
    
    If Not ws Is Nothing Then
        Application.StatusBar = "Sincronizando datos de " & nombreHoja & "..."
        SincronizarHojaDesdeAccess m_ResidenciaActiva, ws
        Application.StatusBar = False
    End If
End Sub

''' Sincroniza TODAS las hojas de las residencias asignadas al usuario.
''' Se llama al arranque (Workbook_Open) para precargar datos.
Public Sub SincronizarTodasMisResidencias()
    If m_ResidenciasAsignadas Is Nothing Then Exit Sub
    
    Dim i As Long
    Dim residencia As String
    Dim nombreHoja As String
    Dim ws As Worksheet
    
    For i = 1 To m_ResidenciasAsignadas.Count
        residencia = m_ResidenciasAsignadas(i)
        nombreHoja = ObtenerNombreHoja(residencia)
        Set ws = Nothing
        
        If Len(nombreHoja) > 0 Then
            Set ws = ObtenerHojaSegura(nombreHoja)
            
            If Not ws Is Nothing Then
                Application.StatusBar = "Sincronizando " & nombreHoja & " (" & i & " de " & m_ResidenciasAsignadas.Count & ")..."
                SincronizarHojaDesdeAccess residencia, ws
            End If
        End If
    Next i
    
    Application.StatusBar = False
End Sub

''' Alias para SincronizarResidenciaActiva().
''' Nombre más intuitivo para usar tras operaciones de escritura.
Public Sub RefrescarCacheVisual()
    SincronizarResidenciaActiva
End Sub

Private Sub DesprotegerEstructuraLibro(ByRef estabaProtegida As Boolean)
    estabaProtegida = ThisWorkbook.ProtectStructure
    If estabaProtegida Then
        On Error Resume Next
        Dim pwd As String
        pwd = ModuloConfigSegura.ObtenerPasswordHojasYEstructura()
        If Len(pwd) > 0 Then
            ThisWorkbook.Unprotect password:=pwd
        End If
        If ThisWorkbook.ProtectStructure Then
            pwd = ModuloConfigSegura.ObtenerPasswordHojas()
            If Len(pwd) > 0 Then ThisWorkbook.Unprotect password:=pwd
        End If
        If ThisWorkbook.ProtectStructure Then
            ThisWorkbook.Unprotect password:=""
        End If
        If ThisWorkbook.ProtectStructure Then
            ThisWorkbook.Unprotect
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub ReprotegerEstructuraLibro(ByVal estabaProtegida As Boolean)
    If estabaProtegida Then
        On Error Resume Next
        Dim pwd As String
        pwd = ModuloConfigSegura.ObtenerPasswordHojasYEstructura()
        If Len(pwd) = 0 Then pwd = ModuloConfigSegura.ObtenerPasswordHojas()
        
        If Len(pwd) > 0 Then
            ThisWorkbook.Protect password:=pwd, Structure:=True, Windows:=False
        Else
            ThisWorkbook.Protect Structure:=True, Windows:=False
        End If
        On Error GoTo 0
    End If
End Sub

''' Muestra las hojas RESIDENCIA, RESUMEN y Calendario de las residencias permitidas al usuario,
''' activa la hoja principal de la residencia seleccionada y oculta la hoja INICIO.
Public Sub MostrarHojasResidencia(Optional ByVal residenciaEspecifica As String = "")
    Dim i As Long
    Dim res As String
    Dim wsResidencia As Worksheet, wsResumen As Worksheet, wsCalendario As Worksheet
    Dim wsActivar As Worksheet
    Dim resSeleccionada As String
    Dim libroProtegida As Boolean
    
    resSeleccionada = UCase(Trim(residenciaEspecifica))
    If Len(resSeleccionada) = 0 Then resSeleccionada = UCase(Trim(m_ResidenciaActiva))
    
    On Error Resume Next
    Application.screenUpdating = False
    
    ' Desproteger la estructura del libro si está protegida (necesario para cambiar Visible de hojas)
    DesprotegerEstructuraLibro libroProtegida
    
    Dim colResidencias As Object
    Set colResidencias = m_ResidenciasAsignadas
    
    If colResidencias Is Nothing Then
        Set colResidencias = CreateObject("System.Collections.ArrayList")
        If Len(resSeleccionada) > 0 Then colResidencias.Add resSeleccionada
    ElseIf colResidencias.Count = 0 Then
        If Len(resSeleccionada) > 0 Then colResidencias.Add resSeleccionada
    End If
    
    For i = 1 To colResidencias.Count
        res = UCase(Trim(colResidencias(i)))
        Set wsResidencia = Nothing
        Set wsResumen = Nothing
        Set wsCalendario = Nothing
        
        Select Case res
            Case "GIJON"
                Set wsResidencia = ObtenerHojaSegura("RESIDENCIA GIJÓN")
                Set wsResumen = ObtenerHojaSegura("RESUMEN GIJÓN")
                Set wsCalendario = ObtenerHojaSegura("Calendario GIJÓN")
            Case "SOTO"
                Set wsResidencia = ObtenerHojaSegura("RESIDENCIA SOTO")
                Set wsResumen = ObtenerHojaSegura("RESUMEN SOTO")
                Set wsCalendario = ObtenerHojaSegura("Calendario SOTO")
            Case "OVIEDO"
                Set wsResidencia = ObtenerHojaSegura("RESIDENCIA OVIEDO")
                Set wsResumen = ObtenerHojaSegura("RESUMEN OVIEDO")
                Set wsCalendario = ObtenerHojaSegura("Calendario OVIEDO")
        End Select
        
        If Not wsResidencia Is Nothing Then wsResidencia.visible = xlSheetVisible
        If Not wsResumen Is Nothing Then wsResumen.visible = xlSheetVisible
        If Not wsCalendario Is Nothing Then wsCalendario.visible = xlSheetVisible
        
        If res = resSeleccionada And Not wsResidencia Is Nothing Then
            Set wsActivar = wsResidencia
        End If
    Next i
    
    If Not wsActivar Is Nothing Then
        wsActivar.Activate
    Else
        Select Case resSeleccionada
            Case "GIJON":  Set wsActivar = ObtenerHojaSegura("RESIDENCIA GIJÓN")
            Case "SOTO":   Set wsActivar = ObtenerHojaSegura("RESIDENCIA SOTO")
            Case "OVIEDO": Set wsActivar = ObtenerHojaSegura("RESIDENCIA OVIEDO")
        End Select
        If Not wsActivar Is Nothing Then
            wsActivar.visible = xlSheetVisible
            wsActivar.Activate
        End If
    End If
    
    Dim wsInicio As Worksheet
    Set wsInicio = ObtenerHojaSegura("INICIO")
    If Not wsInicio Is Nothing Then
        wsInicio.visible = xlSheetHidden
    End If
    
    ' Reproteger la estructura del libro si estaba protegida
    ReprotegerEstructuraLibro libroProtegida
    
    Application.screenUpdating = True
    On Error GoTo 0
End Sub

''' Activa (muestra) la hoja Excel de la residencia actualmente seleccionada.
Public Sub ActivarHojaResidenciaActiva()
    MostrarHojasResidencia m_ResidenciaActiva
End Sub

' ===========================
' GESTIÓN DE USUARIOS (ADMIN)
' ===========================

''' Añade un nuevo usuario al sistema. Solo para administradores.
''' residencias: cadena CSV, ej. "GIJON,SOTO"
Public Function AltaUsuario( _
    ByVal nombreUsuario As String, _
    ByVal nombreCompleto As String, _
    ByVal residencias As String, _
    Optional ByVal rol As String = "OPERADOR" _
) As Boolean
    If m_Rol <> "ADMIN" Then
        MsgBox "Solo los administradores pueden dar de alta usuarios.", vbExclamation
        AltaUsuario = False
        Exit Function
    End If
    
    AltaUsuario = ExecuteNonQuery( _
        "INSERT INTO Usuarios (NombreUsuario, NombreCompleto, Residencias, Rol, Activo, FechaAlta) VALUES (" & _
        "'" & EscaparSQL(nombreUsuario) & "', " & _
        "'" & EscaparSQL(nombreCompleto) & "', " & _
        "'" & EscaparSQL(UCase(residencias)) & "', " & _
        "'" & EscaparSQL(UCase(rol)) & "', " & _
        "True, " & FormatearFechaHoraSQL(Now) & ")" _
    )
End Function

''' Desactiva un usuario del sistema. Solo para administradores.
Public Function BajaUsuario(ByVal nombreUsuario As String) As Boolean
    If m_Rol <> "ADMIN" Then
        MsgBox "Solo los administradores pueden dar de baja usuarios.", vbExclamation
        BajaUsuario = False
        Exit Function
    End If
    
    BajaUsuario = ExecuteNonQuery( _
        "UPDATE Usuarios SET Activo = False WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'" _
    )
End Function

' ===========================
' UTILIDADES
' ===========================

''' Función auxiliar para manejar valores Null de Access.
''' Si el valor es Null, retorna el valor por defecto especificado.
Private Function Nz(ByVal valor As Variant, Optional ByVal valorDefecto As Variant = "") As Variant
    If IsNull(valor) Then
        Nz = valorDefecto
    Else
        Nz = valor
    End If
End Function

' ==============================================================================
' ELIMINACIÓN DE ÓRDENES Y AUDITORÍA
' ==============================================================================

''' Obtiene la clave de residencia normalizada ("GIJON", "SOTO", "OVIEDO")
''' a partir del nombre de una hoja (ej. "RESIDENCIA GIJÓN" -> "GIJON").
Public Function ObtenerClaveResidenciaDesdeHoja(ByVal nombreHoja As String) As String
    Dim n As String
    n = UCase(Trim(nombreHoja))
    If InStr(n, "GIJ") > 0 Then
        ObtenerClaveResidenciaDesdeHoja = "GIJON"
    ElseIf InStr(n, "SOTO") > 0 Then
        ObtenerClaveResidenciaDesdeHoja = "SOTO"
    ElseIf InStr(n, "OVIEDO") > 0 Then
        ObtenerClaveResidenciaDesdeHoja = "OVIEDO"
    Else
        ObtenerClaveResidenciaDesdeHoja = n
    End If
End Function

''' Copia la orden completa a Ordenes_Historico, la elimina de la tabla activa Ordenes
''' y registra la acción en LogActividad. Retorna True si la operación fue exitosa.
Public Function EliminarOrdenBD( _
    ByVal numOrden As Long, _
    ByVal residencia As String _
) As Boolean
    Dim sqlCopia As String
    Dim sqlDelete As String
    Dim exito As Boolean
    Dim claveRes As String
    
    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    If claveRes = "" Then claveRes = UCase(Trim(residencia))
    
    ' 1. Asegurar limpieza de copia previa en histórico si existiera (idempotencia)
    Call ExecuteNonQuery("DELETE FROM Ordenes_Historico WHERE NumOrden = " & numOrden & _
                         " AND Residencia = '" & EscaparSQL(claveRes) & "'")
    
    ' 2. Copiar registro completo a Ordenes_Historico
    sqlCopia = "INSERT INTO Ordenes_Historico SELECT * FROM Ordenes WHERE NumOrden = " & numOrden & _
               " AND Residencia = '" & EscaparSQL(claveRes) & "'"
    Call ExecuteNonQuery(sqlCopia)
    
    ' 3. Eliminar de la tabla activa Ordenes
    sqlDelete = "DELETE FROM Ordenes WHERE NumOrden = " & numOrden & _
                " AND Residencia = '" & EscaparSQL(claveRes) & "'"
    exito = ExecuteNonQuery(sqlDelete)
    
    ' 4. Registrar auditoría completa
    If exito Then
        Dim usr As String
        usr = m_NombreUsuario
        If Len(usr) = 0 Then usr = Environ("USERNAME")
        
        Call InsertarLog(usr, "BORRADO_ORDEN", _
            "Archivado en Histórico y retirado Nº ORDEN " & numOrden & " de " & claveRes)
    End If
    
    EliminarOrdenBD = exito
End Function


