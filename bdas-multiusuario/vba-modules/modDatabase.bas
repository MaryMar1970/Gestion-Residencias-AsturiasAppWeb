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
        MostrarHojasResidencia m_ResidenciaActiva
        InicioGestionResidencias.RefrescarRibbon
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
    
    ' Registrar el login en el log central (Access DB)
    InsertarLog m_NombreUsuario, "LOGIN", _
                "Residencia activa: " & m_ResidenciaActiva & _
                " | Rol: " & m_Rol
    
    ' Registrar el login en la(s) hoja(s) LOG de Excel según la residencia activa seleccionada
    ModuloLOG.RegistrarLoginLOG m_NombreUsuario, m_ResidenciaActiva
    
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
        MostrarHojasResidencia m_ResidenciaActiva
        InsertarLog m_NombreUsuario, "CAMBIO_RESIDENCIA", "Nueva residencia activa: " & m_ResidenciaActiva
        ModuloLOG.RegistrarLoginLOG m_NombreUsuario, m_ResidenciaActiva
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
    InicioGestionResidencias.RefrescarRibbon
End Sub


''' Verifica si el usuario tiene acceso a una residencia específica.
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
        Case "GIJON": ObtenerNombreHoja = "RESIDENCIA GIJÓN"
        Case "SOTO": ObtenerNombreHoja = "RESIDENCIA SOTO"
        Case "OVIEDO": ObtenerNombreHoja = "RESIDENCIA OVIEDO"
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
    buscado = Replace(buscado, "Á", "A")
    buscado = Replace(buscado, "É", "E")
    buscado = Replace(buscado, "Í", "I")
    buscado = Replace(buscado, "Ó", "O")
    buscado = Replace(buscado, "Ú", "U")
    buscado = Replace(buscado, "Ü", "U")
    buscado = Replace(buscado, "Ñ", "N")
    buscado = Replace(buscado, "ÃÂ", "A")
    buscado = Replace(buscado, "ÃÂ", "E")
    buscado = Replace(buscado, "ÃÂ", "I")
    buscado = Replace(buscado, "ÃÂ", "O")
    buscado = Replace(buscado, "ÃÂ", "U")
    buscado = Replace(buscado, "ÃÂ", "N")
    buscado = Replace(buscado, " ", "")
    
    For Each ws In ThisWorkbook.Worksheets
        actual = UCase(Trim(ws.Name))
        actual = Replace(actual, "Á", "A")
        actual = Replace(actual, "É", "E")
        actual = Replace(actual, "Í", "I")
        actual = Replace(actual, "Ó", "O")
        actual = Replace(actual, "Ú", "U")
        actual = Replace(actual, "Ü", "U")
        actual = Replace(actual, "Ñ", "N")
        actual = Replace(actual, "ÃÂ", "A")
        actual = Replace(actual, "ÃÂ", "E")
        actual = Replace(actual, "ÃÂ", "I")
        actual = Replace(actual, "ÃÂ", "O")
        actual = Replace(actual, "ÃÂ", "U")
        actual = Replace(actual, "ÃÂ", "N")
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
' OPERACIONES GENÃÂRICAS
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
''' ÃÂtil para COUNT(*), MAX(...), SELECT @@IDENTITY, etc.
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
        MsgBox "No se pudo asignar un NÂº ORDEN atómico tras varios intentos concurrentes.", vbCritical, "Error Concurrencia"
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

''' Actualiza un campo específico de una orden existente.
''' Ejemplo: ActualizarOrden 1045, "Resolucion", "CONCEDIDA"
''' Ejemplo: ActualizarOrden 1045, "HabitacionesAsignadas", "Hab.3,Hab.5"
Public Function ActualizarOrden( _
    ByVal numOrden As Long, _
    ByVal campo As String, _
    ByVal valor As String, _
    Optional ByVal residencia As String = "" _
) As Boolean
    Dim sql As String
    Dim whereClause As String
    Dim resTarget As String
    
    resTarget = Trim$(residencia)
    
    ' 1. Normalizar si se suministró explícitamente
    If Len(resTarget) > 0 Then resTarget = ObtenerClaveResidenciaDesdeHoja(resTarget)
    
    ' 2. Si no se especificó, inferir de la hoja activa en Excel (inmune a multirresidencia)
    If Len(resTarget) = 0 Or (resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO") Then
        On Error Resume Next
        If Not Application.ActiveSheet Is Nothing Then
            resTarget = ObtenerClaveResidenciaDesdeHoja(Application.ActiveSheet.Name)
        End If
        On Error GoTo 0
    End If
    
    ' 3. Si aún no se determina, usar m_ResidenciaActiva ÚNICAMENTE si es univalor (sin comas)
    If Len(resTarget) = 0 Or (resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO") Then
        If Len(m_ResidenciaActiva) > 0 And InStr(m_ResidenciaActiva, ",") = 0 Then
            resTarget = ObtenerClaveResidenciaDesdeHoja(m_ResidenciaActiva)
        End If
    End If
    
    ' 4. Construir cláusula WHERE blindada solo si la residencia es válida e inequívoca
    If resTarget = "GIJON" Or resTarget = "SOTO" Or resTarget = "OVIEDO" Then
        whereClause = " WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "'"
    Else
        whereClause = " WHERE NumOrden = " & numOrden
    End If

    
    ' Campos de fecha necesitan formato especial
    If LCase(campo) = "fechaentrada" Or LCase(campo) = "fechasalida" Or _
       LCase(campo) = "fechapeticion" Then
        If IsDate(valor) Then
            sql = "UPDATE Ordenes SET " & campo & " = " & FormatearFechaSQL(CDate(valor)) & whereClause
        Else
            ActualizarOrden = False
            Exit Function
        End If
    ' Campos numéricos (gestión segura de vacíos como Null)
    ElseIf LCase(campo) = "numfactura" Or LCase(campo) = "numhabindividuales" Or _
           LCase(campo) = "numhabdobles" Then
        If Len(Trim$(valor)) = 0 Or Not IsNumeric(valor) Then
            sql = "UPDATE Ordenes SET " & campo & " = Null" & whereClause
        Else
            sql = "UPDATE Ordenes SET " & campo & " = " & CLng(valor) & whereClause
        End If

    ' Campos booleanos
    ElseIf LCase(campo) = "solapamiento" Or LCase(campo) = "consentimientorgpd" Then
        sql = "UPDATE Ordenes SET " & campo & " = " & IIf(CBool(valor), "True", "False") & whereClause
    ' Campos de texto
    Else
        sql = "UPDATE Ordenes SET " & campo & " = '" & EscaparSQL(valor) & "'" & whereClause
    End If
    
    ActualizarOrden = ExecuteNonQuery(sql)
End Function

''' Actualiza múltiples campos de una orden en una sola operación con blindaje por Residencia.
''' camposYValores debe ser un diccionario (Scripting.Dictionary) con pares campo->valor.
Public Function ActualizarOrdenMultiple( _
    ByVal numOrden As Long, _
    ByVal camposYValores As Object, _
    Optional ByVal residencia As String = "" _
) As Boolean
    Dim sql As String
    Dim setParts As String
    Dim k As Variant
    Dim whereClause As String
    Dim resTarget As String
    
    resTarget = Trim$(residencia)
    
    ' 1. Normalizar si se suministró explícitamente
    If Len(resTarget) > 0 Then resTarget = ObtenerClaveResidenciaDesdeHoja(resTarget)
    
    ' 2. Si no se especificó, inferir de la hoja activa en Excel (inmune a multirresidencia)
    If Len(resTarget) = 0 Or (resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO") Then
        On Error Resume Next
        If Not Application.ActiveSheet Is Nothing Then
            resTarget = ObtenerClaveResidenciaDesdeHoja(Application.ActiveSheet.Name)
        End If
        On Error GoTo 0
    End If
    
    ' 3. Si aún no se determina, usar m_ResidenciaActiva ÚNICAMENTE si es univalor (sin comas)
    If Len(resTarget) = 0 Or (resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO") Then
        If Len(m_ResidenciaActiva) > 0 And InStr(m_ResidenciaActiva, ",") = 0 Then
            resTarget = ObtenerClaveResidenciaDesdeHoja(m_ResidenciaActiva)
        End If
    End If
    
    ' 4. Construir cláusula WHERE blindada solo si la residencia es válida e inequívoca
    If resTarget = "GIJON" Or resTarget = "SOTO" Or resTarget = "OVIEDO" Then
        whereClause = " WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "'"
    Else
        whereClause = " WHERE NumOrden = " & numOrden
    End If

    
    setParts = ""
    For Each k In camposYValores.Keys
        If Len(setParts) > 0 Then setParts = setParts & ", "
        setParts = setParts & CStr(k) & " = '" & EscaparSQL(CStr(camposYValores(k))) & "'"
    Next k
    
    sql = "UPDATE Ordenes SET " & setParts & whereClause
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

' ==============================================================================
' COMPROBACIÓN Y REFRESCO DE FACTURAS EN ACCESOS MULTIUSUARIO
' ==============================================================================

''' Comprueba en Access si un número de factura ya está asignado a otra orden de la misma residencia.
''' Retorna True si ya existe (duplicado), False si está disponible.
Public Function ExisteFacturaEnBD( _
    ByVal residencia As String, _
    ByVal numFactura As Long, _
    Optional ByVal excluirNumOrden As Long = 0) As Boolean
    
    Dim sql As String
    Dim resCount As Variant
    Dim claveRes As String
    
    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
    
    sql = "SELECT COUNT(*) FROM Ordenes " & _
          "WHERE Residencia = '" & EscaparSQL(claveRes) & "' " & _
          "AND NumFactura = " & numFactura
    
    If excluirNumOrden > 0 Then
        sql = sql & " AND NumOrden <> " & excluirNumOrden
    End If
    
    resCount = ExecuteScalar(sql)
    
    If IsNull(resCount) Then
        ExisteFacturaEnBD = False
    Else
        ExisteFacturaEnBD = (CLng(resCount) > 0)
    End If
End Function

''' Refresca de forma ultraligera únicamente la columna C (NÂº FACTURA) desde Access
''' comparando por NÂº ORDEN, sin alterar fórmulas ni el resto de columnas.
Public Sub RefrescarColumnaFacturasDesdeAccess(ws As Worksheet, ByVal residencia As String)
    Dim cn As Object
    Dim rs As Object
    Dim sql As String
    Dim dictFacturas As Object
    Dim ultimaFila As Long
    Dim r As Long
    Dim numOrd As Variant
    Dim arrA As Variant
    Dim arrC As Variant
    Dim huboCambios As Boolean
    Dim claveRes As String
    
    On Error GoTo ErrorHandler
    
    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
    
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then Exit Sub
    
    ' 1. Consultar únicamente órdenes con factura de esta residencia
    sql = "SELECT NumOrden, NumFactura FROM Ordenes " & _
          "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumFactura > 0"
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    Set rs = cn.Execute(sql)
    
    Set dictFacturas = CreateObject("Scripting.Dictionary")
    Do While Not rs.EOF
        If Not IsNull(rs("NumOrden").Value) And Not IsNull(rs("NumFactura").Value) Then
            dictFacturas(CLng(rs("NumOrden").Value)) = CLng(rs("NumFactura").Value)
        End If
        rs.MoveNext
    Loop
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
    ' 2. Comparación en memoria
    If ultimaFila = 2 Then
        numOrd = ws.Cells(2, "A").Value
        If IsNumeric(numOrd) And CLng(numOrd) > 0 Then
            If dictFacturas.Exists(CLng(numOrd)) Then
                If ws.Cells(2, "C").Value <> dictFacturas(CLng(numOrd)) Then
                    ws.Cells(2, "C").Value = dictFacturas(CLng(numOrd))
                End If
            End If
        End If
        Exit Sub
    End If
    
    arrA = ws.Range("A2:A" & ultimaFila).Value
    arrC = ws.Range("C2:C" & ultimaFila).Value
    huboCambios = False
    
    For r = 1 To UBound(arrA, 1)
        numOrd = arrA(r, 1)
        If IsNumeric(numOrd) And Not IsEmpty(numOrd) Then
            If CLng(numOrd) > 0 Then
                If dictFacturas.Exists(CLng(numOrd)) Then
                    If arrC(r, 1) <> dictFacturas(CLng(numOrd)) Then
                        arrC(r, 1) = dictFacturas(CLng(numOrd))
                        huboCambios = True
                    End If
                End If
            End If
        End If
    Next r
    
    ' 3. Volcado en bloque solo si hubo cambios
    If huboCambios Then
        Application.enableEvents = False
        ws.Range("C2:C" & ultimaFila).Value = arrC
        Application.enableEvents = True
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Sub

' ===========================
' BÃSQUEDAS
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
''' Se llama desde cualquier punto donde antes se escribía en la hoja LOG.
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
' SINCRONIZACIÓN ACCESS -> EXCEL (CACHÃ VISUAL)
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
''' Las celdas funcionan como CACHÃ VISUAL EN MEMORIA para que las macros
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

    sql = "SELECT * FROM Ordenes WHERE Residencia = '" & EscaparSQL(residencia) & "' ORDER BY Id ASC"
    Set rs = GetRecordset(sql)
    
    If rs Is Nothing Then
        GoTo CleanUp
    End If
    
    If rs.EOF Then
        GoTo CleanUp
    End If
    
    Dim resNorm As String
    resNorm = UCase(Trim(residencia))
    
    Dim isOviedo As Boolean, isSoto As Boolean, isGijon As Boolean
    isOviedo = (resNorm = "OVIEDO")
    isSoto = (resNorm = "SOTO")
    isGijon = (resNorm = "GIJON" Or resNorm = "GIJÓN")
    
    ' Limpiar datos antiguos de la hoja preservando estrictamente las columnas de fórmulas
    ultimaFila = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
    If ultimaFila >= filaInicio Then
        If isOviedo Then
            ' OVIEDO: Limpiar solo columnas de datos A:U (1:21) y W:AB (23:28).
            ' Fórmulas blindadas: V (22, IMPORTE), AF (32, NOMBRE PDF), AG (33, TEXTO EMAIL).
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(ultimaFila, 21)).ClearContents    ' Cols A a U
            hoja.Range(hoja.Cells(filaInicio, 23), hoja.Cells(ultimaFila, 28)).ClearContents   ' Cols W a AB
        ElseIf isSoto Then
            ' SOTO: Limpiar solo columnas de datos A:P (1:16), S:T (19:20) y V:AA (22:27).
            ' Fórmulas blindadas: Q (17, PRECIO APTO), R (18, SUPLE OCUPAN), U (21, IMPORTE), AB (28, NOMBRE PDF), AC (29, TEXTO EMAIL).
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(ultimaFila, 16)).ClearContents    ' Cols A a P
            hoja.Range(hoja.Cells(filaInicio, 19), hoja.Cells(ultimaFila, 20)).ClearContents   ' Cols S a T
            hoja.Range(hoja.Cells(filaInicio, 22), hoja.Cells(ultimaFila, 27)).ClearContents   ' Cols V a AA
        Else
            ' GIJÓN: Limpiar solo columnas de datos A:T (1:20) y V:AA (22:27).
            ' Fórmulas blindadas: U (21, IMPORTE), AB (28, NOMBRE PDF), AC (29, TEXTO EMAIL).
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(ultimaFila, 20)).ClearContents    ' Cols A a T
            hoja.Range(hoja.Cells(filaInicio, 22), hoja.Cells(ultimaFila, 27)).ClearContents   ' Cols V a AA
        End If
    End If
    
    Dim vVal As Variant
    
    ' Obtener total de registros
    Dim totalRows As Long
    rs.MoveLast
    totalRows = rs.RecordCount
    rs.MoveFirst
    
    If totalRows > 0 Then
        Dim r As Long
        
        If isOviedo Then
            ' OVIEDO: Bloque 1 (Cols A a U = 1 a 21), Bloque 2 (Cols W a AB = 23 a 28)
            ' Col V (22, IMPORTE), Col AF (32, FÓRMULA PDF) y Col AG (33, FÓRMULA EMAIL) se mantienen intactas
            Dim arrOviedo1() As Variant
            Dim arrOviedo2() As Variant
            ReDim arrOviedo1(1 To totalRows, 1 To 21)
            ReDim arrOviedo2(1 To totalRows, 1 To 6)
            
            For r = 1 To totalRows
                arrOviedo1(r, 1) = ObtenerValorCampo(rs, "NumOrden")                   ' Col A (Nº ORDEN)
                arrOviedo1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")               ' Col B (FECHA PETICION)
                arrOviedo1(r, 3) = ObtenerValorCampo(rs, "NumFactura")                  ' Col C (NÚM FACT)
                arrOviedo1(r, 4) = ObtenerValorCampo(rs, "Finalidad")                   ' Col D (FINALIDAD)
                arrOviedo1(r, 5) = ObtenerValorCampo(rs, "Empleo")                      ' Col E (EMPLEO)
                arrOviedo1(r, 6) = ObtenerValorCampo(rs, "Situacion")                   ' Col F (SITUACION)
                arrOviedo1(r, 7) = ObtenerValorCampo(rs, "Evaluacion")                  ' Col G (EVALUACIÓN)
                
                vVal = ObtenerValorCampo(rs, "Comision")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "Turno")
                arrOviedo1(r, 8) = vVal                                                 ' Col H (COMISIÓN / TURNO)
                
                arrOviedo1(r, 9) = ObtenerValorCampo(rs, "DNI")                         ' Col I (DNI)
                arrOviedo1(r, 10) = ObtenerValorCampo(rs, "Rango")                      ' Col J (RANGO)
                arrOviedo1(r, 11) = ObtenerValorCampo(rs, "Nombre")                     ' Col K (Nombre)
                arrOviedo1(r, 12) = ObtenerValorCampo(rs, "FechaEntrada")               ' Col L (ENTRADA)
                arrOviedo1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")                ' Col M (SALIDA)
                arrOviedo1(r, 14) = ObtenerValorCampo(rs, "DiasUso")                    ' Col N (DIAS USO)
                arrOviedo1(r, 15) = ObtenerValorCampo(rs, "PAX")                        ' Col O (PAX)
                arrOviedo1(r, 16) = ObtenerValorCampo(rs, "Resolucion")                 ' Col P (RESOLUCION)
                arrOviedo1(r, 17) = ObtenerValorCampo(rs, "NumHabIndividuales")         ' Col Q (HAB. IND.)
                arrOviedo1(r, 18) = ObtenerValorCampo(rs, "NumHabDobles")               ' Col R (HAB. DOBLE)
                arrOviedo1(r, 19) = ObtenerValorCampo(rs, "CamaSuple")                  ' Col S (CAMA SUPLE.)
                arrOviedo1(r, 20) = ObtenerValorCampo(rs, "HabitacionesAsignadas")      ' Col T (NÚM HAB.)
                arrOviedo1(r, 21) = ObtenerValorCampo(rs, "DtoFamNum")                  ' Col U (DTO. FAM. NUM.)
                
                ' Col V (22): IMPORTE (FÓRMULA PRESERVADA INTACTA)
                
                arrOviedo2(r, 1) = ObtenerValorCampo(rs, "Telefono")                    ' Col W (TELEFONO)
                arrOviedo2(r, 2) = ObtenerValorCampo(rs, "Direccion")                   ' Col X (DIRECCIÓN)
                
                vVal = ObtenerValorCampo(rs, "CP")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "CodigoPostal")
                arrOviedo2(r, 3) = vVal                                                 ' Col Y (CP)
                
                arrOviedo2(r, 4) = ObtenerValorCampo(rs, "Poblacion")                   ' Col Z (POBLACIÓN)
                arrOviedo2(r, 5) = ObtenerValorCampo(rs, "Provincia")                   ' Col AA (PROVINCIA)
                arrOviedo2(r, 6) = ObtenerValorCampo(rs, "EstadoPago")                  ' Col AB (PAGADO)
                
                ' Col AF (32): FÓRMULA NOMBRE PDF (PRESERVADA INTACTA)
                ' Col AG (33): FÓRMULA COPIAR TEXTO EMAIL (PRESERVADA INTACTA)
                
                rs.MoveNext
            Next r
            
            ' Volcado ultra-rápido en bloques COM preservando columnas de fórmulas
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(filaInicio + totalRows - 1, 21)).Value = arrOviedo1
            hoja.Range(hoja.Cells(filaInicio, 23), hoja.Cells(filaInicio + totalRows - 1, 28)).Value = arrOviedo2
            
        ElseIf isSoto Then
            ' SOTO: Bloque 1 (Cols A a P = 1 a 16), Bloque 2 (Cols S a T = 19 a 20), Bloque 3 (Cols V a AA = 22 a 27)
            ' Col Q (17, PRECIO APTO), Col R (18, SUPLE OCUPAN), Col U (21, IMPORTE),
            ' Col AB (28, NOMBRE PDF) y Col AC (29, TEXTO EMAIL) se mantienen 100% INTACTAS con sus fórmulas
            Dim arrSoto1() As Variant
            Dim arrSoto2() As Variant
            Dim arrSoto3() As Variant
            ReDim arrSoto1(1 To totalRows, 1 To 16)
            ReDim arrSoto2(1 To totalRows, 1 To 2)
            ReDim arrSoto3(1 To totalRows, 1 To 6)
            
            For r = 1 To totalRows
                arrSoto1(r, 1) = ObtenerValorCampo(rs, "NumOrden")                   ' Col A (Nº ORDEN)
                arrSoto1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")               ' Col B (FECHA PETICION)
                arrSoto1(r, 3) = ObtenerValorCampo(rs, "NumFactura")                  ' Col C (NÚM FACT)
                arrSoto1(r, 4) = ObtenerValorCampo(rs, "Finalidad")                   ' Col D (FINALIDAD)
                arrSoto1(r, 5) = ObtenerValorCampo(rs, "Empleo")                      ' Col E (EMPLEO)
                arrSoto1(r, 6) = ObtenerValorCampo(rs, "Situacion")                   ' Col F (SITUACION)
                arrSoto1(r, 7) = ObtenerValorCampo(rs, "Evaluacion")                  ' Col G (EVALUACIÓN)
                
                vVal = ObtenerValorCampo(rs, "Turno")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "Comision")
                arrSoto1(r, 8) = vVal                                                 ' Col H (TURNO)
                
                arrSoto1(r, 9) = ObtenerValorCampo(rs, "DNI")                         ' Col I (DNI)
                arrSoto1(r, 10) = ObtenerValorCampo(rs, "Rango")                      ' Col J (RANGO)
                arrSoto1(r, 11) = ObtenerValorCampo(rs, "Nombre")                     ' Col K (Nombre)
                arrSoto1(r, 12) = ObtenerValorCampo(rs, "FechaEntrada")               ' Col L (ENTRADA)
                arrSoto1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")                ' Col M (SALIDA)
                arrSoto1(r, 14) = ObtenerValorCampo(rs, "DiasUso")                    ' Col N (DIAS USO)
                arrSoto1(r, 15) = ObtenerValorCampo(rs, "PAX")                        ' Col O (PAX)
                arrSoto1(r, 16) = ObtenerValorCampo(rs, "Resolucion")                 ' Col P (RESOLUCION)
                
                ' Col Q (17): PRECIO APTO (FÓRMULA PRESERVADA INTACTA)
                ' Col R (18): SUPLE OCUPAN (FÓRMULA PRESERVADA INTACTA)
                
                arrSoto2(r, 1) = ObtenerValorCampo(rs, "HabitacionesAsignadas")       ' Col S (APARTAMENTO)
                arrSoto2(r, 2) = ObtenerValorCampo(rs, "DtoFamNum")                   ' Col T (DTO. FAM. NUM.)
                
                ' Col U (21): IMPORTE (FÓRMULA PRESERVADA INTACTA)
                
                arrSoto3(r, 1) = ObtenerValorCampo(rs, "Telefono")                    ' Col V (TELEFONO)
                arrSoto3(r, 2) = ObtenerValorCampo(rs, "Direccion")                   ' Col W (DIRECCIÓN)
                
                vVal = ObtenerValorCampo(rs, "CP")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "CodigoPostal")
                arrSoto3(r, 3) = vVal                                                 ' Col X (CP)
                
                arrSoto3(r, 4) = ObtenerValorCampo(rs, "Poblacion")                   ' Col Y (POBLACIÓN)
                arrSoto3(r, 5) = ObtenerValorCampo(rs, "Provincia")                   ' Col Z (PROVINCIA)
                arrSoto3(r, 6) = ObtenerValorCampo(rs, "EstadoPago")                  ' Col AA (PAGADO)
                
                ' Col AB (28): FÓRMULA NOMBRE PDF (PRESERVADA INTACTA)
                ' Col AC (29): FÓRMULA COPIAR TEXTO EMAIL (PRESERVADA INTACTA)
                
                rs.MoveNext
            Next r
            
            ' Volcado ultra-rápido en bloques COM preservando columnas de fórmulas
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(filaInicio + totalRows - 1, 16)).Value = arrSoto1
            hoja.Range(hoja.Cells(filaInicio, 19), hoja.Cells(filaInicio + totalRows - 1, 20)).Value = arrSoto2
            hoja.Range(hoja.Cells(filaInicio, 22), hoja.Cells(filaInicio + totalRows - 1, 27)).Value = arrSoto3
            
        Else
            ' GIJÓN: Bloque 1 (Cols A a T = 1 a 20), Bloque 2 (Cols V a AA = 22 a 27)
            ' Col U (21, IMPORTE), Col AB (28, NOMBRE PDF) y Col AC (29, TEXTO EMAIL) se mantienen 100% INTACTAS con sus fórmulas
            Dim arrGijon1() As Variant
            Dim arrGijon2() As Variant
            ReDim arrGijon1(1 To totalRows, 1 To 20)
            ReDim arrGijon2(1 To totalRows, 1 To 6)
            
            For r = 1 To totalRows
                arrGijon1(r, 1) = ObtenerValorCampo(rs, "NumOrden")                   ' Col A (Nº ORDEN)
                arrGijon1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")               ' Col B (FECHA PETICION)
                arrGijon1(r, 3) = ObtenerValorCampo(rs, "NumFactura")                  ' Col C (NÚM FACT)
                arrGijon1(r, 4) = ObtenerValorCampo(rs, "Finalidad")                   ' Col D (FINALIDAD)
                arrGijon1(r, 5) = ObtenerValorCampo(rs, "Empleo")                      ' Col E (EMPLEO)
                arrGijon1(r, 6) = ObtenerValorCampo(rs, "Situacion")                   ' Col F (SITUACION)
                arrGijon1(r, 7) = ObtenerValorCampo(rs, "Evaluacion")                  ' Col G (EVALUACIÓN)
                
                vVal = ObtenerValorCampo(rs, "Comision")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "Turno")
                arrGijon1(r, 8) = vVal                                                 ' Col H (COMISIÓN)
                
                arrGijon1(r, 9) = ObtenerValorCampo(rs, "DNI")                         ' Col I (DNI)
                arrGijon1(r, 10) = ObtenerValorCampo(rs, "Rango")                      ' Col J (RANGO)
                arrGijon1(r, 11) = ObtenerValorCampo(rs, "Nombre")                     ' Col K (Nombre)
                arrGijon1(r, 12) = ObtenerValorCampo(rs, "FechaEntrada")               ' Col L (ENTRADA)
                arrGijon1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")                ' Col M (SALIDA)
                arrGijon1(r, 14) = ObtenerValorCampo(rs, "DiasUso")                    ' Col N (DIAS USO)
                arrGijon1(r, 15) = ObtenerValorCampo(rs, "PAX")                        ' Col O (PAX)
                arrGijon1(r, 16) = ObtenerValorCampo(rs, "Resolucion")                 ' Col P (RESOLUCION)
                arrGijon1(r, 17) = ObtenerValorCampo(rs, "NumHabIndividuales")         ' Col Q (HAB. IND.)
                arrGijon1(r, 18) = ObtenerValorCampo(rs, "NumHabDobles")               ' Col R (HAB. DOBLE)
                arrGijon1(r, 19) = ObtenerValorCampo(rs, "HabitacionesAsignadas")      ' Col S (NÚM HAB.)
                arrGijon1(r, 20) = ObtenerValorCampo(rs, "DtoFamNum")                  ' Col T (DTO. FAM. NUM.)
                
                ' Col U (21): IMPORTE (FÓRMULA PRESERVADA INTACTA)
                
                arrGijon2(r, 1) = ObtenerValorCampo(rs, "Telefono")                    ' Col V (TELEFONO)
                arrGijon2(r, 2) = ObtenerValorCampo(rs, "Direccion")                   ' Col W (DIRECCIÓN)
                
                vVal = ObtenerValorCampo(rs, "CP")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "CodigoPostal")
                arrGijon2(r, 3) = vVal                                                 ' Col X (CP)
                
                arrGijon2(r, 4) = ObtenerValorCampo(rs, "Poblacion")                   ' Col Y (POBLACIÓN)
                arrGijon2(r, 5) = ObtenerValorCampo(rs, "Provincia")                   ' Col Z (PROVINCIA)
                arrGijon2(r, 6) = ObtenerValorCampo(rs, "EstadoPago")                  ' Col AA (PAGADO)
                
                ' Col AB (28): FÓRMULA NOMBRE PDF (PRESERVADA INTACTA)
                ' Col AC (29): FÓRMULA COPIAR TEXTO EMAIL (PRESERVADA INTACTA)
                
                rs.MoveNext
            Next r
            
            ' Volcado ultra-rápido en bloques COM preservando columnas de fórmulas
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(filaInicio + totalRows - 1, 20)).Value = arrGijon1
            hoja.Range(hoja.Cells(filaInicio, 22), hoja.Cells(filaInicio + totalRows - 1, 27)).Value = arrGijon2
            
        End If
        
        ' =========================================================================
        ' BLINDAJE Y AUTO-REPARACIÓN UNIVERSAL DE FÓRMULAS POST-SINCRONIZACIÓN
        ' Garantiza que cada fila volcada conserve o repare sus fórmulas maestras
        ' =========================================================================
        Dim colsFormula() As Long
        Dim idxF As Long, cCol As Long
        Dim filaRef As Long, fRow As Long
        Dim fmlaMaster As String
        
        If isOviedo Then
            ' OVIEDO: V (22, IMPORTE), AF (32, NOMBRE PDF), AG (33, TEXTO EMAIL)
            ReDim colsFormula(1 To 3)
            colsFormula(1) = 22
            colsFormula(2) = 32
            colsFormula(3) = 33
        ElseIf isSoto Then
            ' SOTO: Q (17, PRECIO APTO), R (18, SUPLE OCUPAN), U (21, IMPORTE), AB (28, NOMBRE PDF), AC (29, TEXTO EMAIL)
            ReDim colsFormula(1 To 5)
            colsFormula(1) = 17
            colsFormula(2) = 18
            colsFormula(3) = 21
            colsFormula(4) = 28
            colsFormula(5) = 29
        Else
            ' GIJÓN: U (21, IMPORTE), AB (28, NOMBRE PDF), AC (29, TEXTO EMAIL)
            ReDim colsFormula(1 To 3)
            colsFormula(1) = 21
            colsFormula(2) = 28
            colsFormula(3) = 29
        End If
        
        For idxF = LBound(colsFormula) To UBound(colsFormula)
            cCol = colsFormula(idxF)
            filaRef = 0
            
            ' 1. Probar en las primeras filas operativas
            If hoja.Cells(filaInicio, cCol).HasFormula Then
                filaRef = filaInicio
            ElseIf hoja.Cells(filaInicio + 1, cCol).HasFormula Then
                filaRef = filaInicio + 1
            Else
                ' 2. Buscar en filas de plantilla posteriores no afectadas por la carga
                Dim rScan As Long
                For rScan = filaInicio + totalRows To filaInicio + totalRows + 100
                    If hoja.Cells(rScan, cCol).HasFormula Then
                        filaRef = rScan
                        Exit For
                    End If
                Next rScan
            End If
            
            ' Si encontramos la fórmula maestra de referencia, propagarla a toda celda que carezca de ella
            If filaRef > 0 Then
                fmlaMaster = hoja.Cells(filaRef, cCol).FormulaR1C1
                For fRow = filaInicio To filaInicio + totalRows - 1
                    If Not hoja.Cells(fRow, cCol).HasFormula Then
                        hoja.Cells(fRow, cCol).FormulaR1C1 = fmlaMaster
                    End If
                Next fRow
            End If
        Next idxF
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
        residencia = UCase(Trim(m_ResidenciasAsignadas(i)))
        
        ' Sincronizar únicamente las residencias seleccionadas para esta sesión
        If InStr(m_ResidenciaActiva, residencia) > 0 Then
            nombreHoja = ObtenerNombreHoja(residencia)
            Set ws = Nothing
            
            If Len(nombreHoja) > 0 Then
                Set ws = ObtenerHojaSegura(nombreHoja)
                
                If Not ws Is Nothing Then
                    Application.StatusBar = "Sincronizando " & nombreHoja & "..."
                    SincronizarHojaDesdeAccess residencia, ws
                End If
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

''' Muestra las hojas RESIDENCIA, Calendario y RESUMEN únicamente de las residencias seleccionadas,
''' oculta las no seleccionadas, activa la hoja de la primera residencia seleccionada y oculta INICIO.
Public Sub MostrarHojasResidencia(Optional ByVal residenciaEspecifica As String = "")
    Dim resSeleccionada As String
    Dim libroProtegida As Boolean
    Dim mostrarGijon As Boolean
    Dim mostrarSoto As Boolean
    Dim mostrarOviedo As Boolean
    Dim wsActivar As Worksheet
    
    resSeleccionada = UCase(Trim(residenciaEspecifica))
    If Len(resSeleccionada) = 0 Then resSeleccionada = UCase(Trim(m_ResidenciaActiva))
    
    ' Detectar qué residencias están incluidas en la selección
    mostrarGijon = (InStr(resSeleccionada, "GIJON") > 0 Or InStr(resSeleccionada, "GIJÓN") > 0)
    mostrarSoto = (InStr(resSeleccionada, "SOTO") > 0)
    mostrarOviedo = (InStr(resSeleccionada, "OVIEDO") > 0)
    
    ' Si por alguna razón ninguna coincidió, activar GIJÓN por defecto
    If Not mostrarGijon And Not mostrarSoto And Not mostrarOviedo Then
        mostrarGijon = True
    End If
    
    On Error Resume Next
    Application.screenUpdating = False
    
    ' Desproteger la estructura del libro si está protegida (necesario para alternar visibilidad)
    DesprotegerEstructuraLibro libroProtegida
    
    ' Referencias a las hojas de cada residencia
    Dim wsResGijon As Worksheet, wsResSoto As Worksheet, wsResOviedo As Worksheet
    Dim wsCalGijon As Worksheet, wsCalSoto As Worksheet, wsCalOviedo As Worksheet
    Dim wsSumGijon As Worksheet, wsSumSoto As Worksheet, wsSumOviedo As Worksheet
    
    Set wsResGijon = ObtenerHojaSegura("RESIDENCIA GIJÓN")
    Set wsCalGijon = ObtenerHojaSegura("Calendario GIJÓN")
    Set wsSumGijon = ObtenerHojaSegura("RESUMEN GIJÓN")
    
    Set wsResSoto = ObtenerHojaSegura("RESIDENCIA SOTO")
    Set wsCalSoto = ObtenerHojaSegura("Calendario SOTO")
    Set wsSumSoto = ObtenerHojaSegura("RESUMEN SOTO")
    
    Set wsResOviedo = ObtenerHojaSegura("RESIDENCIA OVIEDO")
    Set wsCalOviedo = ObtenerHojaSegura("Calendario OVIEDO")
    Set wsSumOviedo = ObtenerHojaSegura("RESUMEN OVIEDO")
    
    ' -------------------------------------------------------------
    ' 1. MOSTRAR PRIMERO las hojas de las residencias seleccionadas
    ' (Garantiza que siempre haya al menos una hoja visible en Excel)
    ' -------------------------------------------------------------
    If mostrarGijon Then
        If Not wsResGijon Is Nothing Then wsResGijon.visible = xlSheetVisible
        If Not wsCalGijon Is Nothing Then wsCalGijon.visible = xlSheetVisible
        If Not wsSumGijon Is Nothing Then wsSumGijon.visible = xlSheetVisible
    End If
    
    If mostrarSoto Then
        If Not wsResSoto Is Nothing Then wsResSoto.visible = xlSheetVisible
        If Not wsCalSoto Is Nothing Then wsCalSoto.visible = xlSheetVisible
        If Not wsSumSoto Is Nothing Then wsSumSoto.visible = xlSheetVisible
    End If
    
    If mostrarOviedo Then
        If Not wsResOviedo Is Nothing Then wsResOviedo.visible = xlSheetVisible
        If Not wsCalOviedo Is Nothing Then wsCalOviedo.visible = xlSheetVisible
        If Not wsSumOviedo Is Nothing Then wsSumOviedo.visible = xlSheetVisible
    End If
    
    ' -------------------------------------------------------------
    ' 2. DETERMINAR Y ACTIVAR la hoja principal en pantalla
    ' -------------------------------------------------------------
    If mostrarGijon And Not wsResGijon Is Nothing Then
        Set wsActivar = wsResGijon
    ElseIf mostrarSoto And Not wsResSoto Is Nothing Then
        Set wsActivar = wsResSoto
    ElseIf mostrarOviedo And Not wsResOviedo Is Nothing Then
        Set wsActivar = wsResOviedo
    End If
    
    If Not wsActivar Is Nothing Then
        wsActivar.Activate
    End If
    
    ' -------------------------------------------------------------
    ' 3. OCULTAR las hojas de las residencias NO seleccionadas
    ' -------------------------------------------------------------
    If Not mostrarGijon Then
        If Not wsResGijon Is Nothing Then wsResGijon.visible = xlSheetHidden
        If Not wsCalGijon Is Nothing Then wsCalGijon.visible = xlSheetHidden
        If Not wsSumGijon Is Nothing Then wsSumGijon.visible = xlSheetHidden
    End If
    
    If Not mostrarSoto Then
        If Not wsResSoto Is Nothing Then wsResSoto.visible = xlSheetHidden
        If Not wsCalSoto Is Nothing Then wsCalSoto.visible = xlSheetHidden
        If Not wsSumSoto Is Nothing Then wsSumSoto.visible = xlSheetHidden
    End If
    
    If Not mostrarOviedo Then
        If Not wsResOviedo Is Nothing Then wsResOviedo.visible = xlSheetHidden
        If Not wsCalOviedo Is Nothing Then wsCalOviedo.visible = xlSheetHidden
        If Not wsSumOviedo Is Nothing Then wsSumOviedo.visible = xlSheetHidden
    End If
    
    ' -------------------------------------------------------------
    ' 4. OCULTAR la hoja INICIO
    ' -------------------------------------------------------------
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

''' Agrega una residencia al campo Residencias de un usuario existente.
''' Si la residencia ya está asignada, no la duplica.
''' Solo para administradores.
Public Function AgregarResidenciaUsuario(ByVal nombreUsuario As String, ByVal residencia As String) As Boolean
    Dim cn As Object
    Dim rs As Object
    Dim sql As String
    Dim resActuales As String
    Dim resNueva As String
    
    If m_Rol <> "ADMIN" Then
        MsgBox "Solo los administradores pueden modificar asignaciones.", vbExclamation
        AgregarResidenciaUsuario = False
        Exit Function
    End If
    
    residencia = UCase(Trim(residencia))
    If Len(residencia) = 0 Then
        AgregarResidenciaUsuario = False
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    
    ' Obtener las residencias actuales del usuario
    sql = "SELECT Residencias FROM Usuarios WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'"
    Set rs = cn.Execute(sql)
    
    If rs.EOF Then
        rs.Close
        cn.Close
        MsgBox "El usuario '" & nombreUsuario & "' no existe en la base de datos.", vbExclamation
        AgregarResidenciaUsuario = False
        GoTo CleanUp
    End If
    
    resActuales = UCase(Trim(Nz(rs("Residencias").Value, "")))
    rs.Close
    
    ' Verificar si ya tiene esa residencia asignada
    If InStr("," & resActuales & ",", "," & residencia & ",") > 0 Then
        MsgBox "El usuario '" & nombreUsuario & "' ya tiene asignada la residencia " & residencia & ".", vbInformation
        AgregarResidenciaUsuario = True
        GoTo CleanUp
    End If
    
    ' Construir la nueva cadena CSV
    If Len(resActuales) > 0 Then
        resNueva = resActuales & "," & residencia
    Else
        resNueva = residencia
    End If
    
    ' Actualizar en Access
    sql = "UPDATE Usuarios SET Residencias = '" & EscaparSQL(resNueva) & "' " & _
          "WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'"
    cn.Execute sql
    cn.Close
    
    AgregarResidenciaUsuario = True
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error al agregar residencia: " & vbCrLf & Err.Description, vbCritical
    AgregarResidenciaUsuario = False

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
    End If
    On Error GoTo 0
End Function

''' Quita una residencia del campo Residencias de un usuario.
''' Si era la última residencia, desactiva el usuario.
''' Solo para administradores.
Public Function QuitarResidenciaUsuario(ByVal nombreUsuario As String, ByVal residencia As String) As Boolean
    Dim cn As Object
    Dim rs As Object
    Dim sql As String
    Dim resActuales As String
    Dim arrRes() As String
    Dim resNueva As String
    Dim i As Long
    
    If m_Rol <> "ADMIN" Then
        MsgBox "Solo los administradores pueden modificar asignaciones.", vbExclamation
        QuitarResidenciaUsuario = False
        Exit Function
    End If
    
    residencia = UCase(Trim(residencia))
    If Len(residencia) = 0 Then
        QuitarResidenciaUsuario = False
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    
    ' Obtener las residencias actuales del usuario
    sql = "SELECT Residencias FROM Usuarios WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'"
    Set rs = cn.Execute(sql)
    
    If rs.EOF Then
        rs.Close
        cn.Close
        MsgBox "El usuario '" & nombreUsuario & "' no existe.", vbExclamation
        QuitarResidenciaUsuario = False
        GoTo CleanUp
    End If
    
    resActuales = UCase(Trim(Nz(rs("Residencias").Value, "")))
    rs.Close
    
    ' Verificar si tiene esa residencia asignada
    If InStr("," & resActuales & ",", "," & residencia & ",") = 0 Then
        MsgBox "El usuario '" & nombreUsuario & "' no tiene asignada la residencia " & residencia & ".", vbInformation
        QuitarResidenciaUsuario = True
        GoTo CleanUp
    End If
    
    ' Reconstruir CSV sin esa residencia
    arrRes = Split(resActuales, ",")
    resNueva = ""
    For i = LBound(arrRes) To UBound(arrRes)
        If UCase(Trim(arrRes(i))) <> residencia Then
            If Len(resNueva) > 0 Then resNueva = resNueva & ","
            resNueva = resNueva & UCase(Trim(arrRes(i)))
        End If
    Next i
    
    If Len(resNueva) = 0 Then
        ' Era la última residencia: desactivar usuario
        sql = "UPDATE Usuarios SET Residencias = '', Activo = False " & _
              "WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'"
        cn.Execute sql
        cn.Close
        MsgBox "Se ha retirado la última residencia del usuario '" & nombreUsuario & "'." & vbCrLf & _
               "El usuario ha sido desactivado.", vbInformation
    Else
        sql = "UPDATE Usuarios SET Residencias = '" & EscaparSQL(resNueva) & "' " & _
              "WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'"
        cn.Execute sql
        cn.Close
    End If
    
    QuitarResidenciaUsuario = True
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error al quitar residencia: " & vbCrLf & Err.Description, vbCritical
    QuitarResidenciaUsuario = False

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
    End If
    On Error GoTo 0
End Function

''' Consulta los usuarios asignados a una residencia específica.
''' Retorna una cadena formateada con la lista de usuarios.
''' Si residencia está vacío, retorna todos los usuarios.
Public Function ConsultarUsuariosPorResidencia(Optional ByVal residencia As String = "") As String
    Dim cn As Object
    Dim rs As Object
    Dim sql As String
    Dim resultado As String
    Dim resField As String
    Dim contador As Long
    
    On Error GoTo ErrorHandler
    
    residencia = UCase(Trim(residencia))
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    
    sql = "SELECT NombreUsuario, NombreCompleto, Residencias, Rol, Activo FROM Usuarios ORDER BY NombreUsuario"
    Set rs = cn.Execute(sql)
    
    resultado = ""
    contador = 0
    
    Do Until rs.EOF
        resField = UCase(Trim(Nz(rs("Residencias").Value, "")))
        
        ' Si se pidió filtrar por residencia, verificar que la contenga
        If Len(residencia) = 0 Or InStr("," & resField & ",", "," & residencia & ",") > 0 Then
            contador = contador + 1
            resultado = resultado & contador & ". " & rs("NombreUsuario").Value
            If Len(Nz(rs("NombreCompleto").Value, "")) > 0 Then
                resultado = resultado & " (" & rs("NombreCompleto").Value & ")"
            End If
            resultado = resultado & vbCrLf
            resultado = resultado & "    Rol: " & Nz(rs("Rol").Value, "OPERADOR")
            resultado = resultado & " | Residencias: " & resField
            If Not CBool(rs("Activo").Value) Then
                resultado = resultado & " [INACTIVO]"
            End If
            resultado = resultado & vbCrLf
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    cn.Close
    
    If contador = 0 Then
        If Len(residencia) > 0 Then
            resultado = "No hay usuarios asignados a la residencia " & residencia & "."
        Else
            resultado = "No hay usuarios registrados en el sistema."
        End If
    Else
        Dim encabezado As String
        If Len(residencia) > 0 Then
            encabezado = "Usuarios asignados a " & residencia & " (" & contador & "):" & vbCrLf & vbCrLf
        Else
            encabezado = "Todos los usuarios del sistema (" & contador & "):" & vbCrLf & vbCrLf
        End If
        resultado = encabezado & resultado
    End If
    
    ConsultarUsuariosPorResidencia = resultado
    Exit Function

ErrorHandler:
    ConsultarUsuariosPorResidencia = "Error al consultar usuarios: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
    End If
    On Error GoTo 0
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
' ELIMINACIÓN DE ÓRDENES Y AUDITORÍA
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
            "Archivado en Histórico y retirado NÂº ORDEN " & numOrden & " de " & claveRes)
    End If
    
    EliminarOrdenBD = exito
End Function

' ==============================================================================
' GESTIÓN DE ARCHIVADO HISTÓRICO Y PURGA RGPD EN ACCESS (ORDENES_HISTORICO)
' ==============================================================================

''' Cuenta cuántas órdenes en la tabla activa 'Ordenes' han superado el plazo
''' de retención operativa (por defecto 90 días desde FechaSalida).
''' Si se omite residencia, cuenta el total de todas las residencias.
Public Function ContarOrdenesArchivablesBD( _
    Optional ByVal residencia As String = "", _
    Optional ByVal diasRetencion As Long = 90 _
) As Long
    Dim sql As String
    Dim whereClause As String
    Dim claveRes As String
    Dim res As Variant
    
    whereClause = "FechaSalida <= DateAdd('d', -" & diasRetencion & ", Date())"
    
    If Len(Trim(residencia)) > 0 Then
        claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
        If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
        whereClause = whereClause & " AND Residencia = '" & EscaparSQL(claveRes) & "'"
    End If
    
    sql = "SELECT COUNT(*) FROM Ordenes WHERE " & whereClause
    res = ExecuteScalar(sql)
    
    If IsNumeric(res) Then
        ContarOrdenesArchivablesBD = CLng(res)
    Else
        ContarOrdenesArchivablesBD = 0
    End If
End Function

''' Traslada de forma atómica y transaccional las órdenes con FechaSalida > diasRetencion
''' desde la tabla activa 'Ordenes' hacia 'Ordenes_Historico' en Access.
''' Retorna el número de órdenes efectivamente archivadas.
Public Function ArchivarOrdenesAntiguasBD( _
    Optional ByVal residencia As String = "", _
    Optional ByVal diasRetencion As Long = 90 _
) As Long
    Dim totalArchivables As Long
    Dim whereClause As String
    Dim claveRes As String
    Dim cn As Object
    Dim exito As Boolean
    Dim usr As String
    
    ' 1. Comprobar si hay registros para archivar
    totalArchivables = ContarOrdenesArchivablesBD(residencia, diasRetencion)
    If totalArchivables = 0 Then
        ArchivarOrdenesAntiguasBD = 0
        Exit Function
    End If
    
    whereClause = "FechaSalida <= DateAdd('d', -" & diasRetencion & ", Date())"
    If Len(Trim(residencia)) > 0 Then
        claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
        If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
        whereClause = whereClause & " AND Residencia = '" & EscaparSQL(claveRes) & "'"
    Else
        claveRes = "TODAS"
    End If
    
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    cn.BeginTrans
    
    ' 2. Limpieza preventiva en histórico para evitar duplicidad de clave primaria (idempotencia)
    cn.Execute "DELETE FROM Ordenes_Historico WHERE Id IN " & _
               "(SELECT Id FROM Ordenes WHERE " & whereClause & ")"
    
    ' 3. Traspaso completo de registros hacia Ordenes_Historico
    cn.Execute "INSERT INTO Ordenes_Historico SELECT * FROM Ordenes WHERE " & whereClause
    
    ' 4. Retirada de los registros de la tabla activa Ordenes
    cn.Execute "DELETE FROM Ordenes WHERE " & whereClause
    
    cn.CommitTrans
    cn.Close
    Set cn = Nothing
    exito = True
    
    ' 5. Registro en auditoría
    usr = m_NombreUsuario
    If Len(usr) = 0 Then usr = Environ("USERNAME")
    Call InsertarLog(usr, "ARCHIVADO_RGPD", _
        "Archivadas " & totalArchivables & " órdenes de " & claveRes & _
        " (> " & diasRetencion & " días) a Ordenes_Historico")
    
    ArchivarOrdenesAntiguasBD = totalArchivables
    Exit Function

ErrorHandler:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then
            cn.RollbackTrans
            cn.Close
        End If
        Set cn = Nothing
    End If
    ArchivarOrdenesAntiguasBD = 0
    MsgBox "Error al archivar órdenes en Access:" & vbCrLf & Err.Description, vbCritical, "Error de Archivado"
End Function

''' Cuenta cuántas órdenes en 'Ordenes_Historico' han superado el plazo máximo
''' legal de conservación (por defecto 3 años según RD 933/2021).
Public Function ContarHistoricoPurgableRGPD( _
    Optional ByVal anosRetencion As Long = 3 _
) As Long
    Dim sql As String
    Dim res As Variant
    
    sql = "SELECT COUNT(*) FROM Ordenes_Historico WHERE FechaSalida <= DateAdd('yyyy', -" & anosRetencion & ", Date())"
    res = ExecuteScalar(sql)
    
    If IsNumeric(res) Then
        ContarHistoricoPurgableRGPD = CLng(res)
    Else
        ContarHistoricoPurgableRGPD = 0
    End If
End Function

''' Ejecuta el borrado físico y definitivo de datos personales de la base de datos
''' Access sobre órdenes históricas con antigüedad superior a anosRetencion (RD 933/2021).
''' RESERVADO EXCLUSIVAMENTE AL PANEL RGPD. Retorna el número de registros suprimidos.
Public Function PurgarHistoricoRGPD( _
    Optional ByVal anosRetencion As Long = 3 _
) As Long
    Dim totalPurgables As Long
    Dim sqlDelete As String
    Dim usr As String
    
    totalPurgables = ContarHistoricoPurgableRGPD(anosRetencion)
    If totalPurgables = 0 Then
        PurgarHistoricoRGPD = 0
        Exit Function
    End If
    
    sqlDelete = "DELETE FROM Ordenes_Historico WHERE FechaSalida <= DateAdd('yyyy', -" & anosRetencion & ", Date())"
    
    If ExecuteNonQuery(sqlDelete) Then
        ' Purgar también de la réplica Snapshot si existiera
        On Error Resume Next
        ExecuteNonQuery "DELETE FROM BDAS_Historico_Snapshot WHERE FechaSalida <= DateAdd('yyyy', -" & anosRetencion & ", Date())"
        On Error GoTo 0
        
        usr = m_NombreUsuario
        If Len(usr) = 0 Then usr = Environ("USERNAME")
        Call InsertarLog(usr, "PURGA_RGPD", _
            "Borrado físico definitivo de " & totalPurgables & " órdenes históricas (> " & anosRetencion & " años) según RD 933/2021")
        
        PurgarHistoricoRGPD = totalPurgables
    Else
        PurgarHistoricoRGPD = 0
    End If
End Function


' ==============================================================================
' CONSULTAS DE CÓDIGOS POSTALES Y MUNICIPIOS (TABLA CodigosPostales EN ACCESS)
' ==============================================================================

''' Normaliza un texto removiendo acentos y convirtiendo a mayúsculas para búsquedas.
Public Function NormalizarTextoBusqueda(ByVal texto As String) As String
    Dim res As String
    res = UCase(Trim(texto))
    If Len(res) = 0 Then
        NormalizarTextoBusqueda = ""
        Exit Function
    End If
    
    res = Replace(res, "Ã", "A")
    res = Replace(res, "Ã", "A")
    res = Replace(res, "Ã", "A")
    res = Replace(res, "Ã", "A")
    res = Replace(res, "Ã", "E")
    res = Replace(res, "Ã", "E")
    res = Replace(res, "Ã", "E")
    res = Replace(res, "Ã", "E")
    res = Replace(res, "Ã", "I")
    res = Replace(res, "Ã", "I")
    res = Replace(res, "Ã", "I")
    res = Replace(res, "Ã", "I")
    res = Replace(res, "Ó", "O")
    res = Replace(res, "Ã", "O")
    res = Replace(res, "Ã", "O")
    res = Replace(res, "Ã", "O")
    res = Replace(res, "Ã", "U")
    res = Replace(res, "Ã", "U")
    res = Replace(res, "Ã", "U")
    res = Replace(res, "Ã", "U")
    res = Replace(res, "Ã", "N")
    res = Replace(res, "Ã", "C")
    NormalizarTextoBusqueda = res
End Function

''' Consulta en Access la población y provincia asociadas a un código postal.
''' Si existen múltiples municipios para el CP, colMunicipios contendrá todos ellos.
''' Retorna True si encontró al menos un registro.
Public Function ObtenerUbicacionPorCP( _
    ByVal cp As String, _
    ByRef colMunicipios As Collection, _
    ByRef provincia As String) As Boolean
    
    Dim rs As Object
    Dim sql As String
    Dim cpLimpio As String
    Dim muni As String
    
    On Error GoTo ErrorHandler
    
    Set colMunicipios = New Collection
    provincia = ""
    ObtenerUbicacionPorCP = False
    
    cpLimpio = Trim(cp)
    If Len(cpLimpio) = 0 Then Exit Function
    If Len(cpLimpio) < 5 And IsNumeric(cpLimpio) Then
        cpLimpio = Right("00000" & cpLimpio, 5)
    End If
    
    sql = "SELECT DISTINCT Municipio, Provincia FROM CodigosPostales " & _
          "WHERE CodigoPostal = '" & EscaparSQL(cpLimpio) & "' ORDER BY Municipio"
          
    Set rs = GetRecordset(sql)
    If rs Is Nothing Then Exit Function
    
    Do While Not rs.EOF
        muni = Trim(CStr(rs("Municipio").Value))
        muni = Replace(muni, ",", ";")
        muni = Replace(muni, Chr(34), "'")
        If Len(muni) > 0 Then
            colMunicipios.Add muni
        End If
        If Len(provincia) = 0 Then
            provincia = Trim(CStr(rs("Provincia").Value))
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    ObtenerUbicacionPorCP = (colMunicipios.Count > 0)
    Exit Function

ErrorHandler:
    ObtenerUbicacionPorCP = False
End Function

''' Consulta en Access los códigos postales y provincia asociados a un municipio.
''' La búsqueda es insensible a mayúsculas y acentos gracias a la columna MunicipioNormalizado.
''' Retorna True si encontró al menos un registro.
Public Function ObtenerCPsPorPoblacion( _
    ByVal poblacion As String, _
    ByRef colCPs As Collection, _
    ByRef provincia As String) As Boolean
    
    Dim rs As Object
    Dim sql As String
    Dim muniNorm As String
    Dim cpVal As String
    
    On Error GoTo ErrorHandler
    
    Set colCPs = New Collection
    provincia = ""
    ObtenerCPsPorPoblacion = False
    
    muniNorm = NormalizarTextoBusqueda(poblacion)
    If Len(muniNorm) = 0 Then Exit Function
    
    sql = "SELECT DISTINCT CodigoPostal, Provincia FROM CodigosPostales " & _
          "WHERE MunicipioNormalizado = '" & EscaparSQL(muniNorm) & "' ORDER BY CodigoPostal"
          
    Set rs = GetRecordset(sql)
    If rs Is Nothing Then Exit Function
    
    Do While Not rs.EOF
        cpVal = Trim(CStr(rs("CodigoPostal").Value))
        If Len(cpVal) > 0 Then
            colCPs.Add cpVal
        End If
        If Len(provincia) = 0 Then
            provincia = Trim(CStr(rs("Provincia").Value))
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    ObtenerCPsPorPoblacion = (colCPs.Count > 0)
    Exit Function

ErrorHandler:
    ObtenerCPsPorPoblacion = False
End Function

' ==============================================================================
' GESTIÓN DE USUARIOS EN ACCESS - MODELO DELEGADO POR RESIDENCIA
' ==============================================================================

''' Comprueba si el usuario actual tiene permisos para gestionar la residencia indicada.
''' Un Administrador puede gestionar cualquier residencia.
''' Un Operador solo puede gestionar la residencia a la que pertenece.
Public Function PuedeGestionarResidencia(ByVal residencia As String) As Boolean
    Dim claveRes As String
    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    
    If m_Rol = "ADMIN" Then
        PuedeGestionarResidencia = True
    Else
        PuedeGestionarResidencia = TieneAccesoAResidencia(claveRes)
    End If
End Function

''' Consulta los datos de un usuario en la tabla Usuarios de Access.
Public Function ConsultarUsuarioAccess( _
    ByVal nombreUsuario As String, _
    ByRef nombreCompleto As String, _
    ByRef residencias As String, _
    ByRef rol As String, _
    ByRef activo As Boolean _
) As Boolean
    Dim cn As Object, rs As Object
    Dim sql As String
    
    On Error GoTo ErrorHandler
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5
    cn.Open GetConnectionString()
    
    sql = "SELECT NombreCompleto, Residencias, Rol, Activo FROM Usuarios " & _
          "WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'"
    Set rs = cn.Execute(sql)
    
    If Not rs.EOF Then
        nombreCompleto = Nz(rs("NombreCompleto").Value, "")
        residencias = Nz(rs("Residencias").Value, "")
        rol = Nz(rs("Rol").Value, "OPERADOR")
        activo = CBool(Nz(rs("Activo").Value, False))
        ConsultarUsuarioAccess = True
    Else
        ConsultarUsuarioAccess = False
    End If
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then On Error Resume Next: rs.Close
    If Not cn Is Nothing Then On Error Resume Next: cn.Close
    ConsultarUsuarioAccess = False
End Function

''' Da de alta un usuario en Access con clave SHA-256.
''' Si lo ejecuta un Operador, solo puede dar de alta para su propia sede y con rol OPERADOR.
Public Function AltaUsuarioDelegado( _
    ByVal nombreUsuario As String, _
    ByVal nombreCompleto As String, _
    ByVal clavePlana As String, _
    ByVal residenciaSede As String, _
    Optional ByVal rol As String = "OPERADOR" _
) As Boolean
    Dim claveRes As String
    claveRes = ObtenerClaveResidenciaDesdeHoja(residenciaSede)
    
    ' Validar permisos sobre la sede
    If Not PuedeGestionarResidencia(claveRes) Then
        MsgBox "No tienes permiso para gestionar usuarios en la sede " & residenciaSede & ".", vbExclamation
        AltaUsuarioDelegado = False
        Exit Function
    End If
    
    ' Un operador no puede crear administradores
    If m_Rol <> "ADMIN" Then rol = "OPERADOR"
    
    Dim claveHash As String
    claveHash = HashSHA256(clavePlana)
    
    Dim sql As String
    sql = "INSERT INTO Usuarios (NombreUsuario, NombreCompleto, Clave, Residencias, Rol, Activo, FechaAlta) VALUES (" & _
          "'" & EscaparSQL(nombreUsuario) & "', " & _
          "'" & EscaparSQL(nombreCompleto) & "', " & _
          "'" & EscaparSQL(claveHash) & "', " & _
          "'" & EscaparSQL(claveRes) & "', " & _
          "'" & EscaparSQL(UCase(rol)) & "', " & _
          "True, " & FormatearFechaHoraSQL(Now) & ")"
          
    If ExecuteNonQuery(sql) Then
        InsertarLog m_NombreUsuario, "ALTA_USUARIO", _
            "Alta de usuario " & nombreUsuario & " (" & rol & ") en [" & claveRes & "]"
        AltaUsuarioDelegado = True
    Else
        AltaUsuarioDelegado = False
    End If
End Function

''' Concede acceso a una sede a un usuario existente.
Public Function AgregarResidenciaUsuarioDelegado( _
    ByVal nombreUsuario As String, _
    ByVal nuevaResidencia As String _
) As Boolean
    Dim claveRes As String
    claveRes = ObtenerClaveResidenciaDesdeHoja(nuevaResidencia)
    
    If Not PuedeGestionarResidencia(claveRes) Then
        MsgBox "No tienes permiso para gestionar asignaciones en la sede " & nuevaResidencia & ".", vbExclamation
        AgregarResidenciaUsuarioDelegado = False
        Exit Function
    End If
    
    Dim nomComp As String, resActuales As String, rolUsr As String, actUsr As Boolean
    If Not ConsultarUsuarioAccess(nombreUsuario, nomComp, resActuales, rolUsr, actUsr) Then
        MsgBox "El usuario no existe en la base de datos.", vbExclamation
        AgregarResidenciaUsuarioDelegado = False
        Exit Function
    End If
    
    ' Si el objetivo es ADMIN y el actual no es ADMIN, no puede modificarlo
    If rolUsr = "ADMIN" And m_Rol <> "ADMIN" Then
        MsgBox "No tienes permiso para modificar una cuenta de Administrador.", vbCritical
        AgregarResidenciaUsuarioDelegado = False
        Exit Function
    End If
    
    ' Comprobar si ya la tiene asignada
    Dim arrRes() As String, i As Long
    arrRes = Split(resActuales, ",")
    For i = LBound(arrRes) To UBound(arrRes)
        If Trim(UCase(arrRes(i))) = claveRes Then
            AgregarResidenciaUsuarioDelegado = True
            Exit Function
        End If
    Next i
    
    Dim resFinal As String
    If Len(Trim(resActuales)) = 0 Then
        resFinal = claveRes
    Else
        resFinal = Trim(resActuales) & "," & claveRes
    End If
    
    Dim sql As String
    sql = "UPDATE Usuarios SET Residencias = '" & EscaparSQL(resFinal) & "', Activo = True " & _
          "WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'"
          
    If ExecuteNonQuery(sql) Then
        InsertarLog m_NombreUsuario, "MODIFICAR_USUARIO", _
            "Concedido acceso a " & claveRes & " a usuario: " & nombreUsuario
        AgregarResidenciaUsuarioDelegado = True
    Else
        AgregarResidenciaUsuarioDelegado = False
    End If
End Function

''' Retira el acceso a una sede para un usuario.
Public Function QuitarResidenciaUsuarioDelegado( _
    ByVal nombreUsuario As String, _
    ByVal residenciaAQuitar As String _
) As Boolean
    Dim claveRes As String
    claveRes = ObtenerClaveResidenciaDesdeHoja(residenciaAQuitar)
    
    If Not PuedeGestionarResidencia(claveRes) Then
        MsgBox "No tienes permiso para retirar accesos en la sede " & residenciaAQuitar & ".", vbExclamation
        QuitarResidenciaUsuarioDelegado = False
        Exit Function
    End If
    
    Dim nomComp As String, resActuales As String, rolUsr As String, actUsr As Boolean
    If Not ConsultarUsuarioAccess(nombreUsuario, nomComp, resActuales, rolUsr, actUsr) Then
        MsgBox "El usuario no existe en la base de datos.", vbExclamation
        QuitarResidenciaUsuarioDelegado = False
        Exit Function
    End If
    
    If rolUsr = "ADMIN" And m_Rol <> "ADMIN" Then
        MsgBox "No tienes permiso para modificar una cuenta de Administrador.", vbCritical
        QuitarResidenciaUsuarioDelegado = False
        Exit Function
    End If
    
    Dim arrRes() As String, i As Long, resFinal As String
    arrRes = Split(resActuales, ",")
    resFinal = ""
    
    For i = LBound(arrRes) To UBound(arrRes)
        Dim item As String
        item = Trim(UCase(arrRes(i)))
        If item <> claveRes And Len(item) > 0 Then
            If Len(resFinal) = 0 Then
                resFinal = item
            Else
                resFinal = resFinal & "," & item
            End If
        End If
    Next i
    
    Dim sql As String
    sql = "UPDATE Usuarios SET Residencias = '" & EscaparSQL(resFinal) & "' " & _
          "WHERE NombreUsuario = '" & EscaparSQL(nombreUsuario) & "'"
          
    If ExecuteNonQuery(sql) Then
        InsertarLog m_NombreUsuario, "MODIFICAR_USUARIO", _
            "Retirado acceso a " & claveRes & " para usuario: " & nombreUsuario
        QuitarResidenciaUsuarioDelegado = True
    Else
        QuitarResidenciaUsuarioDelegado = False
    End If
End Function

''' Lista usuarios con acceso a la sede indicada.
Public Function ObtenerListaUsuariosDelegada(ByVal residenciaSede As String) As String
    Dim claveRes As String, filtroToken As String
    claveRes = ObtenerClaveResidenciaDesdeHoja(residenciaSede)
    
    If m_Rol = "ADMIN" And (claveRes <> "GIJON" And claveRes <> "SOTO" And claveRes <> "OVIEDO") Then
        filtroToken = "" ' Administrador viendo todos
    Else
        filtroToken = claveRes
    End If
    
    Dim cn As Object, rs As Object
    Dim sql As String, resultado As String
    sql = "SELECT NombreUsuario, NombreCompleto, Rol, Residencias, Activo FROM Usuarios ORDER BY NombreUsuario"
    
    On Error GoTo ErrorHandler
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5
    cn.Open GetConnectionString()
    Set rs = cn.Execute(sql)
    
    resultado = ""
    Do While Not rs.EOF
        Dim resStr As String, tieneAcceso As Boolean
        resStr = Nz(rs("Residencias").Value, "")
        
        If Len(filtroToken) = 0 Then
            tieneAcceso = True
        Else
            tieneAcceso = (InStr("," & UCase(resStr) & ",", "," & filtroToken & ",") > 0)
        End If
        
        If tieneAcceso Then
            Dim estado As String
            estado = IIf(CBool(Nz(rs("Activo").Value, False)), "ACTIVO", "INACTIVO")
            resultado = resultado & "• " & rs("NombreUsuario").Value & " (" & Nz(rs("NombreCompleto").Value, "-") & ") " & _
                        "[" & Nz(rs("Rol").Value, "OPERADOR") & "] - Estado: " & estado & vbCrLf & _
                        "   Residencias: " & resStr & vbCrLf & vbCrLf
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
    If Len(resultado) = 0 Then resultado = "No hay usuarios registrados para esta sede."
    ObtenerListaUsuariosDelegada = resultado
    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then On Error Resume Next: rs.Close
    If Not cn Is Nothing Then On Error Resume Next: cn.Close
    ObtenerListaUsuariosDelegada = "Error al consultar usuarios en Access: " & Err.Description
End Function
