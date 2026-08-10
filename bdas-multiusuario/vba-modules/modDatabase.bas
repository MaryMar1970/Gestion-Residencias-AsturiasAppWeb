Attribute VB_Name = "modDatabase"
Option Explicit

' ==============================================================================
' Módulo: modDatabase.bas
' Propósito: Capa de acceso a datos ADO entre Excel Front-End y MS Access
'            Back-End (H:\ResidenciaBD\Residencia_BE.accdb)
'
' Arquitectura: Archivo .xlsm ÚNICO compartido en H:\ (hasta 15 operadores)
'   - El .xlsm se abre en memoria por cada usuario (solo lectura aceptable)
'   - Access es la ÚNICA fuente de verdad
'   - Las celdas de Excel son caché visual en memoria (NO se guardan al .xlsm)
'   - Login automático por Environ("USERNAME") contra tabla Usuarios de Access
'
' Todas las conexiones siguen el patrón de VIDA CORTA:
'   Abrir -> Ejecutar -> Cerrar (inmediatamente)
'
' Funciones disponibles:
'   CONEXIÓN Y LOGIN:
'   - GetConnectionString()              -> Cadena de conexión OLEDB
'   - ComprobarConexion()                -> Test de conectividad con Access
'   - LoginUsuario()                     -> Login automático por Windows username
'   - ObtenerUsuarioActual()             -> Nombre de usuario Windows actual
'   - ObtenerNombreCompleto()            -> Nombre completo del usuario logueado
'   - ObtenerRolUsuario()                -> Rol del usuario (OPERADOR/ADMIN)
'
'   RESIDENCIAS:
'   - ObtenerResidenciasAsignadas()      -> Collection con las residencias del usuario
'   - ObtenerResidenciaActiva()          -> Residencia seleccionada actualmente
'   - CambiarResidenciaActiva(res)       -> Cambiar la residencia activa
'   - TieneAccesoAResidencia(res)        -> Verificar permiso sobre una residencia
'
'   OPERACIONES CRUD:
'   - ExecuteNonQuery(sql)               -> INSERT/UPDATE/DELETE
'   - ExecuteScalar(sql)                 -> SELECT que retorna un solo valor
'   - GetRecordset(sql)                  -> SELECT que retorna un Recordset desconectado
'   - InsertarOrdenAtomica(...)          -> Alta de solicitud con N Orden atómico
'   - ActualizarOrden(...)               -> UPDATE genérico por N Orden
'   - ActualizarOrdenMultiple(...)       -> UPDATE múltiples campos por N Orden
'   - ObtenerSiguienteNumFactura()       -> Incremento atómico del contador de facturas
'
'   BÚSQUEDAS:
'   - BuscarEnOrdenes(...)               -> SELECT con filtro para búsquedas
'   - ExisteDuplicado(...)               -> Verificar duplicados de solicitud
'   - ComprobarListaNegra(dni)           -> Verificar si un DNI está vetado
'   - AnadirAListaNegra(...)             -> Añadir DNI a lista negra
'
'   SINCRONIZACIÓN:
'   - SincronizarHojaDesdeAccess(...)    -> Volcar datos de Access a una hoja Excel
'   - SincronizarResidenciaActiva()      -> Sincronizar solo la residencia seleccionada
'   - SincronizarTodasMisResidencias()   -> Sincronizar todas las residencias del usuario
'   - RefrescarCacheVisual()             -> Alias para SincronizarResidenciaActiva
'   - ActivarHojaResidenciaActiva()      -> Navegar a la hoja de la residencia activa
'
'   LOG:
'   - InsertarLog(...)                   -> Escribir en LogActividad
'
'   UTILIDADES:
'   - EscaparSQL(texto)                  -> Sanitizar texto contra inyección SQL
'   - FormatearFechaSQL(fecha)           -> Formato fecha para Access SQL
'   - FormatearFechaHoraSQL(fecha)       -> Formato fecha+hora para Access SQL
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
    
    ' Login exitoso - ahora seleccionar residencia
    If m_ResidenciasAsignadas.Count = 1 Then
        ' Solo tiene una residencia -> activar directamente
        m_ResidenciaActiva = m_ResidenciasAsignadas(1)
    Else
        ' Tiene varias -> mostrar selector
        frmSelectorResidencia.Show vbModal
        
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
    ' Nada que limpiar aquí, la conexión se gestiona en ValidarCredenciales
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
        Case "GIJON": ObtenerNombreHoja = "BDAS GIJÓN"
        Case "SOTO": ObtenerNombreHoja = "BDAS SOTO"
        Case "OVIEDO": ObtenerNombreHoja = "BDAS OVIEDO"
        Case Else: ObtenerNombreHoja = ""
    End Select
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
' OPERACIONES GENÉRICAS
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
''' Útil para COUNT(*), MAX(...), SELECT @@IDENTITY, etc.
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
    
    On Error GoTo ErrorHandler
    
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    
    ' Iniciar transacción atómica
    cn.BeginTrans
    
    sql = "INSERT INTO Ordenes (" & _
          "Residencia, FechaPeticion, DNI, Nombre, Apellidos, TipoHuesped, " & _
          "NumHabIndividuales, NumHabDobles, FechaEntrada, FechaSalida, " & _
          "Resolucion, EstadoPago, Observaciones, " & _
          "Telefono, Email, Direccion, CodigoPostal, Poblacion, Provincia, " & _
          "Solapamiento, ConsentimientoRGPD, " & _
          "FechaCreacion, UsuarioCreacion" & _
          ") VALUES ("
          
    sql = sql & "'" & EscaparSQL(residencia) & "', " & _
          FormatearFechaSQL(fechaPeticion) & ", " & _
          "'" & EscaparSQL(dni) & "', " & _
          "'" & EscaparSQL(nombre) & "', " & _
          "'" & EscaparSQL(apellidos) & "', " & _
          "'" & EscaparSQL(tipoHuesped) & "', " & _
          numHabInd & ", " & numHabDob & ", " & _
          FormatearFechaSQL(fechaEntrada) & ", " & _
          FormatearFechaSQL(fechaSalida) & ", " & _
          "'', '', " & _
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
    
    ' Obtener el autonumérico asignado por Access en la misma conexión
    Set rs = cn.Execute("SELECT @@IDENTITY")
    If Not rs.EOF Then
        nuevoNumOrden = CLng(rs(0).Value)
    End If
    rs.Close
    
    ' Confirmar transacción
    cn.CommitTrans
    cn.Close
    
    InsertarOrdenAtomica = nuevoNumOrden
    GoTo CleanUp

ErrorHandler:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then
            On Error Resume Next
            cn.RollbackTrans
            On Error GoTo 0
        End If
    End If
    MsgBox "Error al grabar la nueva orden en Access:" & vbCrLf & _
           Err.Description, vbCritical, "Error de Inserción"
    InsertarOrdenAtomica = 0

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

''' Busca órdenes en Access según un criterio.
''' criterio: "DNI", "NOMBRE", "NUMORDEN", "NUMFACTURA"
''' valor: el texto o número a buscar
''' residencia: opcional, filtrar por residencia ("GIJON", "SOTO", "OVIEDO", "" = todas)
''' Retorna un Recordset desconectado con los resultados.
Public Function BuscarEnOrdenes( _
    ByVal criterio As String, _
    ByVal valor As String, _
    Optional ByVal residencia As String = "" _
) As Object
    Dim sql As String
    Dim whereClause As String
    
    Select Case UCase(criterio)
        Case "DNI"
            whereClause = "DNI LIKE '%" & EscaparSQL(valor) & "%'"
        Case "NOMBRE"
            whereClause = "(Nombre LIKE '%" & EscaparSQL(valor) & "%' OR Apellidos LIKE '%" & EscaparSQL(valor) & "%')"
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
    
    sql = "SELECT * FROM Ordenes WHERE " & whereClause & " ORDER BY NumOrden DESC"
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
Public Function ComprobarListaNegra(ByVal dni As String) As Boolean
    Dim resultado As Variant
    
    resultado = ExecuteScalar( _
        "SELECT COUNT(*) FROM ListaNegra WHERE DNI = '" & EscaparSQL(dni) & "' AND Activo = True" _
    )
    
    ComprobarListaNegra = (Not IsNull(resultado) And CLng(Nz(resultado, 0)) > 0)
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
' SINCRONIZACIÓN ACCESS -> EXCEL (CACHÉ VISUAL)
' ===========================

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
        hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(ultimaFila, 30)).ClearContents
    End If
    
    ' Volcar datos del Recordset a las celdas
    ' NOTA: El mapeo de columnas debe coincidir con las columnas de la hoja original
    ' Columna A = NumOrden, C = NumFactura, K = Nombre, L = FechaEntrada, etc.
    Dim fila As Long
    fila = filaInicio
    
    Do While Not rs.EOF
        hoja.Cells(fila, 1).Value = rs("NumOrden").Value                  ' Col A
        If Not IsNull(rs("NumFactura").Value) Then _
            hoja.Cells(fila, 3).Value = rs("NumFactura").Value            ' Col C
        hoja.Cells(fila, 11).Value = rs("Nombre").Value                   ' Col K
        hoja.Cells(fila, 12).Value = rs("FechaEntrada").Value             ' Col L
        hoja.Cells(fila, 13).Value = rs("FechaSalida").Value              ' Col M
        hoja.Cells(fila, 16).Value = rs("Resolucion").Value               ' Col P
        
        ' Columna de habitaciones: S para Gijón/Soto, T para Oviedo
        If UCase(residencia) = "OVIEDO" Then
            hoja.Cells(fila, 20).Value = rs("HabitacionesAsignadas").Value ' Col T
        Else
            hoja.Cells(fila, 19).Value = rs("HabitacionesAsignadas").Value ' Col S
        End If
        
        ' Columna de estado de pago: AA para Gijón/Soto, AB para Oviedo
        If UCase(residencia) = "OVIEDO" Then
            hoja.Cells(fila, 28).Value = rs("EstadoPago").Value           ' Col AB
        Else
            hoja.Cells(fila, 27).Value = rs("EstadoPago").Value           ' Col AA
        End If
        
        ' Campos adicionales que también están en las hojas
        hoja.Cells(fila, 5).Value = rs("DNI").Value                       ' Col E (aprox)
        hoja.Cells(fila, 6).Value = rs("Apellidos").Value                 ' Col F (aprox)
        hoja.Cells(fila, 17).Value = rs("NumHabIndividuales").Value       ' Col Q
        hoja.Cells(fila, 18).Value = rs("NumHabDobles").Value             ' Col R
        
        rs.MoveNext
        fila = fila + 1
    Loop
    
    rs.Close
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error al sincronizar datos desde Access:" & vbCrLf & _
           Err.Description, vbCritical, "Error de Sincronización"

CleanUp:
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
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    
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
        
        If Len(nombreHoja) > 0 Then
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(nombreHoja)
            On Error GoTo 0
            
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

''' Activa (muestra) la hoja Excel de la residencia actualmente seleccionada.
Public Sub ActivarHojaResidenciaActiva()
    Dim nombreHoja As String
    nombreHoja = ObtenerNombreHoja(m_ResidenciaActiva)
    
    If Len(nombreHoja) > 0 Then
        On Error Resume Next
        ThisWorkbook.Sheets(nombreHoja).Activate
        On Error GoTo 0
    End If
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
