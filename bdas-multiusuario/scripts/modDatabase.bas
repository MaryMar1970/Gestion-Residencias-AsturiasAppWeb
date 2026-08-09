Attribute VB_Name = "modDatabase"
Option Explicit

' ==============================================================================
' Módulo: modDatabase.bas
' Propósito: Capa de acceso a datos ADO entre Excel Front-End y MS Access
'            Back-End (H:\ResidenciaBD\Residencia_BE.accdb)
'
' Solución puente multiusuario BDAS v16.5.5
' Todas las conexiones siguen el patrón de VIDA CORTA:
'   Abrir → Ejecutar → Cerrar (inmediatamente)
'
' Funciones disponibles:
'   - GetConnectionString()         → Cadena de conexión OLEDB
'   - EscaparSQL(texto)             → Sanitizar texto contra inyección SQL
'   - ExecuteNonQuery(sql)          → INSERT/UPDATE/DELETE
'   - ExecuteScalar(sql)            → SELECT que retorna un solo valor
'   - GetRecordset(sql)             → SELECT que retorna un Recordset desconectado
'   - InsertarOrdenAtomica(...)     → Alta de solicitud con Nº Orden atómico
'   - ActualizarOrden(...)          → UPDATE genérico por Nº Orden
'   - ObtenerSiguienteNumFactura()  → Incremento atómico del contador de facturas
'   - BuscarEnOrdenes(...)          → SELECT con filtro para búsquedas
'   - ComprobarListaNegra(dni)      → Verificar si un DNI está vetado
'   - InsertarLog(...)              → Escribir en LogActividad
'   - SincronizarHojaDesdeAccess()  → Volcar datos de Access a una hoja Excel
'   - ComprobarConexion()           → Test de conectividad con Access
' ==============================================================================

Private Const DB_PATH As String = "H:\ResidenciaBD\Residencia_BE.accdb"
Private Const PROVIDER_STRING As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="

' Constantes de estado de conexión ADO
Private Const adStateClosed As Long = 0
Private Const adStateOpen As Long = 1

' ===========================
' RESIDENCIA ASIGNADA AL USUARIO
' ===========================
' Cada usuario/operador está asignado a una residencia concreta.
' Esta variable se establece al arrancar el libro (Workbook_Open)
' y optimiza la sincronización para cargar solo los datos de su residencia.
Private m_ResidenciaAsignada As String

''' Establece la residencia asignada al usuario actual.
''' Se llama una vez al arrancar el libro, después del login.
''' residencia: "GIJON", "SOTO" o "OVIEDO"
Public Sub EstablecerResidenciaAsignada(ByVal residencia As String)
    m_ResidenciaAsignada = UCase(Trim(residencia))
End Sub

''' Retorna la residencia asignada al usuario actual.
Public Function ObtenerResidenciaAsignada() As String
    ObtenerResidenciaAsignada = m_ResidenciaAsignada
End Function

''' Solicita al usuario que seleccione su residencia asignada.
''' Se puede llamar desde Workbook_Open o desde frmLogin.
Public Sub SeleccionarResidencia()
    Dim respuesta As String
    respuesta = InputBox("Selecciona tu residencia asignada:" & vbCrLf & vbCrLf & _
                         "  1 = GIJÓN" & vbCrLf & _
                         "  2 = SOTO" & vbCrLf & _
                         "  3 = OVIEDO" & vbCrLf & vbCrLf & _
                         "Escribe 1, 2 o 3:", _
                         "Configuración de Residencia", "1")
    
    Select Case Trim(respuesta)
        Case "1", "GIJON", "GIJÓN"
            m_ResidenciaAsignada = "GIJON"
        Case "2", "SOTO"
            m_ResidenciaAsignada = "SOTO"
        Case "3", "OVIEDO"
            m_ResidenciaAsignada = "OVIEDO"
        Case Else
            m_ResidenciaAsignada = "GIJON" ' Por defecto
            MsgBox "Residencia no reconocida. Se asigna GIJÓN por defecto.", vbExclamation
    End Select
End Sub

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

''' Inserta un registro de Orden/Reserva de forma atómica y retorna el Nº ORDEN asignado.
''' El autonumérico de Access garantiza unicidad incluso con 6 usuarios simultáneos.
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
          ") VALUES (" & _
          "'" & EscaparSQL(residencia) & "', " & _
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
          "'" & EscaparSQL(Environ("USERNAME")) & "'" & _
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
''' camposYValores debe ser un diccionario (Scripting.Dictionary) con pares campo→valor.
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
' SINCRONIZACIÓN ACCESS → EXCEL
' ===========================

''' Descarga todos los registros de una residencia desde Access y los vuelca
''' en la hoja Excel correspondiente, sobreescribiendo los datos existentes.
'''
''' Esto permite que las macros de calendario, solapamientos e impresión
''' sigan funcionando sin cambios, ya que encuentran los datos en las celdas
''' habituales de la hoja.
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
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False  ' Evitar que Worksheet_Change se dispare durante la carga
    
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
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Not rs Is Nothing Then Set rs = Nothing
End Sub

' ===========================
' SINCRONIZACIÓN OPTIMIZADA AL ARRANQUE
' ===========================

''' Sincroniza SOLO la residencia asignada al usuario actual.
''' Se llama desde Workbook_Open después del login y la selección de residencia.
''' Es más rápida que sincronizar las 3 residencias porque solo descarga 1.
Public Sub SincronizarResidenciaAsignada()
    If Len(m_ResidenciaAsignada) = 0 Then
        MsgBox "No se ha establecido la residencia asignada.", vbExclamation
        Exit Sub
    End If
    
    Dim nombreHoja As String
    Select Case m_ResidenciaAsignada
        Case "GIJON": nombreHoja = "RESIDENCIA GIJÓN"
        Case "SOTO": nombreHoja = "RESIDENCIA SOTO"
        Case "OVIEDO": nombreHoja = "RESIDENCIA OVIEDO"
    End Select
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Application.StatusBar = "Sincronizando datos de " & nombreHoja & "..."
        SincronizarHojaDesdeAccess m_ResidenciaAsignada, ws
        Application.StatusBar = False
    End If
End Sub

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
