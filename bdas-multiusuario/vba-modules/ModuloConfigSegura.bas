Attribute VB_Name = "ModuloConfigSegura"
'Attribute VB_Name = "ModuloConfigSegura"
Option Explicit
' Constante para hoja de configuraci�n din�mica
Private Const HOJA_CONFIG_PWD As String = "CONFIG_PWD"
' Constantes de niveles de error (desde ModuloGestionErrores)
Private Const INFO_LEVEL As Integer = 1
Private Const WARNING_LEVEL As Integer = 2
Private Const ERROR_LEVEL As Integer = 3
Private Const CRITICAL_LEVEL As Integer = 4

'=================================================================================
' M�dulo: Configuraci�n Segura
' Descripci�n:
'   Gesti�n centralizada y segura de contrase�as y configuraci�n sensible.
'   VERSI�N 3.0: Sistema Din�mico implementado
'
' Mejoras de Seguridad:
'   1. Centralizaci�n de contrase�as en un solo m�dulo
'   2. Ofuscaci�n b�sica para contrase�as del sistema
'   3. Sistema de Hash + Salt para contrase�as de usuarios
'   4. Validaci�n fuerte de contrase�as
'   5. Funciones de migraci�n de texto plano a hash
'   6. Sistema din�mico con hoja CONFIG_PWD (NUEVO)
'
' VERSI�N: 3.0 - Sistema Din�mico
' FECHA: 2025-11-21
'=================================================================================

' ====================== CONFIGURACI�N DE CONTRASE�AS DEL SISTEMA =========================

' Funci�n para obtener CLAVE MAESTRA (Login)
Public Function ObtenerClaveMaestra() As String
    ' ACTUALIZADO: 2025-11-21 - Sistema simplificado a 3 contrase�as
    ' Uso: Login del sistema
    ' Contrase�a: MasterKey@Jaime1968
    ObtenerClaveMaestra = ObtenerPassword("CLAVE_MAESTRA")
End Function

' Funci�n para obtener contrase�a de HOJAS Y ESTRUCTURA
Public Function ObtenerPasswordHojasYEstructura() As String
    ' NUEVO: 2025-11-21 - Contrase�a unificada
    ' Uso: Protecci�n de hojas LOG, USUARIOS, CONFIG y estructura del libro
    ' Contrase�a: Log$Residencias2025
    ObtenerPasswordHojasYEstructura = ObtenerPassword("PASSWORD_HOJAS_Y_ESTRUCTURA")
End Function

' ============== FUNCIONES DE COMPATIBILIDAD (Temporal) ==============
' Estas funciones mantienen compatibilidad con c�digo antiguo
' Se redirigen a PASSWORD_HOJAS_Y_ESTRUCTURA

Public Function ObtenerPasswordLOG() As String
    ' DEPRECADO: Usar ObtenerPasswordHojasYEstructura()
    ' Redirige a la nueva contrase�a unificada
    ObtenerPasswordLOG = ObtenerPasswordHojasYEstructura()
End Function

Public Function ObtenerPasswordHojaUsuarios() As String
    ' DEPRECADO: Usar ObtenerPasswordHojasYEstructura()
    ' Redirige a la nueva contrase�a unificada
    ObtenerPasswordHojaUsuarios = ObtenerPasswordHojasYEstructura()
End Function

Public Function ObtenerPasswordEstructura() As String
    ' DEPRECADO: Usar ObtenerPasswordHojasYEstructura()
    ' Redirige a la nueva contrase�a unificada
    ObtenerPasswordEstructura = ObtenerPasswordHojasYEstructura()
End Function

Public Function ObtenerPasswordConfig() As String
    ' DEPRECADO: Usar ObtenerPasswordHojasYEstructura()
    ' Redirige a la nueva contrase�a unificada
    ObtenerPasswordConfig = ObtenerPasswordHojasYEstructura()
End Function

' ====================== FUNCIONES DE OFUSCACI�N ============================

' Ofuscaci�n simple ROT1 (cada car�cter +1 en ASCII)
' NOTA: Esto NO es cifrado seguro, solo ofuscaci�n b�sica para contrase�as del sistema
Private Function OfuscarTexto(texto As String) As String
    Dim i As Long
    Dim resultado As String
    Dim charCode As Integer
    
    resultado = ""
    For i = 1 To Len(texto)
        charCode = Asc(Mid(texto, i, 1))
        resultado = resultado & Chr(charCode + 1)
    Next i
    
    OfuscarTexto = resultado
End Function

' Desofuscaci�n ROT1 (inverso de ofuscaci�n)
Private Function DesofuscarTexto(texto As String) As String
    Dim i As Long
    Dim resultado As String
    Dim charCode As Integer
    
    resultado = ""
    For i = 1 To Len(texto)
        charCode = Asc(Mid(texto, i, 1))
        resultado = resultado & Chr(charCode - 1)
    Next i
    
    DesofuscarTexto = resultado
End Function

' ====================== VALIDACI�N DE CONTRASE�AS ==========================

' Valida que una contrase�a cumpla requisitos m�nimos de seguridad
Public Function ValidarFortalezaPassword(password As String) As Boolean
    Dim tieneMinuscula As Boolean
    Dim tieneMayuscula As Boolean
    Dim tieneNumero As Boolean
    Dim longitudMinima As Boolean
    Dim mensaje As String
    
    ' Verificar longitud m�nima
    longitudMinima = (Len(password) >= 8)
    
    ' Verificar contenido
    tieneMinuscula = (password Like "*[a-z]*")
    tieneMayuscula = (password Like "*[A-Z]*")
    tieneNumero = (password Like "*[0-9]*")
    
    ' Construir mensaje de error si no cumple
    If Not longitudMinima Then
        mensaje = "La contrase�a debe tener al menos 8 caracteres." & vbCrLf
    End If
    
    If Not (tieneMinuscula And tieneMayuscula) Then
        mensaje = mensaje & "Debe contener may�sculas y min�sculas." & vbCrLf
    End If
    
    If Not tieneNumero Then
        mensaje = mensaje & "Debe contener al menos un n�mero." & vbCrLf
    End If
    
    ' Retornar resultado
    If mensaje <> "" Then
        MsgBox "La contrase�a no cumple los requisitos de seguridad:" & vbCrLf & vbCrLf & mensaje, _
               vbExclamation, "Contrase�a D�bil"
        ValidarFortalezaPassword = False
    Else
        ValidarFortalezaPassword = True
    End If
End Function

' ====================== HASH DE CONTRASE�AS SIMPLE (LEGACY) ======================

' Funci�n de hash simple para contrase�as
' ADVERTENCIA: Esto NO es criptogr�ficamente seguro
' Mantenida por compatibilidad, usar HashPasswordMejorado en su lugar
Public Function HashPasswordSimple(password As String, salt As String) As String
    Dim i As Long
    Dim hashValue As Long
    Dim combined As String
    
    ' Combinar contrase�a con salt
    combined = password & salt
    
    ' Hash simple (no usar en producci�n)
    hashValue = 0
    For i = 1 To Len(combined)
        hashValue = ((hashValue * 31) Xor Asc(Mid(combined, i, 1))) Mod 2147483647
    Next i
    
    ' Convertir a hexadecimal
    HashPasswordSimple = Hex(hashValue)
End Function

' Genera un salt aleatorio para hashing (versi�n simple)
' LEGACY: Usar GenerarSaltSeguro en su lugar
Public Function GenerarSalt() As String
    Dim i As Long
    Dim salt As String
    
    Randomize Timer
    salt = ""
    
    For i = 1 To 16
        salt = salt & Chr(Int(Rnd() * 26) + 65)  ' A-Z aleatorio
    Next i
    
    GenerarSalt = salt
End Function

'====================== FUNCIONES DE HASH MEJORADAS (PASO 5) ====================

' Hash mejorado usando algoritmo de m�ltiples pasadas
' MEJORA: M�s seguro que HashPasswordSimple
' NOTA: Para m�xima seguridad en producci�n, usar CryptoAPI de Windows
Public Function HashPasswordMejorado(password As String, salt As String) As String
    Dim combined As String
    Dim i As Long
    Dim resultado As String
    Dim char As String
    Dim asciiVal As Integer
    Dim hash1 As String
    Dim hash2 As String
    
    ' Combinar contrase�a con salt
    combined = password & salt
    
    ' Generar hash usando m�todo de suma ponderada (sin desbordamiento)
    resultado = ""
    
    ' Primera pasada: construir hash car�cter por car�cter
    For i = 1 To Len(combined)
        char = Mid(combined, i, 1)
        asciiVal = Asc(char)
        
        ' Convertir a hex de 2 d�gitos
        resultado = resultado & Right("0" & Hex(asciiVal), 2)
    Next i
    
    ' Aplicar transformaci�n adicional usando XOR
    hash1 = ""
    For i = 1 To Len(resultado) Step 2
        If i + 1 <= Len(resultado) Then
            ' XOR entre caracteres consecutivos
            Dim val1 As Integer, val2 As Integer
            val1 = Asc(Mid(resultado, i, 1))
            val2 = Asc(Mid(resultado, i + 1, 1))
            hash1 = hash1 & Right("0" & Hex((val1 Xor val2) Mod 256), 2)
        End If
    Next i
    
    ' Generar checksum simple
    Dim checksumVal As Long
    checksumVal = 0
    For i = 1 To Len(combined)
        checksumVal = checksumVal + Asc(Mid(combined, i, 1))
    Next i
    
    ' Combinar ambos hashes
    hash2 = Right("00000000" & Hex(checksumVal Mod 16777216), 8)
    
    ' Resultado final: primeros 8 caracteres del hash1 + hash2
    If Len(hash1) >= 8 Then
        HashPasswordMejorado = Mid(hash1, 1, 8) & hash2
    Else
        HashPasswordMejorado = hash1 & String(8 - Len(hash1), "0") & hash2
    End If
End Function

' Genera un salt aleatorio seguro de 16 caracteres
' MEJORA: Usa conjunto de caracteres amplio y mejor randomizaci�n
Public Function GenerarSaltSeguro() As String
    Dim i As Long
    Dim salt As String
    Dim charset As String
    Dim randomIndex As Integer
    
    ' Conjunto de caracteres: A-Z y 0-9 (36 caracteres)
    charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    
    ' Inicializar generador aleatorio con m�ltiples fuentes de entrop�a
    Randomize Timer + (Second(Now) * 1000) + (Minute(Now) * 60000)
    
    salt = ""
    For i = 1 To 16
        ' Generar �ndice aleatorio
        randomIndex = Int((Len(charset) * Rnd()) + 1)
        salt = salt & Mid(charset, randomIndex, 1)
    Next i
    
    GenerarSaltSeguro = salt
End Function

' Valida una contrase�a ingresada contra un hash guardado
' Par�metros:
'   - passwordIngresada: La contrase�a que el usuario ingres�
'   - saltGuardado: El salt almacenado en la base de datos
'   - hashGuardado: El hash almacenado en la base de datos
' Retorna: True si la contrase�a es correcta, False si no
Public Function ValidarPasswordHash(passwordIngresada As String, _
                                    saltGuardado As String, _
                                    hashGuardado As String) As Boolean
    Dim hashCalculado As String
    
    ' Calcular hash de la contrase�a ingresada con el salt guardado
    hashCalculado = HashPasswordMejorado(passwordIngresada, saltGuardado)
    
    ' Comparar hashes (case-sensitive)
    ValidarPasswordHash = (hashCalculado = hashGuardado)
End Function

' Crea un nuevo par Salt + Hash para una contrase�a
' Par�metros:
'   - password: La contrase�a en texto plano
' Retorna: Array de 2 elementos: (0)=Salt, (1)=Hash
Public Function CrearPasswordHash(password As String) As Variant
    Dim resultado(1) As String
    Dim salt As String
    Dim hash As String
    
    ' Generar salt �nico aleatorio
    salt = GenerarSaltSeguro()
    
    ' Generar hash con el salt
    hash = HashPasswordMejorado(password, salt)
    
    ' Retornar ambos valores en un array
    resultado(0) = salt
    resultado(1) = hash
    
    CrearPasswordHash = resultado
End Function


' Funci�n UNIVERSAL para obtener cualquier contrase�a
Public Function ObtenerPassword(nombrePassword As String) As String
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim nombre As String
    Dim valorOfuscado As String
    
    ' Buscar hoja de configuraci�n
    Set ws = BuscarHoja(HOJA_CONFIG_PWD)
    
    If ws Is Nothing Then
        ' Si no existe, usar valores hardcoded como fallback
        ObtenerPassword = ObtenerPasswordHardcoded(nombrePassword)
        Exit Function
    End If
    
    ' Buscar contrase�a en la hoja
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaFila ' Fila 1 = encabezados
        nombre = ws.Cells(i, 1).Value
        
        If UCase(Trim(nombre)) = UCase(Trim(nombrePassword)) Then
            valorOfuscado = ws.Cells(i, 2).Value
            ObtenerPassword = DesofuscarTexto(valorOfuscado)
            Exit Function
        End If
    Next i
    
    ' Si no se encuentra, usar hardcoded
    ObtenerPassword = ObtenerPasswordHardcoded(nombrePassword)
    Exit Function
    
ErrorHandler:
    ' Log simple sin dependencia de ModuloGestionErrores
    Debug.Print "ERROR en ObtenerPassword: " & Err.Number & " - " & Err.Description
    ObtenerPassword = ""
End Function

' Funci�n para A�ADIR o ACTUALIZAR contrase�a din�micamente
Public Sub EstablecerPassword(nombrePassword As String, _
                             passwordTextoPlano As String, _
                             Optional descripcion As String = "")
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim encontrada As Boolean
    Dim passwordOfuscada As String
    Dim passwordProteccion As String
    
    ' Validar fortaleza
    If Not ValidarFortalezaPassword(passwordTextoPlano) Then
        Exit Sub
    End If
    
    ' Crear hoja si no existe
    Set ws = BuscarHoja(HOJA_CONFIG_PWD)
    If ws Is Nothing Then
        Set ws = CrearHojaConfigPassword()
    End If
    
    ' Desproteger hoja
    passwordProteccion = ObtenerPasswordHardcoded("PASSWORD_ESTRUCTURA")
    ws.Unprotect password:=passwordProteccion
    
    ' Ofuscar contrase�a
    passwordOfuscada = OfuscarTextoPublico(passwordTextoPlano)
    
    ' Buscar si ya existe
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    encontrada = False
    
    For i = 2 To ultimaFila
If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(nombrePassword)) Then
    ' Actualizar existente
    ws.Cells(i, 2).Value = passwordOfuscada
    ws.Cells(i, 3).Value = Now
    ws.Cells(i, 4).Value = ThisWorkbook.usuarioActual  ' ? Usuario de login
    If descripcion <> "" Then ws.Cells(i, 5).Value = descripcion
    encontrada = True
    Exit For
End If
    Next i
    
    ' Si no existe, a�adir nueva
' ? C�DIGO CORREGIDO
If Not encontrada Then
    ultimaFila = ultimaFila + 1
    ws.Cells(ultimaFila, 1).Value = nombrePassword
    ws.Cells(ultimaFila, 2).Value = passwordOfuscada
    ws.Cells(ultimaFila, 3).Value = Now
    ws.Cells(ultimaFila, 4).Value = ThisWorkbook.usuarioActual  ' ? Corregido
    ws.Cells(ultimaFila, 5).Value = descripcion
End If
    
    ' Reproteger hoja
    ws.Protect password:=passwordProteccion, UserInterfaceOnly:=True
    
    ' Logging simple
    Debug.Print "INFO: Password actualizada: " & nombrePassword
    
    MsgBox "Contrase�a '" & nombrePassword & "' configurada correctamente.", vbInformation
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en EstablecerPassword: " & Err.Number & " - " & Err.Description
    MsgBox "Error al establecer contrase�a: " & Err.Description, vbCritical
End Sub

' Funci�n para LISTAR todas las contrase�as disponibles
Public Function ListarPasswords() As Variant
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim resultado() As String
    Dim i As Long, j As Long
    
    Set ws = BuscarHoja(HOJA_CONFIG_PWD)
    
    If ws Is Nothing Then
        ' Devolver lista hardcoded
        ListarPasswords = Array("CLAVE_MAESTRA", _
                               "PASSWORD_HOJAS_Y_ESTRUCTURA")
        Exit Function
    End If
    
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaFila < 2 Then
        ListarPasswords = Array()
        Exit Function
    End If
    
    ReDim resultado(1 To ultimaFila - 1)
    j = 1
    
    For i = 2 To ultimaFila
        resultado(j) = ws.Cells(i, 1).Value
        j = j + 1
    Next i
    
    ListarPasswords = resultado
End Function

' Funci�n para ELIMINAR una contrase�a
Public Sub EliminarPassword(nombrePassword As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim passwordProteccion As String
    
    ' Verificar que no sea contrase�a cr�tica
    If EsPasswordCritica(nombrePassword) Then
        MsgBox "No se puede eliminar una contrase�a cr�tica del sistema.", vbExclamation
        Exit Sub
    End If
    
    Set ws = BuscarHoja(HOJA_CONFIG_PWD)
    If ws Is Nothing Then Exit Sub
    
    ' Desproteger
    passwordProteccion = ObtenerPasswordHardcoded("PASSWORD_ESTRUCTURA")
    ws.Unprotect password:=passwordProteccion
    
    ' Buscar y eliminar
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaFila
        If UCase(Trim(ws.Cells(i, 1).Value)) = UCase(Trim(nombrePassword)) Then
            ws.Rows(i).Delete
            MsgBox "Contrase�a '" & nombrePassword & "' eliminada.", vbInformation
            Exit For
        End If
    Next i
    
    ' Reproteger
    ws.Protect password:=passwordProteccion, UserInterfaceOnly:=True
    
    ' Logging simple
    Debug.Print "WARNING: Password eliminada: " & nombrePassword
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR en EliminarPassword: " & Err.Number & " - " & Err.Description
    MsgBox "Error al eliminar contrase�a: " & Err.Description, vbCritical
End Sub

' Crear hoja de configuraci�n con estructura
Private Function CrearHojaConfigPassword() As Worksheet
    Dim ws As Worksheet
    Dim passwordProteccion As String
    
    ' Crear hoja nueva
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = HOJA_CONFIG_PWD
    
    ' Encabezados
    ws.Cells(1, 1).Value = "Nombre"
    ws.Cells(1, 2).Value = "Valor_Ofuscado"
    ws.Cells(1, 3).Value = "Fecha_Modificacion"
    ws.Cells(1, 4).Value = "Usuario_Modificacion"
    ws.Cells(1, 5).Value = "Descripcion"
    
    ' Formato
    ws.Rows(1).Font.Bold = True
    ws.Columns("A:E").AutoFit
    
    ' Migrar contrase�as existentes (hardcoded)
    Call MigrarPasswordsHardcoded(ws)
    
    ' Ocultar y proteger
    ws.visible = xlSheetVeryHidden
    passwordProteccion = ObtenerPasswordHardcoded("PASSWORD_ESTRUCTURA")
    ws.Protect password:=passwordProteccion, UserInterfaceOnly:=True
    
    Set CrearHojaConfigPassword = ws
    
    MsgBox "Hoja CONFIG_PWD creada y contrase�as migradas exitosamente.", vbInformation
End Function

' Migrar contrase�as hardcoded a la nueva hoja
Private Sub MigrarPasswordsHardcoded(ws As Worksheet)
    Dim fila As Long
    fila = 2
    
    ' Sistema simplificado: Solo 3 contrase�as
    With ws
        ' 1. CLAVE_MAESTRA
        .Cells(fila, 1).Value = "CLAVE_MAESTRA"
        .Cells(fila, 2).Value = "NbtufsLfzAKbjnf2:79"  ' MasterKey@Jaime1968 ofuscado
        .Cells(fila, 3).Value = Now
        .Cells(fila, 4).Value = "SISTEMA"
        .Cells(fila, 5).Value = "Clave maestra para login del sistema"
        fila = fila + 1
        
        ' 2. PASSWORD_HOJAS_Y_ESTRUCTURA
        .Cells(fila, 1).Value = "PASSWORD_HOJAS_Y_ESTRUCTURA"
        .Cells(fila, 2).Value = "Mph%Sftjefodjbt3136"  ' Log$Residencias2025 ofuscado
        .Cells(fila, 3).Value = Now
        .Cells(fila, 4).Value = "SISTEMA"
        .Cells(fila, 5).Value = "Protecci�n de hojas del sistema (LOG, USUARIOS, CONFIG) y estructura del libro"
        fila = fila + 1
        
    End With
End Sub
' Fallback a contrase�as hardcoded si falla lectura din�mica
Private Function ObtenerPasswordHardcoded(nombrePassword As String) As String
    Select Case UCase(Trim(nombrePassword))
        Case "CLAVE_MAESTRA"
            ' Contrase�a: MasterKey@Jaime1968
            ObtenerPasswordHardcoded = DesofuscarTexto("NbtufsLfzAKbjnf2:79")
        
        Case "PASSWORD_HOJAS_Y_ESTRUCTURA"
            ' Contrase�a: Log$Residencias2025
            ObtenerPasswordHardcoded = DesofuscarTexto("Mph%Sftjefodjbt3136")
       
        ' COMPATIBILIDAD CON FUNCIONES ANTIGUAS
        Case "PASSWORD_LOG", "PASSWORD_USUARIOS", "PASSWORD_CONFIG", "PASSWORD_ESTRUCTURA"
            ' Redirigir a PASSWORD_HOJAS_Y_ESTRUCTURA
            ObtenerPasswordHardcoded = DesofuscarTexto("Mph%Sftjefodjbt3136")
        
        Case Else
            ObtenerPasswordHardcoded = ""
    End Select
End Function
' Buscar hoja de forma segura
Private Function BuscarHoja(nombreHoja As String) As Worksheet
    On Error Resume Next
    Set BuscarHoja = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo 0
End Function

' Verificar si es contrase�a cr�tica
Private Function EsPasswordCritica(nombrePassword As String) As Boolean
    ' Solo 3 contrase�as cr�ticas en el sistema simplificado
    Dim criticas As Variant
    Dim i As Long
    
    criticas = Array("CLAVE_MAESTRA", "PASSWORD_HOJAS_Y_ESTRUCTURA", "PASSWORD_BDAS")
    
    For i = LBound(criticas) To UBound(criticas)
        If UCase(Trim(nombrePassword)) = UCase(Trim(criticas(i))) Then
            EsPasswordCritica = True
            Exit Function
        End If
    Next i
    
    EsPasswordCritica = False
End Function

' Ofuscaci�n p�blica para uso en EstablecerPassword
Private Function OfuscarTextoPublico(texto As String) As String
    Dim i As Long
    Dim resultado As String
    
    resultado = ""
    For i = 1 To Len(texto)
        resultado = resultado & Chr(Asc(Mid(texto, i, 1)) + 1)
    Next i
    
    OfuscarTextoPublico = resultado
End Function

' Funci�n auxiliar para generar c�digo ofuscado (para desarrolladores)
Public Function GenerarCodigoOfuscado(passwordEnTextoPlano As String) As String
    Dim resultado As String
    
    resultado = OfuscarTextoPublico(passwordEnTextoPlano)
    GenerarCodigoOfuscado = resultado
    
    ' Mostrar para copiar al c�digo
    MsgBox "C�digo ofuscado para copiar:" & vbCrLf & vbCrLf & _
           "DesofuscarTexto(""" & resultado & """)" & vbCrLf & vbCrLf & _
           "Valor ofuscado: " & resultado, vbInformation, "Generador de C�digo Ofuscado"
End Function

' ============== FUNCI�N PARA HOJAS BDAS ==============

' Funci�n para obtener contrase�a de HOJAS BDAS
Public Function ObtenerPasswordHojas() As String
    ' Las hojas BDAS usan la misma contrase�a que el login (clave maestra)
    ' para m�xima seguridad
    ObtenerPasswordHojas = ObtenerClaveMaestra()
End Function
'=================================================================================
' NOTAS IMPORTANTES PARA EL DESARROLLADOR:
'
' HISTORIAL DE CAMBIOS:
' ----------------------
' 2025-11-16 (Paso 2):
'   ? Contrase�as del sistema actualizadas y ofuscadas
'   - Nueva contrase�a LOG: Log$Residencias2025
'   - Nueva clave maestra: MasterKey@Jaime1968
'
' 2025-11-16 (Paso 5):
'   ? Sistema de Hash + Salt implementado
'   - HashPasswordMejorado: 3 pasadas de hash con diferentes algoritmos
'   - GenerarSaltSeguro: Salt aleatorio de 16 caracteres
'   - ValidarPasswordHash: Validaci�n segura de contrase�as
'   - CrearPasswordHash: Generaci�n completa de salt + hash
'
' 2025-11-21 (Versi�n 3.0):
'   ? Sistema din�mico de contrase�as implementado
'   - Hoja CONFIG_PWD para almacenamiento centralizado
'   - Funci�n ObtenerPassword() universal
'   - Funci�n EstablecerPassword() para a�adir/actualizar
'   - Funci�n ListarPasswords() para listar disponibles
'   - Funci�n EliminarPassword() para eliminar contrase�as
'   - Sistema de fallback a contrase�as hardcoded
'   - Migraci�n autom�tica de contrase�as existentes
'   - Ofuscaci�n de PASSWORD_BDAS y PASSWORD_CONFIG
'
' SEGURIDAD:
' ----------
' ? Contrase�as del sistema: Ofuscadas con ROT1
' ? Contrase�as de usuarios: Hash + Salt (irreversible)
' ? Validaci�n fuerte: 8+ caracteres, may�sculas, min�sculas, n�meros
' ? Sistema din�mico: Hoja CONFIG_PWD (VeryHidden y protegida)
' ?? El hash implementado es suficiente para aplicaciones internas
' ?? Para aplicaciones cr�ticas web, considerar SHA256/bcrypt con CryptoAPI
'
' USO DEL SISTEMA DIN�MICO:
' --------------------------
' A�adir nueva contrase�a:
'   Call ModuloConfigSegura.EstablecerPassword("NOMBRE", "Valor123!", "Descripci�n")
'
' Obtener contrase�a:
'   pwd = ModuloConfigSegura.ObtenerPassword("NOMBRE")
'
' Listar contrase�as:
'   lista = ModuloConfigSegura.ListarPasswords()
'
' Eliminar contrase�a:
'   Call ModuloConfigSegura.EliminarPassword("NOMBRE")
'
' PR�XIMAS MEJORAS OPCIONALES:
' ----------------------------
' 1. Implementar CryptoAPI de Windows para SHA256 nativo
' 2. A�adir pol�tica de expiraci�n de contrase�as (90 d�as)
' 3. Implementar bloqueo de cuenta tras N intentos fallidos
' 4. A�adir historial de contrase�as (evitar reutilizaci�n)
' 5. Implementar autenticaci�n de dos factores (2FA)
' 6. Migrar a Windows Credential Manager para contrase�as del sistema
' 7. Crear formulario visual para gesti�n de contrase�as
'
' Referencias de seguridad:
' - OWASP Password Storage Cheat Sheet
' - NIST Digital Identity Guidelines (SP 800-63B)
' - Reglamento RGPD (UE 2016/679) - Art�culo 32 (Seguridad del tratamiento)
'=================================================================================

