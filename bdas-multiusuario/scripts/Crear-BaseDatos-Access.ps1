# ==============================================================================
# Script: Crear-BaseDatos-Access.ps1
# Propósito: Crea la base de datos Microsoft Access (Residencia_BE.accdb)
#            en H:\ResidenciaBD\ con TODAS las tablas necesarias para la
#            solución puente multiusuario del BDAS v16.5.5 (hasta 15 operadores).
#
# Tablas creadas:
#   - Ordenes (unifica RESIDENCIA GIJON + SOTO + OVIEDO)
#   - ContadorFacturas (numeración atómica de facturas por residencia)
#   - LogActividad (reemplaza hoja LOG)
#   - ListaNegra (reemplaza hoja LISTA NEGRA)
#   - Usuarios (gestión de operadores y residencias asignadas)
#
# Uso:
#   powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\bdas-multiusuario\scripts\Crear-BaseDatos-Access.ps1" -StartingNumOrden 1541
# ==============================================================================

Param(
    [string]$TargetFolder = "H:\ResidenciaBD",
    [string]$DbName = "Residencia_BE.accdb",
    [int]$StartingNumOrden = 1,
    [int]$UltimaFacturaGijon = 0,
    [int]$UltimaFacturaSoto = 0,
    [int]$UltimaFacturaOviedo = 0,
    [int]$Ejercicio = (Get-Date).Year,
    [switch]$ForceRecreate = $false
)

$ErrorActionPreference = "Stop"

# Auto-reejecucion en PowerShell 32-bit si la sesion actual es 64-bit (el proveedor OLEDB ACE es 32-bit)
if ([Environment]::Is64BitProcess -and (Test-Path "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe")) {
    Write-Host "[INFO] Reejecutando script en PowerShell 32-bit para soporte OLEDB ACE Access..." -ForegroundColor Yellow
    & "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy Bypass -File $MyInvocation.MyCommand.Path @PSBoundParameters
    exit $LASTEXITCODE
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  BDAS - Inicializador de Base de Datos Access (H:\)           " -ForegroundColor Cyan
Write-Host "  Solucion puente multiusuario para hasta 15 operadores        " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# -----------------------------------------------------------------------------
# 1. Crear directorio de destino
# -----------------------------------------------------------------------------
if (-not (Test-Path -Path $TargetFolder)) {
    New-Item -Path $TargetFolder -ItemType Directory | Out-Null
    Write-Host "[OK] Directorio '$TargetFolder' creado." -ForegroundColor Green
} else {
    Write-Host "[INFO] El directorio '$TargetFolder' ya existe." -ForegroundColor Yellow
}

# Crear subdirectorio para la plantilla del Front-End
$plantillaDir = Join-Path -Path $TargetFolder -ChildPath "Plantilla"
if (-not (Test-Path -Path $plantillaDir)) {
    New-Item -Path $plantillaDir -ItemType Directory | Out-Null
    Write-Host "[OK] Directorio de plantilla '$plantillaDir' creado." -ForegroundColor Green
}

$dbPath = Join-Path -Path $TargetFolder -ChildPath $DbName

if (Test-Path -Path $dbPath) {
    if ($ForceRecreate) {
        Write-Host "[INFO] Recreando base de datos (-ForceRecreate)..." -ForegroundColor Yellow
        Remove-Item -Path $dbPath -Force
    } else {
        Write-Host "[AVISO] La base de datos '$dbPath' ya existe." -ForegroundColor Red
        Write-Host "        No se sobrescribira para preservar datos existentes." -ForegroundColor Red
        Write-Host "        Si desea recrearla, elimine el archivo manualmente o use -ForceRecreate." -ForegroundColor Red
        exit 0
    }
}

Write-Host "[PROCESANDO] Creando base de datos en: $dbPath ..." -ForegroundColor Cyan

# -----------------------------------------------------------------------------
# 2. Crear archivo .accdb mediante ADOX (ComObject)
# -----------------------------------------------------------------------------
try {
    $cat = New-Object -ComObject ADOX.Catalog
    $connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
    $cat.Create($connectionString)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($cat) | Out-Null
    Write-Host "[OK] Archivo .accdb creado correctamente." -ForegroundColor Green
} catch {
    Write-Error "Error al crear la base de datos Access.`nAsegurate de tener instalado Microsoft Access o los controladores redistribuibles de Access Database Engine.`nDescarga: https://www.microsoft.com/en-us/download/details.aspx?id=54920`nDetalle: $_"
    exit 1
}

# -----------------------------------------------------------------------------
# 3. Crear tablas mediante ADODB Connection
# -----------------------------------------------------------------------------
try {
    $cn = New-Object -ComObject ADODB.Connection
    $cn.Open($connectionString)

    # =========================================================================
    # TABLA: Ordenes
    # Unifica las hojas RESIDENCIA GIJON + SOTO + OVIEDO
    # El campo NumOrden es AUTOINCREMENT (autonumerico atomico)
    # =========================================================================
    Write-Host "[PROCESANDO] Creando tabla Ordenes..." -ForegroundColor Cyan

    $sqlOrdenes = @"
CREATE TABLE Ordenes (
    Id AUTOINCREMENT PRIMARY KEY,
    NumOrden LONG NOT NULL,
    NumFactura LONG,
    Residencia VARCHAR(20) NOT NULL,
    FechaPeticion DATETIME,
    DNI VARCHAR(20),
    Nombre VARCHAR(100),
    Apellidos VARCHAR(100),
    TipoHuesped VARCHAR(30),
    NumHabIndividuales INTEGER,
    NumHabDobles INTEGER,
    FechaEntrada DATETIME,
    FechaSalida DATETIME,
    Resolucion VARCHAR(30),
    HabitacionesAsignadas VARCHAR(200),
    EstadoPago VARCHAR(30),
    Solapamiento BIT,
    Observaciones MEMO,
    Telefono VARCHAR(20),
    Email VARCHAR(100),
    Direccion VARCHAR(200),
    CodigoPostal VARCHAR(10),
    Poblacion VARCHAR(100),
    Provincia VARCHAR(50),
    ConsentimientoRGPD BIT,
    PagadoResidencia VARCHAR(30),
    FechaCreacion DATETIME,
    UsuarioCreacion VARCHAR(50)
);
"@
    [void]$cn.Execute($sqlOrdenes)
    Write-Host "[OK] Tabla 'Ordenes' creada." -ForegroundColor Green

    # Ajustar valor de inicio del autonumerico
    if ($StartingNumOrden -gt 1) {
        try {
            $alterSql = "ALTER TABLE Ordenes ALTER COLUMN NumOrden COUNTER($StartingNumOrden, 1);"
            [void]$cn.Execute($alterSql)
            Write-Host "[OK] Semilla de NumOrden fijada en $StartingNumOrden." -ForegroundColor Green
        } catch {
            Write-Host "[AVISO] No se pudo establecer la semilla del autonumerico directamente." -ForegroundColor Yellow
            Write-Host "        Se insertara un registro semilla y se eliminara." -ForegroundColor Yellow
        }
    }

    # Crear indices utiles para rendimiento
    [void]$cn.Execute("CREATE INDEX idx_Ordenes_Residencia ON Ordenes (Residencia);")
    [void]$cn.Execute("CREATE INDEX idx_Ordenes_DNI ON Ordenes (DNI);")
    [void]$cn.Execute("CREATE INDEX idx_Ordenes_FechaEntrada ON Ordenes (FechaEntrada);")
    [void]$cn.Execute("CREATE INDEX idx_Ordenes_FechaSalida ON Ordenes (FechaSalida);")
    [void]$cn.Execute("CREATE INDEX idx_Ordenes_Resolucion ON Ordenes (Resolucion);")
    Write-Host "[OK] Indices de Ordenes creados." -ForegroundColor Green

    # =========================================================================
    # TABLA: ContadorFacturas
    # Numeracion atomica de facturas, independiente por residencia y ejercicio
    # Reemplaza la logica de bloques de AsignarNumFactura.bas
    # =========================================================================
    Write-Host "[PROCESANDO] Creando tabla ContadorFacturas..." -ForegroundColor Cyan

    $sqlContadores = @"
CREATE TABLE ContadorFacturas (
    Id AUTOINCREMENT PRIMARY KEY,
    Residencia VARCHAR(20) NOT NULL,
    Ejercicio INTEGER NOT NULL,
    UltimoNumero LONG NOT NULL
);
"@
    [void]$cn.Execute($sqlContadores)
    [void]$cn.Execute("CREATE UNIQUE INDEX idx_CF_ResEjer ON ContadorFacturas (Residencia, Ejercicio);")
    Write-Host "[OK] Tabla 'ContadorFacturas' creada." -ForegroundColor Green

    # Insertar registros iniciales para las 3 residencias
    [void]$cn.Execute("INSERT INTO ContadorFacturas (Residencia, Ejercicio, UltimoNumero) VALUES ('GIJON', $Ejercicio, $UltimaFacturaGijon);")
    [void]$cn.Execute("INSERT INTO ContadorFacturas (Residencia, Ejercicio, UltimoNumero) VALUES ('SOTO', $Ejercicio, $UltimaFacturaSoto);")
    [void]$cn.Execute("INSERT INTO ContadorFacturas (Residencia, Ejercicio, UltimoNumero) VALUES ('OVIEDO', $Ejercicio, $UltimaFacturaOviedo);")
    Write-Host "[OK] Contadores de factura inicializados (GIJ=$UltimaFacturaGijon, SOT=$UltimaFacturaSoto, OVI=$UltimaFacturaOviedo)." -ForegroundColor Green

    # =========================================================================
    # TABLA: LogActividad
    # Reemplaza la hoja LOG del Excel
    # =========================================================================
    Write-Host "[PROCESANDO] Creando tabla LogActividad..." -ForegroundColor Cyan

    $sqlLog = @"
CREATE TABLE LogActividad (
    Id AUTOINCREMENT PRIMARY KEY,
    FechaHora DATETIME NOT NULL,
    Usuario VARCHAR(50),
    Accion VARCHAR(200),
    Detalle MEMO
);
"@
    [void]$cn.Execute($sqlLog)
    [void]$cn.Execute("CREATE INDEX idx_Log_FechaHora ON LogActividad (FechaHora);")
    [void]$cn.Execute("CREATE INDEX idx_Log_Usuario ON LogActividad (Usuario);")
    Write-Host "[OK] Tabla 'LogActividad' creada." -ForegroundColor Green

    # =========================================================================
    # TABLA: ListaNegra
    # Reemplaza la hoja LISTA NEGRA del Excel
    # =========================================================================
    Write-Host "[PROCESANDO] Creando tabla ListaNegra..." -ForegroundColor Cyan

    $sqlListaNegra = @"
CREATE TABLE ListaNegra (
    Id AUTOINCREMENT PRIMARY KEY,
    DNI VARCHAR(20) NOT NULL,
    Nombre VARCHAR(150),
    Motivo MEMO,
    FechaAlta DATETIME,
    Activo BIT NOT NULL
);
"@
    [void]$cn.Execute($sqlListaNegra)
    [void]$cn.Execute("CREATE UNIQUE INDEX idx_LN_DNI ON ListaNegra (DNI);")
    Write-Host "[OK] Tabla 'ListaNegra' creada." -ForegroundColor Green

    # =========================================================================
    # TABLA: Usuarios
    # Gestion de operadores y sus residencias asignadas.
    # Hasta 15 usuarios simultaneos, cada uno asignado a 1-3 residencias.
    # Login manual con usuario + contrasena (hash SHA-256).
    # =========================================================================
    Write-Host "[PROCESANDO] Creando tabla Usuarios..." -ForegroundColor Cyan

    $sqlUsuarios = @"
CREATE TABLE Usuarios (
    Id AUTOINCREMENT PRIMARY KEY,
    NombreUsuario VARCHAR(50) NOT NULL,
    Clave VARCHAR(128) NOT NULL,
    NombreCompleto VARCHAR(100),
    Residencias VARCHAR(60) NOT NULL,
    Rol VARCHAR(20) DEFAULT 'OPERADOR',
    Activo BIT NOT NULL DEFAULT True,
    FechaAlta DATETIME
);
"@
    [void]$cn.Execute($sqlUsuarios)
    [void]$cn.Execute("CREATE UNIQUE INDEX idx_Usr_NombreUsuario ON Usuarios (NombreUsuario);")
    Write-Host "[OK] Tabla 'Usuarios' creada." -ForegroundColor Green

    # Funcion para generar hash SHA-256 (misma logica que en VBA para compatibilidad)
    function Get-SHA256Hash([string]$text) {
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
        $hash = $sha256.ComputeHash($bytes)
        return ($hash | ForEach-Object { $_.ToString("x2") }) -join ""
    }

    # Insertar un usuario administrador de ejemplo con clave "admin" (CAMBIAR tras despliegue)
    $adminUser = "admin"
    $adminClave = Get-SHA256Hash "admin"
    [void]$cn.Execute("INSERT INTO Usuarios (NombreUsuario, Clave, NombreCompleto, Residencias, Rol, Activo, FechaAlta) VALUES ('$adminUser', '$adminClave', 'Administrador', 'GIJON,SOTO,OVIEDO', 'ADMIN', True, Now());")
    Write-Host "[OK] Usuario administrador '$adminUser' creado (clave: admin - CAMBIAR)." -ForegroundColor Green
    Write-Host "     Hash SHA-256: $adminClave" -ForegroundColor DarkGray

    # Cerrar conexion
    $cn.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($cn) | Out-Null

} catch {
    Write-Error "Error al crear las tablas en Access: $_"
    if ($cn -and $cn.State -eq 1) { $cn.Close() }
    exit 1
}

# -----------------------------------------------------------------------------
# 4. Crear archivo de version
# -----------------------------------------------------------------------------
$versionFile = Join-Path -Path $TargetFolder -ChildPath "version.txt"
$nowDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
"1.0.0|$nowDate|Creacion inicial de la BD multiusuario" | Out-File -FilePath $versionFile -Encoding utf8
Write-Host "[OK] Archivo de version creado: $versionFile" -ForegroundColor Green

# -----------------------------------------------------------------------------
# 5. Resumen final
# -----------------------------------------------------------------------------
Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  BASE DE DATOS CREADA CON EXITO                               " -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Ubicacion:        $dbPath" -ForegroundColor White
Write-Host "  Tablas creadas:   Ordenes, ContadorFacturas, LogActividad, ListaNegra, Usuarios" -ForegroundColor White
Write-Host "  NumOrden desde:   $StartingNumOrden" -ForegroundColor White
Write-Host "  Facturas desde:   GIJ=$UltimaFacturaGijon SOT=$UltimaFacturaSoto OVI=$UltimaFacturaOviedo" -ForegroundColor White
Write-Host "  Ejercicio:        $Ejercicio" -ForegroundColor White
Write-Host ""
Write-Host "  Siguiente paso:   Importar modDatabase.bas en el Excel" -ForegroundColor Yellow
Write-Host "                    (ALT+F11 -> Archivo -> Importar archivo)" -ForegroundColor Yellow
Write-Host ""
