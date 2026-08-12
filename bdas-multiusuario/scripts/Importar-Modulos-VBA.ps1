# ==============================================================================
# Script: Importar-Modulos-VBA.ps1
# Propósito: Importa automáticamente los módulos modDatabase.bas, modMigracion.bas,
#            los UserForms frmLogin.frm y frmSelectorResidencia.frm, y los eventos
#            de ThisWorkbook en un archivo Excel (.xlsm) destino.
#
# Requisito: Excel debe tener activada la opción:
#   "Confiar en el acceso al modelo de objetos de proyectos de VBA"
#   (Archivo -> Opciones -> Centro de confianza -> Configuración del Centro de confianza -> Configuración de macros)
# ==============================================================================

Param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelPath,
    [string]$ScriptsDir = "H:\ResidenciaApp\bdas-multiusuario\scripts"
)

$ErrorActionPreference = "Stop"

# Habilitar automatica y temporalmente AccessVBOM en el Registro de Windows
foreach ($regPath in @("HKCU:\Software\Microsoft\Office\16.0\Excel\Security", "HKCU:\Software\WOW6432Node\Microsoft\Office\16.0\Excel\Security")) {
    try {
        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
        Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    } catch {}
}



if (-not (Test-Path -Path $ExcelPath)) {
    Write-Error "El archivo Excel '$ExcelPath' no existe."
    exit 1
}

$modDatabasePath = Join-Path $ScriptsDir "modDatabase.bas"
$modMigracionPath = Join-Path $ScriptsDir "modMigracion.bas"
$frmLoginPath = Join-Path $ScriptsDir "frmLogin.frm"
$frmSelectorPath = Join-Path $ScriptsDir "frmSelectorResidencia.frm"
$thisWorkbookEventsPath = Join-Path $ScriptsDir "ThisWorkbook_Events.bas"

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  BDAS - Importador Automatizado de Módulos VBA               " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Destino Excel: $ExcelPath" -ForegroundColor White

# Habilitar automatica y temporalmente AccessVBOM en el Registro de Windows
foreach ($regPath in @("HKCU:\Software\Microsoft\Office\16.0\Excel\Security", "HKCU:\Software\WOW6432Node\Microsoft\Office\16.0\Excel\Security")) {
    try {
        if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
        Set-ItemProperty -Path $regPath -Name "AccessVBOM" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    } catch {}
}

# Definir P/Invoke para simular tecla Shift al abrir el libro y evitar ejecucion de macros de inicio
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class WinKeyboard {
    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
}
"@ -ErrorAction SilentlyContinue
} catch {}

$excel = $null
$wb = $null

try {
    Write-Host "[PROCESANDO] Abriendo Microsoft Excel..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.WindowState = -4137 # xlMaximized
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false

    # Simular pulsacion de tecla Shift (VK_SHIFT = 0x10) para omitir macros Workbook_Open
    if ("WinKeyboard" -as [type]) {
        [WinKeyboard]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero)
    }

    try {
        $wb = $excel.Workbooks.Open($ExcelPath, 0, $false)
    } finally {
        if ("WinKeyboard" -as [type]) {
            [WinKeyboard]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)
        }
    }

    $vbProj = $wb.VBProject

    function Import-OrUpdate-VBComponent($proj, $compName, $filePath) {
        $comp = $null
        try { $comp = $proj.VBComponents.Item($compName) } catch {}
        $code = Get-Content $filePath -Raw -Encoding utf8
        if ($comp -ne $null) {
            Write-Host "[INFO] Actualizando contenido del módulo existente '$compName'..." -ForegroundColor Yellow
            if ($comp.CodeModule.CountOfLines -gt 0) {
                $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines)
            }
            $comp.CodeModule.AddFromString($code)
        } else {
            Write-Host "[INFO] Importando módulo '$compName'..." -ForegroundColor Green
            [void]$proj.VBComponents.Import($filePath)
        }
    }

    # 1. Importar/Actualizar modDatabase
    Import-OrUpdate-VBComponent $vbProj "modDatabase" $modDatabasePath
    Write-Host "[OK] Módulo 'modDatabase' procesado." -ForegroundColor Green

    # 2. Importar/Actualizar modMigracion
    Import-OrUpdate-VBComponent $vbProj "modMigracion" $modMigracionPath
    Write-Host "[OK] Módulo 'modMigracion' procesado." -ForegroundColor Green

    # 3. Importar frmLogin (UserForm)
    try {
        $c = $vbProj.VBComponents.Item("frmLogin")
        if ($c) { $vbProj.VBComponents.Remove($c) }
    } catch {}
    [void]$vbProj.VBComponents.Import($frmLoginPath)
    Write-Host "[OK] UserForm 'frmLogin' importado." -ForegroundColor Green

    # 4. Importar frmSelectorResidencia (UserForm)
    try {
        $c = $vbProj.VBComponents.Item("frmSelectorResidencia")
        if ($c) { $vbProj.VBComponents.Remove($c) }
    } catch {}
    [void]$vbProj.VBComponents.Import($frmSelectorPath)
    Write-Host "[OK] UserForm 'frmSelectorResidencia' importado." -ForegroundColor Green

    # 4.5. Re-importar/Actualizar los 15 módulos VBA adaptados
    $vbaDir = Join-Path (Split-Path $ScriptsDir) "vba-modules"
    if (Test-Path $vbaDir) {
        $modsToImport = @("FechasPeticion", "EvitarDuplicidadSolicitudesGyS", "AsignarNumFactura", `
                          "FacturacionMesGIJON", "FacturacionMesOVIEDO", "FacturacionMesSOTO", `
                          "MarcarSiPagadosEnResidencia", "ModuloCalendarioGijon", `
                          "ModuloCalendarioOviedo", "ModuloCalendarioSoto", `
                          "BusquedaDNIResidencias", "BusquedaOrdenNombreFactura", `
                          "ModuloLOG", "ModListaNegra", "ReevaluacionSolicitudes")
        foreach ($mName in $modsToImport) {
            $mPath = Join-Path $vbaDir "$mName.bas"
            if (Test-Path $mPath) {
                Import-OrUpdate-VBComponent $vbProj $mName $mPath
                Write-Host "[OK] Módulo adaptado '$mName' procesado." -ForegroundColor Green
            }
        }
    }

    # 5. Inyectar eventos en ThisWorkbook
    $thisWorkbookComp = $vbProj.VBComponents.Item("ThisWorkbook")
    $codeModule = $thisWorkbookComp.CodeModule

    if (Test-Path $thisWorkbookEventsPath) {
        $eventsCode = Get-Content $thisWorkbookEventsPath -Raw -Encoding utf8
        if ($codeModule.CountOfLines -gt 0) {
            $codeModule.DeleteLines(1, $codeModule.CountOfLines)
        }
        $codeModule.AddFromString($eventsCode)
        Write-Host "[OK] Código de eventos inyectado en 'ThisWorkbook'." -ForegroundColor Green
    }

    Write-Host "[PROCESANDO] Guardando cambios en el libro Excel..." -ForegroundColor Cyan
    $wb.Save()
    Write-Host "[OK] Libro Excel guardado correctamente." -ForegroundColor Green

} catch {
    Write-Error "Error durante la importacion de modulos VBA: $_`nAsegurate de que 'Confiar en el acceso al modelo de objetos de proyectos de VBA' este activado en Excel."
} finally {
    if ($wb) { $wb.Close($false); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if ($excel) { $excel.Quit(); [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "  IMPORTACION FINALIZADA CON EXITO                             " -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
