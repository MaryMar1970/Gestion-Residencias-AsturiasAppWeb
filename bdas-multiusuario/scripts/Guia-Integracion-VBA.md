# Guía Paso a Paso: Integración de Excel Front-End con MS Access (Unidad H:)

Esta guía explica cómo adaptar el archivo Excel actual para conectarlo a la base de datos Microsoft Access compartida en la unidad `H:\`.

---

## Paso 1: Crear la Base de Datos en `H:\`

Puedes ejecutar el script de PowerShell incluido en el proyecto (`scripts/Crear-BaseDatos-Access.ps1`) o ejecutar el siguiente comando en consola de PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\scripts\Crear-BaseDatos-Access.ps1" -StartingNumOrden 1000
```

Esto creará automáticamente el archivo `H:\ResidenciaBD\Residencia_BE.accdb` con la tabla `Ordenes` y el campo `NumOrden` configurado como autonumérico atómico.

---

## Paso 2: Importar Módulos y UserForms en el Excel Front-End (Fase 2)

### 2.1. Importación de Módulos Estándar (`.bas`) y UserForms (`.frm`)
1. Abre tu archivo Excel maestro (`.xlsm`).
2. Presiona `ALT + F11` para abrir el editor de Visual Basic para Aplicaciones (VBA).
3. Haz clic con el botón derecho sobre el proyecto en el árbol de navegación (izquierda) y selecciona **Importar archivo...** (o pulsa `Ctrl + M`).
4. Importa uno a uno los siguientes 4 archivos ubicados en `H:\ResidenciaApp\bdas-multiusuario\scripts\`:
   - `modDatabase.bas` — Capa de acceso a datos ADO + Autenticación SHA-256 + Sincronización.
   - `modMigracion.bas` — Módulo para la migración inicial de datos históricos a Access.
   - `frmLogin.frm` — Formulario visual de login (usuario + contraseña).
   - `frmSelectorResidencia.frm` — Formulario visual de selección de residencia activa.

> [!TIP]
> Al importar `frmLogin.frm` y `frmSelectorResidencia.frm`, Excel asociará automáticamente sus archivos de diseño `.frx` correspondientes.

### 2.2. Configuración de Eventos de Inicio y Cierre en `ThisWorkbook`
1. En el árbol de navegación del Editor VBA (panel izquierdo), haz doble clic sobre el objeto **`ThisWorkbook`**.
2. Abre el archivo `H:\ResidenciaApp\bdas-multiusuario\scripts\ThisWorkbook_Events.bas` en el Bloc de notas o editor.
3. Copia todo su contenido y pégalo en el módulo de código de **`ThisWorkbook`**.
4. Este código gestiona:
   - `Workbook_Open`: Comprueba la conexión a Access en `H:\`, muestra el formulario de login `frmLogin` y `frmSelectorResidencia`, y sincroniza la residencia activa.
   - `Workbook_BeforeClose`: Registra la salida (`LOGOUT`) y evita que el archivo `.xlsm` pregunte si guardar cambios (Access es la fuente de verdad).

### 2.3. Verificación de Compilación
1. En la barra superior del Editor VBA, haz clic en **Depuración** -> **Compilar VBAProject**.
2. Guarda el libro como **`BDAS_v16.5.5_FrontEnd.xlsm`** (o `Residencia_FrontEnd.xlsm`).

---

## Paso 3: Adaptar las Macros Existentes

### 3.1. Reemplazar la Macro de "Grabar Nueva Reserva"

Busca la macro o evento del botón **Guardar** (en tu formulario `UserForm` o módulo estándar).

#### Código Antiguo (Excel en pestañas):
```vba
' ANTIGUO: Escribía directamente en las celdas
Dim ultFila As Long
ultFila = Sheets("Datos").Cells(Rows.Count, 1).End(xlUp).Row + 1

' Generaba el Nº Orden de forma manual (propenso a duplicados)
Dim nuevoOrden As Long
nuevoOrden = Sheets("Datos").Cells(ultFila - 1, 1).Value + 1 

Sheets("Datos").Cells(ultFila, 1).Value = nuevoOrden
Sheets("Datos").Cells(ultFila, 2).Value = txtHuesped.Text
Sheets("Datos").Cells(ultFila, 3).Value = txtNIF.Text
'...
```

#### Código Nuevo (VBA + Access vía ADO):
```vba
' NUEVO: Invoca la función atómica del módulo modDatabase
Dim asignadoNumOrden As Long

asignadoNumOrden = modDatabase.InsertarOrdenAtomica( _
    huespedId:=txtHuespedId.Text, _
    nombreHuesped:=txtHuesped.Text, _
    nif:=txtNIF.Text, _
    residencia:=cmbResidencia.Text, _
    tipoHabitacion:=cmbTipoHabitacion.Text, _
    fechaEntrada:=CDate(txtFechaEntrada.Text), _
    fechaSalida:=CDate(txtFechaSalida.Text), _
    estado:="Pendiente", _
    resolucion:="", _
    observaciones:=txtObservaciones.Text _
)

If asignadoNumOrden > 0 Then
    MsgBox "Reserva grabada con éxito." & vbCrLf & _
           "Nº ORDEN Asignado: " & asignadoNumOrden, vbInformation, "Registro Completado"
Else
    MsgBox "No se pudo grabar la reserva. Inténtelo de nuevo.", vbExclamation, "Error"
End If
```

---

### 3.2. Cargar Listados o Buscar Datos desde Access

Para mostrar listados en formularios o volcarlos en pestañas de vista:

```vba
Sub CargarListadoEnHoja()
    Dim rs As Object
    Dim sql As String
    
    sql = "SELECT NumOrden, NombreHuesped, Residencia, FechaEntrada, FechaSalida, Estado FROM Ordenes ORDER BY NumOrden DESC"
    
    Set rs = modDatabase.GetRecordset(sql)
    
    If Not rs Is Nothing Then
        ' Limpiar contenido antiguo en la hoja de vistas
        Sheets("VistaOrdenes").Range("A5:F10000").ClearContents
        
        ' Volcar el Recordset directamente a las celdas
        Sheets("VistaOrdenes").Range("A5").CopyFromRecordset rs
    End If
End Sub
```

---

## Paso 4: Despliegue a los 6 Usuarios

1. Guarda una copia limpia del Excel como **`Residencia_FrontEnd.xlsm`**.
2. Copia este archivo `Residencia_FrontEnd.xlsm` en el escritorio local de cada uno de los 6 ordenadores.
3. ¡Listo! Todos los ordenadores trabajarán localmente contra la base de datos central en `H:\ResidenciaBD\Residencia_BE.accdb` simultáneamente sin bloqueos ni duplicados de Nº ORDEN.
