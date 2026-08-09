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

## Paso 2: Importar el Módulo VBA en el Excel Front-End

1. Abre tu archivo Excel maestro.
2. Presiona `ALT + F11` para abrir el editor de Visual Basic para Aplicaciones (VBA).
3. Haz clic con el botón derecho sobre el árbol de proyectos del libro (izquierda) y selecciona **Importar archivo...** (`Ctrl + M`).
4. Selecciona el archivo `H:\ResidenciaApp\scripts\modDatabase.bas`.
5. Verás aparecer un nuevo módulo llamado `modDatabase` con todas las funciones de conexión necesarias.

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
