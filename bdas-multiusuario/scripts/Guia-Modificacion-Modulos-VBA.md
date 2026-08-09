# Guía de Modificación de Módulos VBA — BDAS Multiusuario

Guía paso a paso para adaptar los 15 módulos VBA críticos del BDAS v16.5.5.
Cada sección indica **qué buscar** en el código existente y **por qué sustituirlo**.

> **REGLA DE ORO:** Cada módulo se modifica para escribir **solo en Access** (fuente de verdad).
> Después de cada escritura, se llama a `modDatabase.RefrescarCacheVisual` para actualizar
> las celdas en memoria (caché visual para calendarios, solapamientos e impresión).
> **NO se usa doble escritura** — Access es la única fuente de verdad.

---

## PREVIO: Importar los módulos de soporte

Antes de modificar cualquier módulo, importar en el proyecto VBA (ALT+F11 → Archivo → Importar):

1. `H:\ResidenciaApp\bdas-multiusuario\scripts\modDatabase.bas` — Capa de acceso a datos + login
2. `H:\ResidenciaApp\bdas-multiusuario\scripts\modMigracion.bas` — Migración inicial (se usará una vez)

Además, crear los UserForms siguiendo las instrucciones en:
3. `H:\ResidenciaApp\bdas-multiusuario\scripts\frmLogin.frm` — Login manual
4. `H:\ResidenciaApp\bdas-multiusuario\scripts\frmSelectorResidencia.frm` — Selector de residencia

---

## 1. FechasPeticion.bas — Alta de Solicitudes

### Qué buscar
Localiza el bloque donde se calcula la última fila con datos y se escribe el nuevo registro:
```vba
' Patrón típico a encontrar:
ultFila = Sheets("RESIDENCIA GIJÓN").Cells(Rows.Count, 1).End(xlUp).Row + 1
' ... asignar Nº Orden manual ...
Sheets("RESIDENCIA GIJÓN").Cells(ultFila, 1).Value = nuevoOrden
```

### Por qué sustituirlo
```vba
' === NUEVO: Inserción atómica en Access ===
Dim nuevoOrden As Long
nuevoOrden = modDatabase.InsertarOrdenAtomica( _
    residencia:="GIJON", _
    fechaPeticion:=Date, _
    dni:=txtDNI.Text, _
    nombre:=txtNombre.Text, _
    apellidos:=txtApellidos.Text, _
    tipoHuesped:=txtTipoHuesped.Text, _
    numHabInd:=CInt(txtHabInd.Text), _
    numHabDob:=CInt(txtHabDob.Text), _
    fechaEntrada:=CDate(txtFechaEntrada.Text), _
    fechaSalida:=CDate(txtFechaSalida.Text), _
    observaciones:=txtObservaciones.Text _
)

If nuevoOrden > 0 Then
    ' Doble escritura: también escribir en la hoja local para las macros visuales
    ultFila = Sheets("RESIDENCIA GIJÓN").Cells(Rows.Count, 1).End(xlUp).Row + 1
    Sheets("RESIDENCIA GIJÓN").Cells(ultFila, 1).Value = nuevoOrden
    ' ... escribir el resto de columnas en la hoja como antes ...
    
    MsgBox "Solicitud registrada. Nº ORDEN: " & nuevoOrden, vbInformation
Else
    MsgBox "Error al registrar la solicitud.", vbCritical
End If
```

### Notas
- Repetir para las 3 residencias (Gijón, Soto, Oviedo)
- Si el alta se hace desde un UserForm (`frmReservasUnidad`), el cambio es en el evento del botón Guardar

---

## 2. EvitarDuplicidadSolicitudesGyS.bas — Control de Duplicados

### Qué buscar
```vba
' Patrón típico: bucle For recorriendo celdas para buscar duplicados
For i = 2 To ultFila
    If Cells(i, colDNI).Value = dniBuscado And _
       Cells(i, colFechaEntrada).Value = fechaEntrada Then
        ' Duplicado encontrado
    End If
Next i
```

### Por qué sustituirlo
```vba
' === NUEVO: Consulta directa a Access ===
If modDatabase.ExisteDuplicado(dniBuscado, residencia, fechaEntrada, fechaSalida) Then
    MsgBox "Ya existe una solicitud con este DNI para las mismas fechas.", vbExclamation
    Exit Sub
End If
```

---

## 3. AsignarNumFactura.bas — Numeración de Facturas

### Qué buscar
La función `ObtenerNumeroFacturaInteligente()` que escanea la columna C buscando bloques:
```vba
' Patrón típico: bucle escaneando columna C
For i = 2 To ultFila
    If Cells(i, 3).Value <> "" Then
        ultimaFactura = CLng(Cells(i, 3).Value)
    End If
Next i
siguienteFactura = ultimaFactura + 1
' ... lógica de bloques con salto máximo de 3 ...
```

### Por qué sustituirlo
```vba
' === NUEVO: Contador atómico en Access ===
Dim nuevaFactura As Long
nuevaFactura = modDatabase.ObtenerSiguienteNumFactura("GIJON", Year(Now))

If nuevaFactura > 0 Then
    ' Asignar el número en Access
    modDatabase.ActualizarOrden numOrden, "NumFactura", CStr(nuevaFactura)
    
    ' Doble escritura: también en la hoja local
    Cells(filaReserva, 3).Value = nuevaFactura
    
    MsgBox "Factura nº " & nuevaFactura & " asignada.", vbInformation
End If
```

### Nota importante
Esto **elimina completamente** el riesgo de números de factura duplicados.
El sistema de bloques con salto máximo de 3 ya no es necesario, porque Access gestiona
la secuencia de forma atómica.

---

## 4. FacturacionMesGIJON/OVIEDO/SOTO.bas — Marcar como Pagado

### Qué buscar
El bloque que escribe el estado de pago en la columna AA/AB:
```vba
Cells(fila, colPagado).Value = formaPago  ' SMS/Bizum/Transferencia/etc.
```

### Por qué sustituirlo
```vba
' === NUEVO: Actualizar en Access Y en la hoja ===
Dim numOrden As Long
numOrden = CLng(Cells(fila, 1).Value)

' Actualizar en Access (fuente de verdad)
modDatabase.ActualizarOrden numOrden, "EstadoPago", formaPago

' Doble escritura: también en la hoja local
Cells(fila, colPagado).Value = formaPago
```

---

## 5. MarcarSiPagadosEnResidencia.bas — Sincronización con RESUMEN

### Qué buscar
Escritura en la hoja RESUMEN correspondiente:
```vba
Sheets("RESUMEN GIJÓN").Cells(filaResumen, colPagadoResumen).Value = "SI"
Sheets("RESUMEN GIJÓN").Cells(filaResumen, colFacturaResumen).Value = numFactura
```

### Por qué sustituirlo
```vba
' === NUEVO: Actualizar campo PagadoResidencia en Access ===
modDatabase.ActualizarOrden numOrden, "PagadoResidencia", "SI"

' Doble escritura en RESUMEN local (para impresión/visualización)
Sheets("RESUMEN GIJÓN").Cells(filaResumen, colPagadoResumen).Value = "SI"
Sheets("RESUMEN GIJÓN").Cells(filaResumen, colFacturaResumen).Value = numFactura
```

---

## 6. ModuloCalendarioGijon/Oviedo/Soto.bas — Renderizado del Calendario

### Qué buscar
La función principal que pinta el calendario (suele tener un bucle que lee filas):
```vba
Sub PintarCalendario_Gijon()
    ' ... código que recorre filas de la hoja RESIDENCIA GIJÓN ...
    For i = 2 To ultFila
        numOrden = Sheets("RESIDENCIA GIJÓN").Cells(i, 1).Value
        ' ... pintar en el calendario ...
    Next i
End Sub
```

### Qué añadir (AL INICIO de la función, antes del bucle)
```vba
Sub PintarCalendario_Gijon()
    ' === NUEVO: Sincronizar datos desde Access antes de pintar ===
    ' Esto garantiza que el usuario ve las reservas grabadas por otros operadores
    modDatabase.SincronizarHojaDesdeAccess "GIJON", Sheets("RESIDENCIA GIJÓN")
    
    ' ... el resto del código de pintado del calendario NO cambia ...
    For i = 2 To ultFila
        numOrden = Sheets("RESIDENCIA GIJÓN").Cells(i, 1).Value
        ' ... pintar en el calendario (código existente intacto) ...
    Next i
End Sub
```

### Nota
Esto es el cambio **más potente y sencillo** de toda la migración: añadir UNA LÍNEA
al inicio de cada función de calendario. El resto del código de pintado sigue
leyendo celdas normales — pero ahora esas celdas contienen datos frescos de Access.

---

## 7. BusquedaDNIResidencias.bas + BusquedaOrdenNombreFactura.bas

### Qué buscar
```vba
Set rng = Sheets("RESIDENCIA GIJÓN").Range("E:E").Find(What:=dniBuscado)
```

### Por qué sustituirlo
```vba
' === NUEVO: Búsqueda en Access (ve datos de TODOS los usuarios) ===
Dim rs As Object
Set rs = modDatabase.BuscarEnOrdenes("DNI", dniBuscado, "GIJON")

If Not rs Is Nothing Then
    If Not rs.EOF Then
        ' Mostrar resultados en el formulario de resultados
        ' (volcar a hoja temporal o rellenar un ListBox)
        Sheets("TempBusqueda").Range("A2").CopyFromRecordset rs
    Else
        MsgBox "No se encontraron resultados.", vbInformation
    End If
End If
```

---

## 8. ModuloLOG.bas — Registro de Actividad

### Qué buscar
```vba
ultFila = Sheets("LOG").Cells(Rows.Count, 1).End(xlUp).Row + 1
Sheets("LOG").Cells(ultFila, 1).Value = Now
Sheets("LOG").Cells(ultFila, 2).Value = usuario
Sheets("LOG").Cells(ultFila, 3).Value = accion
```

### Por qué sustituirlo
```vba
' === NUEVO: Log centralizado en Access ===
modDatabase.InsertarLog Environ("USERNAME"), accion, detalle
```

### Nota
Este es el módulo donde el cambio es **más limpio**: una sola línea sustituye 4.

---

## 9. ModListaNegra.bas — Verificación de DNI Vetado

### Qué buscar
```vba
Set rng = Sheets("LISTA NEGRA").Range("A:A").Find(What:=dni)
If Not rng Is Nothing Then
    MsgBox "Este DNI está en la lista negra.", vbCritical
End If
```

### Por qué sustituirlo
```vba
' === NUEVO: Consulta centralizada en Access ===
If modDatabase.ComprobarListaNegra(dni) Then
    MsgBox "Este DNI está en la lista negra.", vbCritical
    Exit Sub
End If
```

### Nota
Ventaja crítica: si el operador del PC #1 añade un DNI a la lista negra,
el operador del PC #2 lo verá de inmediato en su siguiente solicitud.
Con el sistema anterior basado en hojas, esto no ocurría.

---

## 10. ReevaluacionSolicitudes.bas — Reevaluación tras Renuncias

### Qué buscar
El módulo lee filas de candidatos en lista de espera y las modifica:
```vba
For i = 2 To ultFila
    If Cells(i, colResolucion).Value = "DENEGADA" Then
        ' ... comprobar si puede reasignarse ...
        Cells(i, colResolucion).Value = "REEVALUADA"
    End If
Next i
```

### Qué añadir
```vba
' Al INICIO del proceso de reevaluación, sincronizar datos frescos:
modDatabase.SincronizarHojaDesdeAccess "GIJON", Sheets("RESIDENCIA GIJÓN")

For i = 2 To ultFila
    If Cells(i, colResolucion).Value = "DENEGADA" Then
        ' ... comprobar si puede reasignarse (código existente) ...
        
        ' Si se reevalúa, actualizar en Access Y en la hoja local:
        Dim numOrden As Long
        numOrden = CLng(Cells(i, 1).Value)
        modDatabase.ActualizarOrden numOrden, "Resolucion", "REEVALUADA"
        Cells(i, colResolucion).Value = "REEVALUADA"
    End If
Next i
```

---

## 11. Eventos de Hoja (Clases .cls) — Propagación de Cambios

### Para las clases de hoja de RESIDENCIA GIJÓN, SOTO, OVIEDO

Si la hoja tiene un evento `Worksheet_Change` que reacciona a cambios manuales
en las celdas de datos, añadir la propagación a Access:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    ' ... código existente de validación/formateo (NO cambiar) ...
    
    ' === NUEVO: Propagar cambio manual a Access ===
    If Not Intersect(Target, Me.Range("A:AB")) Is Nothing Then
        If Target.Cells.Count = 1 Then  ' Solo cambios celda a celda
            Dim numOrden As Long
            numOrden = CLng(Me.Cells(Target.Row, 1).Value)
            If numOrden > 0 Then
                Dim campo As String
                campo = MapearColumnaACampoAccess(Target.Column)
                If Len(campo) > 0 Then
                    modDatabase.ActualizarOrden numOrden, campo, CStr(Target.Value)
                End If
            End If
        End If
    End If
End Sub

' Función auxiliar para mapear columna Excel → campo Access
Private Function MapearColumnaACampoAccess(ByVal col As Long) As String
    Select Case col
        Case 1: MapearColumnaACampoAccess = ""  ' NumOrden no se modifica
        Case 3: MapearColumnaACampoAccess = "NumFactura"
        Case 11: MapearColumnaACampoAccess = "Nombre"
        Case 12: MapearColumnaACampoAccess = "FechaEntrada"
        Case 13: MapearColumnaACampoAccess = "FechaSalida"
        Case 16: MapearColumnaACampoAccess = "Resolucion"
        Case 19: MapearColumnaACampoAccess = "HabitacionesAsignadas"  ' Col S (Gij/Soto)
        Case 20: MapearColumnaACampoAccess = "HabitacionesAsignadas"  ' Col T (Oviedo)
        Case 27: MapearColumnaACampoAccess = "EstadoPago"             ' Col AA (Gij/Soto)
        Case 28: MapearColumnaACampoAccess = "EstadoPago"             ' Col AB (Oviedo)
        Case Else: MapearColumnaACampoAccess = ""
    End Select
End Function
```

---

## 12. Workbook_Open — Sincronización al Arranque

### Añadir al evento `Workbook_Open` (en ThisWorkbook)

Después del login (`frmLogin`), sincronizar datos desde Access:

```vba
Private Sub Workbook_Open()
    ' ... código existente de login (frmLogin.Show) ...
    
    ' === NUEVO: Cargar datos frescos desde Access ===
    If modDatabase.ComprobarConexion() Then
        Application.StatusBar = "Sincronizando datos desde la base de datos central..."
        modDatabase.SincronizarHojaDesdeAccess "GIJON", Sheets("RESIDENCIA GIJÓN")
        modDatabase.SincronizarHojaDesdeAccess "SOTO", Sheets("RESIDENCIA SOTO")
        modDatabase.SincronizarHojaDesdeAccess "OVIEDO", Sheets("RESIDENCIA OVIEDO")
        Application.StatusBar = False
    Else
        MsgBox "AVISO: No se pudo conectar con la base de datos central en H:\." & vbCrLf & _
               "Estás trabajando con datos locales que pueden no estar actualizados.", _
               vbExclamation, "Sin conexión a Access"
    End If
End Sub
```

---

## Resumen de cambios por módulo

| # | Módulo | Tipo de cambio | Complejidad |
|---|--------|---------------|-------------|
| 1 | `FechasPeticion.bas` | Sustituir bloque de grabación | ⭐⭐ Media |
| 2 | `EvitarDuplicidadSolicitudesGyS.bas` | Sustituir bucle For por consulta SQL | ⭐ Baja |
| 3 | `AsignarNumFactura.bas` | Sustituir función de bloques por contador atómico | ⭐⭐ Media |
| 4 | `FacturacionMesGIJON.bas` | Añadir línea de UPDATE tras escritura en celda | ⭐ Baja |
| 5 | `FacturacionMesOVIEDO.bas` | Ídem | ⭐ Baja |
| 6 | `FacturacionMesSOTO.bas` | Ídem | ⭐ Baja |
| 7 | `MarcarSiPagadosEnResidencia.bas` | Añadir línea de UPDATE | ⭐ Baja |
| 8 | `ModuloCalendarioGijon.bas` | Añadir 1 línea de sincronización al inicio | ⭐ Baja |
| 9 | `ModuloCalendarioOviedo.bas` | Ídem | ⭐ Baja |
| 10 | `ModuloCalendarioSoto.bas` | Ídem | ⭐ Baja |
| 11 | `BusquedaDNIResidencias.bas` | Sustituir Find por consulta SQL | ⭐ Baja |
| 12 | `BusquedaOrdenNombreFactura.bas` | Ídem | ⭐ Baja |
| 13 | `ModuloLOG.bas` | Sustituir 4 líneas por 1 | ⭐ Baja |
| 14 | `ModListaNegra.bas` | Sustituir Find por consulta SQL | ⭐ Baja |
| 15 | `ReevaluacionSolicitudes.bas` | Añadir sincronización + UPDATE en bucle | ⭐⭐ Media |
