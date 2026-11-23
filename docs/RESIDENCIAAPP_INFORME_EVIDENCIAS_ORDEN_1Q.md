# RESIDENCIAAPP — INFORME DE EVIDENCIAS Y EXTRACCIÓN TÉCNICA (ORDEN 1Q)
## Verificación Integral de Dependencias, Código Fuente y Estructura Access para el Diseño Arquitectónico

- **Fecha de Emisión**: 2026-11-17
- **Naturaleza de la Intervención**: Estrictamente lectura forense, extracción íntegra de código y comprobación técnica de solo lectura.
- **Destinatario**: Dirección Arquitectónica de ResidenciaApp (ChatGPT).
- **Estado de las Restricciones Absolutas**: Cumplimiento del 100% (cero modificaciones de código, cero pseudocódigo, cero propuestas de diseño, cero commits/pushes, regla literal GIJÓN respetada, RGPD protegido sin datos personales).

---

## 1. FUENTES DE INFORMACIÓN Y TRAZABILIDAD CRIPTOGRÁFICA (SECCIÓN 3)

Se han inspeccionado y utilizado exclusivamente las fuentes autorizadas del proyecto, verificando sus dimensiones y sumas de verificación criptográfica SHA-256:

| Componente / Archivo | Ruta Absoluta | Tamaño (Bytes) | Hash SHA-256 | Estado de Verificación |
| :--- | :--- | :--- | :--- | :--- |
| **Libro Excel de Producción** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.621.963 | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` | Idéntico a Orden 1M |
| **Módulo Base Working Tree** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\modDatabase.bas` | 298.794 | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` | 100% Byte a byte idéntico al libro |
| **Extracción Forense 1M** | `C:\Users\...\modDatabase_FORENSE_ORDEN_1M.bas` | 298.794 | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` | Idéntico a Orden 1M |
| **Base de Datos Access (BE)** | `H:\ResidenciaBD\Residencia_BE.accdb` | 7.639.040 | `853eb4e736c02934668f0cc10d5a4d393f3cd66ceb40bf887b8993115f0ed430` | Activa de pruebas (2.127 órdenes) |

> [!NOTE]
> Todas las fuentes coinciden rigurosamente con los hashes acreditados en la Orden 1M. No existe ninguna discrepancia física entre el libro activo, los módulos del repositorio y la extracción forense de referencia.

---

## 2. EXTRACCIÓN Y ANÁLISIS DEL ARCHIVADO Y ELIMINACIÓN (SECCIÓN 4)

Se ha realizado la extracción completa, sin omisiones ni puntos suspensivos, de los cinco procedimientos centrales de supresión y archivado de órdenes.

### 2.1. Procedimiento `EliminarOrdenBD` (`modDatabase.bas`, líneas 4261-4420)

```vba
4261: Public Function EliminarOrdenBD( _
4262:     ByVal residencia As String, _
4263:     ByVal numOrden As Long _
4264: ) As Boolean
4265: 
4266:     Dim cn As Object
4267:     Dim claveRes As String
4268:     Dim enTransaccion As Boolean
4269: 
4270:     On Error GoTo ErrEliminar
4271:     EliminarOrdenBD = False
4272: 
4273:     If numOrden <= 0 Then Exit Function
4274:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
4275:         EliminarOrdenBD = False
4276:         Exit Function
4277:     End If
4278: 
4279:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
4280:     If claveRes <> "GIJON" And claveRes <> "SOTO" And claveRes <> "OVIEDO" Then Exit Function
4281: 
4282:     Call AsegurarEsquemaConcurrenciaBD
4283: 
4284:     Set cn = CreateObject("ADODB.Connection")
4285:     cn.ConnectionTimeout = 5
4286:     cn.Open GetConnectionString()
4287: 
4288:     cn.BeginTrans
4289:     enTransaccion = True
4290: 
4291:     ' === [CIERRE V5.2 - Punto 6: Adquisición atómica del Lease Lock dentro de la transacción antes de borrar] ===
4292:     Dim tokenActual As String, usrActualName As String
4293:     Dim sqlLockDel As String, recsLockedDel As Long
4294:     tokenActual = ObtenerTokenSesionInstancia()
4295:     usrActualName = Trim$(ObtenerUsuarioActual())
4296:     If Len(usrActualName) = 0 Then usrActualName = "admin"
4297: 
4298:     sqlLockDel = "UPDATE Ordenes SET [UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "', " & _
4299:                  "[UltimaModificacion] = " & FormatearFechaHoraSQL(Now) & _
4300:                  " WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(claveRes) & "' " & _
4301:                  "AND ([UsuarioModificacion] IS NULL OR [UsuarioModificacion] = '' " & _
4302:                  "OR [UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "' " & _
4303:                  "OR [UsuarioModificacion] = '" & EscaparSQL(usrActualName) & "' " & _
4304:                  "OR [UltimaModificacion] IS NULL " & _
4305:                  "OR DateDiff('s', [UltimaModificacion], Now()) >= " & LEASE_LOCK_SEGUNDOS & _
4306:                  " OR DateDiff('s', [UltimaModificacion], Now()) < 0)"
4307: 
4308:     recsLockedDel = 0
4309:     cn.Execute sqlLockDel, recsLockedDel
4310: 
4311:     Dim rsCheck As Object
4312:     Dim usrLock As String, fechaLock As Variant
4313:     Dim tieneDatosReales As Boolean
4314:     tieneDatosReales = False
4315: 
4316:     If recsLockedDel = 0 Then
4317:         ' No se pudo adquirir el bloqueo: comprobar si está bloqueado por otro operador o si ya no existe
4318:         Set rsCheck = cn.Execute("SELECT UsuarioCreacion, FechaCreacion, UsuarioModificacion, UltimaModificacion " & _
4319:                                  "FROM Ordenes WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(claveRes) & "'")
4320:         If Not rsCheck.EOF Then
4321:             usrLock = Trim$(CStr(Nz(rsCheck("UsuarioModificacion").Value, "")))
4322:             fechaLock = rsCheck("UltimaModificacion").Value
4323:             If Len(usrLock) = 0 Or IsNull(fechaLock) Or Not IsDate(fechaLock) Then
4324:                 usrLock = Trim$(CStr(Nz(rsCheck("UsuarioCreacion").Value, "")))
4325:                 fechaLock = rsCheck("FechaCreacion").Value
4326:             End If
4327:             rsCheck.Close
4328:             Set rsCheck = Nothing
4329:             cn.RollbackTrans
4330:             enTransaccion = False
4331:             cn.Close
4332:             Set cn = Nothing
4333: 
4334:             Dim segRestDel As Long
4335:             segRestDel = LEASE_LOCK_SEGUNDOS
4336:             If IsDate(fechaLock) Then
4337:                 segRestDel = LEASE_LOCK_SEGUNDOS - DateDiff("s", CDate(fechaLock), Now)
4338:                 If segRestDel < 1 Then segRestDel = 1
4339:             End If
4340: 
4341:             MsgBox "BORRADO BLOQUEADO POR SEGURIDAD MULTIUSUARIO:" & vbCrLf & vbCrLf & _
4342:                    "El Nº ORDEN " & numOrden & " (" & claveRes & ") está siendo editado en este instante por:" & vbCrLf & _
4343:                    "   Operador: " & FormatearPropietarioBloqueo(usrLock) & _
4344:                    " (bloqueo activo: " & segRestDel & " s restantes)." & vbCrLf & vbCrLf & _
4345:                    "No se ha eliminado de la base de datos central. Se refrescará la hoja.", _
4346:                    vbCritical, "Protección contra Borrado Cruzado - " & claveRes
4347:             Dim wsRefrescoDel As Worksheet
4348:             Set wsRefrescoDel = ObtenerHojaSegura(ObtenerNombreHoja(claveRes))
4349:             If Not wsRefrescoDel Is Nothing Then Call SincronizarHojaDesdeAccess(claveRes, wsRefrescoDel)
4350:             EliminarOrdenBD = False
4351:             Exit Function
4352:         Else
4353:             rsCheck.Close
4354:             Set rsCheck = Nothing
4355:             If enTransaccion Then cn.CommitTrans
4356:             enTransaccion = False
4357:             cn.Close
4358:             Set cn = Nothing
4359:             EliminarOrdenBD = True ' Idempotente: la orden ya no existe en la base de datos central
4360:             Exit Function
4361:         End If
4362:     End If
4363: 
4364:     ' Tenemos el bloqueo exclusivo dentro de la transacción -> comprobar si tiene datos reales de huésped
4365:     Set rsCheck = cn.Execute("SELECT DNI, Nombre FROM Ordenes WHERE NumOrden = " & numOrden & _
4366:                              " AND Residencia = '" & EscaparSQL(claveRes) & "'")
4367:     If Not rsCheck.EOF Then
4368:         If Len(Trim$(CStr(Nz(rsCheck("DNI").Value, "")))) > 0 Or _
4369:            Len(Trim$(CStr(Nz(rsCheck("Nombre").Value, "")))) > 0 Then
4370:             tieneDatosReales = True
4371:         End If
4372:     End If
4373:     rsCheck.Close
4374:     Set rsCheck = Nothing
4375: 
4376:     ' 1. Copiar a Ordenes_Historico SOLO si la orden tenía datos reales (evita archivar reservas en blanco canceladas)
4377:     If tieneDatosReales Then
4378:         cn.Execute "DELETE FROM Ordenes_Historico WHERE NumOrden = " & numOrden & _
4379:                    " AND Residencia = '" & EscaparSQL(claveRes) & "'"
4380: 
4381:         cn.Execute "INSERT INTO Ordenes_Historico SELECT * FROM Ordenes WHERE NumOrden = " & numOrden & _
4382:                    " AND Residencia = '" & EscaparSQL(claveRes) & "' " & _
4383:                    "AND [UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "'"
4384:     End If
4385: 
4386:     ' 2. Eliminar de la tabla activa Ordenes con predicado de autorización y comprobación de filas afectadas
4387:     Dim recsDeleted As Long
4388:     recsDeleted = 0
4389:     cn.Execute "DELETE FROM Ordenes WHERE NumOrden = " & numOrden & _
4390:                " AND Residencia = '" & EscaparSQL(claveRes) & "' " & _
4391:                "AND [UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "'", recsDeleted
4392: 
4393:     If recsDeleted <= 0 Then
4394:         cn.RollbackTrans
4395:         enTransaccion = False
4396:         cn.Close
4397:         Set cn = Nothing
4398:         EliminarOrdenBD = False
4399:         Exit Function
4400:     End If
4401: 
4402:     cn.CommitTrans
4403:     enTransaccion = False
4404: 
4405:     cn.Close
4406:     Set cn = Nothing
4407: 
4408:     EliminarOrdenBD = True
4409:     Exit Function
4410: 
4411: ErrEliminar:
4412:     On Error Resume Next
4413:     If Not cn Is Nothing Then
4414:         If enTransaccion Then cn.RollbackTrans
4415:         If cn.State = adStateOpen Then cn.Close
4416:     End If
4417:     Set cn = Nothing
4418:     EliminarOrdenBD = False
4419:     On Error GoTo 0
4420: End Function

```

### 2.2. Procedimiento `EliminarOrdenesLoteBD` (`modDatabase.bas`, líneas 4424-4552)

```vba
4424: Public Function EliminarOrdenesLoteBD( _
4425:     ByVal residencia As String, _
4426:     ByVal colOrdenes As Collection, _
4427:     Optional ByRef outTotalEliminadas As Long = 0 _
4428: ) As Boolean
4429:     Dim cn As Object, rsLock As Object
4430:     Dim claveRes As String, tokenActual As String, usrActualName As String
4431:     Dim enTransaccion As Boolean
4432:     Dim i As Long, contChunk As Long
4433:     Dim inClause As String
4434:     Dim recsAfectados As Long, totalAcum As Long
4435: 
4436:     On Error GoTo ErrLote
4437:     EliminarOrdenesLoteBD = False
4438:     outTotalEliminadas = 0
4439:     If colOrdenes Is Nothing Then Exit Function
4440:     If colOrdenes.Count = 0 Then
4441:         EliminarOrdenesLoteBD = True
4442:         Exit Function
4443:     End If
4444: 
4445:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
4446:     If claveRes <> "GIJON" And claveRes <> "SOTO" And claveRes <> "OVIEDO" Then Exit Function
4447: 
4448:     ' Soporte Contingencia SICO-N2 (Encolado optimizado en bloque matricial con 1 solo volcado a disco)
4449:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
4450:         If ModuloSeguridadLocal.EncolarBorradosLoteOffline(claveRes, colOrdenes) Then
4451:             outTotalEliminadas = colOrdenes.Count
4452:             EliminarOrdenesLoteBD = True
4453:         Else
4454:             EliminarOrdenesLoteBD = False
4455:         End If
4456:         Exit Function
4457:     End If
4458: 
4459:     Call AsegurarEsquemaConcurrenciaBD
4460:     tokenActual = ObtenerTokenSesionInstancia()
4461:     usrActualName = Trim$(ObtenerUsuarioActual())
4462:     If Len(usrActualName) = 0 Then usrActualName = "admin"
4463: 
4464:     Set cn = CreateObject("ADODB.Connection")
4465:     cn.ConnectionTimeout = 5
4466:     cn.Open GetConnectionString()
4467:     cn.BeginTrans
4468:     enTransaccion = True
4469: 
4470:     inClause = ""
4471:     contChunk = 0
4472:     totalAcum = 0
4473: 
4474:     For i = 1 To colOrdenes.Count
4475:         If IsNumeric(colOrdenes(i)) Then
4476:             If Len(inClause) > 0 Then inClause = inClause & ","
4477:             inClause = inClause & CLng(colOrdenes(i))
4478:             contChunk = contChunk + 1
4479:         End If
4480: 
4481:         If contChunk = 200 Or (i = colOrdenes.Count And contChunk > 0) Then
4482:             ' 0. Verificación V5.2 de Lease Lock activo por otro operador en este bloque
4483:             Set rsLock = cn.Execute("SELECT TOP 1 NumOrden, UsuarioModificacion, UltimaModificacion FROM Ordenes " & _
4484:                 "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden IN (" & inClause & ") " & _
4485:                 "AND [UsuarioModificacion] IS NOT NULL AND [UsuarioModificacion] <> '' " & _
4486:                 "AND [UsuarioModificacion] <> '" & EscaparSQL(tokenActual) & "' " & _
4487:                 "AND [UsuarioModificacion] <> '" & EscaparSQL(usrActualName) & "' " & _
4488:                 "AND [UltimaModificacion] IS NOT NULL " & _
4489:                 "AND DateDiff('s', [UltimaModificacion], Now()) >= 0 " & _
4490:                 "AND DateDiff('s', [UltimaModificacion], Now()) < " & LEASE_LOCK_SEGUNDOS)
4491:             If Not rsLock.EOF Then
4492:                 Dim ordBloq As Long, usrBloq As String
4493:                 ordBloq = CLng(Nz(rsLock("NumOrden").Value, 0))
4494:                 usrBloq = FormatearPropietarioBloqueo(CStr(Nz(rsLock("UsuarioModificacion").Value, "")))
4495:                 rsLock.Close
4496:                 Set rsLock = Nothing
4497:                 cn.RollbackTrans
4498:                 enTransaccion = False
4499:                 cn.Close
4500:                 Set cn = Nothing
4501:                 MsgBox "BORRADO DE LOTE CANCELADO POR SEGURIDAD MULTIUSUARIO (V5.2):" & vbCrLf & vbCrLf & _
4502:                        "El expediente Nº ORDEN " & ordBloq & " (" & claveRes & ") está siendo editado por:" & vbCrLf & _
4503:                        "   Operador: " & usrBloq & vbCrLf & vbCrLf & _
4504:                        "No se ha eliminado ningún registro para preservar la integridad transaccional.", _
4505:                        vbCritical, "Protección Lease Lock V5.2 - " & claveRes
4506:                 EliminarOrdenesLoteBD = False
4507:                 Exit Function
4508:             End If
4509:             rsLock.Close
4510:             Set rsLock = Nothing
4511: 
4512:             ' 1. Limpiar posibles duplicados previos en histórico para este lote
4513:             cn.Execute "DELETE FROM Ordenes_Historico WHERE Residencia = '" & EscaparSQL(claveRes) & _
4514:                        "' AND NumOrden IN (" & inClause & ")"
4515: 
4516:             ' 2. Archivar en histórico solo registros con datos reales (DNI o Nombre no vacíos)
4517:             cn.Execute "INSERT INTO Ordenes_Historico SELECT * FROM Ordenes " & _
4518:                        "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden IN (" & inClause & ") " & _
4519:                        "AND ((DNI IS NOT NULL AND Trim(DNI) <> '') OR (Nombre IS NOT NULL AND Trim(Nombre) <> ''))"
4520: 
4521:             ' 3. Eliminar del maestro activo Ordenes
4522:             recsAfectados = 0
4523:             cn.Execute "DELETE FROM Ordenes WHERE Residencia = '" & EscaparSQL(claveRes) & _
4524:                        "' AND NumOrden IN (" & inClause & ")", recsAfectados
4525:             totalAcum = totalAcum + recsAfectados
4526: 
4527:             inClause = ""
4528:             contChunk = 0
4529:         End If
4530:     Next i
4531: 
4532:     cn.CommitTrans
4533:     enTransaccion = False
4534:     cn.Close
4535:     Set cn = Nothing
4536: 
4537:     outTotalEliminadas = totalAcum
4538:     EliminarOrdenesLoteBD = True
4539:     Exit Function
4540: 
4541: ErrLote:
4542:     On Error Resume Next
4543:     If Not rsLock Is Nothing Then If rsLock.State = adStateOpen Then rsLock.Close
4544:     If Not cn Is Nothing Then
4545:         If enTransaccion Then cn.RollbackTrans
4546:         If cn.State = adStateOpen Then cn.Close
4547:     End If
4548:     Set rsLock = Nothing
4549:     Set cn = Nothing
4550:     EliminarOrdenesLoteBD = False
4551:     On Error GoTo 0
4552: End Function

```

### 2.3. Procedimiento `EliminarTodasOrdenesSedeBD` (`modDatabase.bas`, líneas 4556-4659)

```vba
4556: Public Function EliminarTodasOrdenesSedeBD( _
4557:     ByVal residencia As String, _
4558:     Optional ByRef outTotalEliminadas As Long = 0, _
4559:     Optional ByVal colOrdenesOffline As Collection = Nothing _
4560: ) As Boolean
4561:     Dim cn As Object, rsLock As Object
4562:     Dim claveRes As String, tokenActual As String, usrActualName As String
4563:     Dim enTransaccion As Boolean
4564:     Dim recsAfectados As Long
4565: 
4566:     On Error GoTo ErrTotalSede
4567:     EliminarTodasOrdenesSedeBD = False
4568:     outTotalEliminadas = 0
4569: 
4570:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
4571:     If claveRes <> "GIJON" And claveRes <> "SOTO" And claveRes <> "OVIEDO" Then Exit Function
4572: 
4573:     ' Soporte Contingencia SICO-N2 para Borrado Total
4574:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
4575:         If Not colOrdenesOffline Is Nothing Then
4576:             If ModuloSeguridadLocal.EncolarBorradosLoteOffline(claveRes, colOrdenesOffline) Then
4577:                 outTotalEliminadas = colOrdenesOffline.Count
4578:                 EliminarTodasOrdenesSedeBD = True
4579:             End If
4580:         End If
4581:         Exit Function
4582:     End If
4583: 
4584:     Call AsegurarEsquemaConcurrenciaBD
4585:     tokenActual = ObtenerTokenSesionInstancia()
4586:     usrActualName = Trim$(ObtenerUsuarioActual())
4587:     If Len(usrActualName) = 0 Then usrActualName = "admin"
4588: 
4589:     Set cn = CreateObject("ADODB.Connection")
4590:     cn.ConnectionTimeout = 5
4591:     cn.Open GetConnectionString()
4592:     cn.BeginTrans
4593:     enTransaccion = True
4594: 
4595:     ' 0. Verificar que ningún otro puesto tenga un Lease Lock activo (< 120 s) en esta sede
4596:     Set rsLock = cn.Execute("SELECT TOP 1 NumOrden, UsuarioModificacion FROM Ordenes " & _
4597:         "WHERE Residencia = '" & EscaparSQL(claveRes) & "' " & _
4598:         "AND [UsuarioModificacion] IS NOT NULL AND [UsuarioModificacion] <> '' " & _
4599:         "AND [UsuarioModificacion] <> '" & EscaparSQL(tokenActual) & "' " & _
4600:         "AND [UsuarioModificacion] <> '" & EscaparSQL(usrActualName) & "' " & _
4601:         "AND [UltimaModificacion] IS NOT NULL " & _
4602:         "AND DateDiff('s', [UltimaModificacion], Now()) >= 0 " & _
4603:         "AND DateDiff('s', [UltimaModificacion], Now()) < " & LEASE_LOCK_SEGUNDOS)
4604:     If Not rsLock.EOF Then
4605:         Dim ordBloqSede As Long, usrBloqSede As String
4606:         ordBloqSede = CLng(Nz(rsLock("NumOrden").Value, 0))
4607:         usrBloqSede = FormatearPropietarioBloqueo(CStr(Nz(rsLock("UsuarioModificacion").Value, "")))
4608:         rsLock.Close
4609:         Set rsLock = Nothing
4610:         cn.RollbackTrans
4611:         enTransaccion = False
4612:         cn.Close
4613:         Set cn = Nothing
4614:         MsgBox "BORRADO TOTAL CANCELADO POR SEGURIDAD MULTIUSUARIO (V5.2):" & vbCrLf & vbCrLf & _
4615:                "El expediente Nº ORDEN " & ordBloqSede & " (" & claveRes & ") está en edición activa por:" & vbCrLf & _
4616:                "   Operador: " & usrBloqSede & vbCrLf & vbCrLf & _
4617:                "Espere a que finalice la edición antes de vaciar la tabla de la sede.", _
4618:                vbCritical, "Protección Lease Lock V5.2 - " & claveRes
4619:         EliminarTodasOrdenesSedeBD = False
4620:         Exit Function
4621:     End If
4622:     rsLock.Close
4623:     Set rsLock = Nothing
4624: 
4625:     ' 1. Evitar colisión en histórico para las órdenes activas con datos reales de la sede
4626:     cn.Execute "DELETE FROM Ordenes_Historico WHERE Residencia = '" & EscaparSQL(claveRes) & _
4627:                "' AND NumOrden IN (SELECT NumOrden FROM Ordenes WHERE Residencia = '" & EscaparSQL(claveRes) & _
4628:                "' AND ((DNI IS NOT NULL AND Trim(DNI) <> '') OR (Nombre IS NOT NULL AND Trim(Nombre) <> '')))"
4629: 
4630:     ' 2. Copiar en bloque a Ordenes_Historico todas las órdenes con datos reales de la sede
4631:     cn.Execute "INSERT INTO Ordenes_Historico SELECT * FROM Ordenes " & _
4632:                "WHERE Residencia = '" & EscaparSQL(claveRes) & "' " & _
4633:                "AND ((DNI IS NOT NULL AND Trim(DNI) <> '') OR (Nombre IS NOT NULL AND Trim(Nombre) <> ''))"
4634: 
4635:     ' 3. Vaciar las órdenes activas de la sede en una sola sentencia
4636:     recsAfectados = 0
4637:     cn.Execute "DELETE FROM Ordenes WHERE Residencia = '" & EscaparSQL(claveRes) & "'", recsAfectados
4638: 
4639:     cn.CommitTrans
4640:     enTransaccion = False
4641:     cn.Close
4642:     Set cn = Nothing
4643: 
4644:     outTotalEliminadas = recsAfectados
4645:     EliminarTodasOrdenesSedeBD = True
4646:     Exit Function
4647: 
4648: ErrTotalSede:
4649:     On Error Resume Next
4650:     If Not rsLock Is Nothing Then If rsLock.State = adStateOpen Then rsLock.Close
4651:     If Not cn Is Nothing Then
4652:         If enTransaccion Then cn.RollbackTrans
4653:         If cn.State = adStateOpen Then cn.Close
4654:     End If
4655:     Set rsLock = Nothing
4656:     Set cn = Nothing
4657:     EliminarTodasOrdenesSedeBD = False
4658:     On Error GoTo 0
4659: End Function

```

### 2.4. Procedimiento `ArchivarOrdenesAntiguasBD` (`modDatabase.bas`, líneas 4784-4950)

```vba
4784: Public Function ArchivarOrdenesAntiguasBD( _
4785:     Optional ByVal residencia As String = "", _
4786:     Optional ByVal diasRetencion As Long = 90 _
4787: ) As Long
4788:     On Error GoTo ErrorHandler
4789:     ArchivarOrdenesAntiguasBD = 0
4790: 
4791:     If Len(Trim$(residencia)) = 0 Then
4792:         ArchivarOrdenesAntiguasBD = _
4793:             ArchivarOrdenesAntiguasBD("GIJON", diasRetencion) + _
4794:             ArchivarOrdenesAntiguasBD("SOTO", diasRetencion) + _
4795:             ArchivarOrdenesAntiguasBD("OVIEDO", diasRetencion)
4796:         Exit Function
4797:     End If
4798: 
4799:     Dim claveRes As String
4800:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
4801:     If claveRes <> "GIJON" And claveRes <> "SOTO" And claveRes <> "OVIEDO" Then Exit Function
4802: 
4803:     Dim colFilas As Collection, colOrdenes As Collection
4804:     Dim totalCandidatos As Long
4805:     totalCandidatos = ObtenerFilasYOrdenesArchivablesSede(claveRes, diasRetencion, colFilas, colOrdenes)
4806:     If totalCandidatos = 0 Then Exit Function
4807: 
4808:     ' 1. Archivar en Ordenes_Historico y eliminar de Ordenes en Access (o encolar en SICO-N2 si estamos Offline)
4809:     Dim eliminadasBD As Long
4810:     If colOrdenes.Count > 0 Then
4811:         If Not EliminarOrdenesLoteBD(claveRes, colOrdenes, eliminadasBD) Then
4812:             ArchivarOrdenesAntiguasBD = 0
4813:             Exit Function
4814:         End If
4815:     End If
4816: 
4817:         ' 2. Eliminar físicamente las filas antiguas en la hoja RESIDENCIA (compactación nativa Shift:=xlUp)
4818:     Dim ws As Worksheet
4819:     Set ws = ObtenerHojaSegura(ObtenerNombreHoja(claveRes))
4820:     If Not ws Is Nothing And colFilas.Count > 0 Then
4821:         Dim evOrig As Boolean, scrOrig As Boolean, calcOrig As XlCalculation
4822:         Dim estabaProt As Boolean, pwdHojas As String
4823:         Dim nomHojaCal As String, nomHojaSum As String
4824:         Dim wsCal As Worksheet, wsSum As Worksheet
4825: 
4826:         evOrig = Application.enableEvents
4827:         scrOrig = Application.screenUpdating
4828:         calcOrig = Application.calculation
4829: 
4830:         Application.enableEvents = False
4831:         Application.screenUpdating = False
4832:         Application.calculation = xlCalculationManual
4833: 
4834:         ' Suspender transitoriamente el cálculo en hojas satélite dependientes para evitar recálculos en cascada y parpadeo DWM
4835:         Select Case claveRes
4836:             Case "GIJON":  nomHojaCal = "Calendario GIJÓN": nomHojaSum = "RESUMEN GIJÓN"
4837:             Case "SOTO":   nomHojaCal = "Calendario SOTO":  nomHojaSum = "RESUMEN SOTO"
4838:             Case "OVIEDO": nomHojaCal = "Calendario OVIEDO": nomHojaSum = "RESUMEN OVIEDO"
4839:         End Select
4840: 
4841:         On Error Resume Next
4842:         If Len(nomHojaCal) > 0 Then Set wsCal = ThisWorkbook.Worksheets(nomHojaCal)
4843:         If Len(nomHojaSum) > 0 Then Set wsSum = ThisWorkbook.Worksheets(nomHojaSum)
4844:         If Not wsCal Is Nothing Then wsCal.EnableCalculation = False
4845:         If Not wsSum Is Nothing Then wsSum.EnableCalculation = False
4846:         On Error GoTo ErrorHandler
4847: 
4848:         estabaProt = ws.ProtectContents
4849:         pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()
4850:         If estabaProt Then
4851:             On Error Resume Next
4852:             ws.Unprotect password:=pwdHojas
4853:             On Error GoTo ErrorHandler
4854:         End If
4855:         If ws.FilterMode Then
4856:             On Error Resume Next
4857:             ws.ShowAllData
4858:             On Error GoTo ErrorHandler
4859:         End If
4860: 
4861:         Dim ultFilaA As Long
4862:         ultFilaA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
4863:         Dim totalFilasHoja As Long
4864:         totalFilasHoja = IIf(ultFilaA >= 2, ultFilaA - 1, 0)
4865: 
4866:         If colFilas.Count >= totalFilasHoja And ultFilaA >= 2 Then
4867:             ' Si se archivan todas las filas de la hoja, limpiar contenidos preservando fórmulas de plantilla
4868:             Select Case claveRes
4869:                 Case "GIJON"
4870:                     ws.Range(ws.Cells(2, 1), ws.Cells(ultFilaA, 20)).ClearContents
4871:                     ws.Range(ws.Cells(2, 22), ws.Cells(ultFilaA, 27)).ClearContents
4872:                     ws.Range(ws.Cells(2, 30), ws.Cells(ultFilaA, 30)).ClearContents
4873:                 Case "SOTO"
4874:                     ws.Range(ws.Cells(2, 1), ws.Cells(ultFilaA, 16)).ClearContents
4875:                     ws.Range(ws.Cells(2, 19), ws.Cells(ultFilaA, 20)).ClearContents
4876:                     ws.Range(ws.Cells(2, 22), ws.Cells(ultFilaA, 27)).ClearContents
4877:                     ws.Range(ws.Cells(2, 30), ws.Cells(ultFilaA, 30)).ClearContents
4878:                 Case "OVIEDO"
4879:                     ws.Range(ws.Cells(2, 1), ws.Cells(ultFilaA, 21)).ClearContents
4880:                     ws.Range(ws.Cells(2, 23), ws.Cells(ultFilaA, 31)).ClearContents
4881:             End Select
4882:             ws.Range("P2:P" & ultFilaA).ClearComments
4883:         Else
4884:             ' Agrupar en bloques de hasta 100 filas para evitar el límite de áreas en Application.Union
4885:             Dim rngBatch As Range, idxF As Long, contAreas As Long
4886:             For idxF = colFilas.Count To 1 Step -1
4887:                 If rngBatch Is Nothing Then
4888:                     Set rngBatch = ws.Rows(CLng(colFilas(idxF)))
4889:                     contAreas = 1
4890:                 Else
4891:                     Set rngBatch = Application.Union(rngBatch, ws.Rows(CLng(colFilas(idxF))))
4892:                     contAreas = contAreas + 1
4893:                 End If
4894:                 If contAreas >= 100 Or idxF = 1 Then
4895:                     rngBatch.Delete Shift:=xlUp
4896:                     Set rngBatch = Nothing
4897:                     contAreas = 0
4898:                 End If
4899:             Next idxF
4900:         End If
4901: 
4902:         If estabaProt Then
4903:             On Error Resume Next
4904:             ws.Protect password:=pwdHojas, UserInterfaceOnly:=True, _
4905:                 AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
4906:                 AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
4907:             On Error GoTo ErrorHandler
4908:         End If
4909: 
4910:         ' Reactivación escalonada del cálculo en hojas satélite dependientes con DoEvents
4911:         If Not wsCal Is Nothing Then
4912:             On Error Resume Next
4913:             wsCal.EnableCalculation = True
4914:             wsCal.Calculate
4915:             On Error GoTo ErrorHandler
4916:         End If
4917:         If Not wsSum Is Nothing Then
4918:             On Error Resume Next
4919:             wsSum.EnableCalculation = True
4920:             wsSum.Calculate
4921:             On Error GoTo ErrorHandler
4922:         End If
4923:         DoEvents
4924: 
4925:         Application.calculation = calcOrig
4926:         Application.enableEvents = evOrig
4927:         Application.screenUpdating = scrOrig
4928:     End If
4929: 
4930:     Dim usr As String
4931:     usr = m_NombreUsuario
4932:     If Len(usr) = 0 Then usr = Environ("USERNAME")
4933:     Call InsertarLog(usr, "ARCHIVADO_RGPD", _
4934:         "Archivadas " & totalCandidatos & " órdenes de " & claveRes & _
4935:         " (> " & diasRetencion & " días) a Ordenes_Historico")
4936: 
4937:     ArchivarOrdenesAntiguasBD = totalCandidatos
4938:     Exit Function
4939: 
4940: ErrorHandler:
4941:     On Error Resume Next
4942:     If Not wsCal Is Nothing Then wsCal.EnableCalculation = True
4943:     If Not wsSum Is Nothing Then wsSum.EnableCalculation = True
4944:     Application.calculation = xlCalculationAutomatic
4945:     Application.enableEvents = True
4946:     Application.screenUpdating = True
4947:     On Error GoTo 0
4948:     ArchivarOrdenesAntiguasBD = 0
4949:     MsgBox "Error al archivar órdenes:" & vbCrLf & Err.Description, vbCritical, "Error de Archivado"
4950: End Function

```

### 2.5. Procedimiento `PurgarHistoricoRGPD` (`modDatabase.bas`, líneas 4973-5003)

```vba
4973: Public Function PurgarHistoricoRGPD( _
4974:     Optional ByVal anosRetencion As Long = 3 _
4975: ) As Long
4976:     Dim totalPurgables As Long
4977:     Dim sqlDelete As String
4978:     Dim usr As String
4979:     
4980:     totalPurgables = ContarHistoricoPurgableRGPD(anosRetencion)
4981:     If totalPurgables = 0 Then
4982:         PurgarHistoricoRGPD = 0
4983:         Exit Function
4984:     End If
4985:     
4986:     sqlDelete = "DELETE FROM Ordenes_Historico WHERE FechaSalida <= DateAdd('yyyy', -" & anosRetencion & ", Date())"
4987:     
4988:     If ExecuteNonQuery(sqlDelete) Then
4989:         ' Purgar también de la réplica Snapshot si existiera
4990:         On Error Resume Next
4991:         ExecuteNonQuery "DELETE FROM BDAS_Historico_Snapshot WHERE FechaSalida <= DateAdd('yyyy', -" & anosRetencion & ", Date())"
4992:         On Error GoTo 0
4993:         
4994:         usr = m_NombreUsuario
4995:         If Len(usr) = 0 Then usr = Environ("USERNAME")
4996:         Call InsertarLog(usr, "PURGA_RGPD", _
4997:             "Borrado físico definitivo de " & totalPurgables & " óórdenes históricas (> " & anosRetencion & " años) según RD 933/2021")
4998:         
4999:         PurgarHistoricoRGPD = totalPurgables
5000:     Else
5001:         PurgarHistoricoRGPD = 0
5002:     End If
5003: End Function

```

### 2.6. Análisis Técnico Exhaustivo de los Procedimientos de Borrado y Archivado

| Aspecto Analizado | `EliminarOrdenBD` | `EliminarOrdenesLoteBD` | `EliminarTodasOrdenesSedeBD` | `ArchivarOrdenesAntiguasBD` | `PurgarHistoricoRGPD` |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Firma Completa** | `Public Function EliminarOrdenBD(ByVal residencia As String, ByVal numOrden As Long) As Boolean` | `Public Function EliminarOrdenesLoteBD(ByVal residencia As String, ByVal colOrdenes As Collection, Optional ByRef outTotalEliminadas As Long = 0) As Boolean` | `Public Function EliminarTodasOrdenesSedeBD(ByVal residencia As String, Optional ByRef outTotalEliminadas As Long = 0, Optional ByVal colOrdenesOffline As Collection = Nothing) As Boolean` | `Public Function ArchivarOrdenesAntiguasBD(Optional ByVal residencia As String = "", Optional ByVal diasRetencion As Long = 90) As Long` | `Public Function PurgarHistoricoRGPD(Optional ByVal anosRetencion As Long = 3) As Long` |
| **Módulo y Líneas** | `modDatabase.bas` (4261-4420) | `modDatabase.bas` (4424-4552) | `modDatabase.bas` (4556-4659) | `modDatabase.bas` (4784-4950) | `modDatabase.bas` (4973-5003) |
| **Llamadores Directos** | - `BorradoNumOrden.GestionarBorradoOrden` (l. 58)<br>- `ModuloSeguridadLocal.SincronizarColaOfflineHaciaAccess` (l. 993) | - `BorradoSolicitudes.BorrarSolicitudesPorOrden` (l. 389)<br>- `modDatabase.ArchivarOrdenesAntiguasBD` (l. 4811) | - `BorradoDatosAlojamiento.BorrarDatosAlojamientoOptimizado` (l. 150) | - `ModuloArchivadoDatosBDAS.ArchivarDatosAntiguosAlAbrir` (l. 109-111)<br>- `ModuloArchivadoDatosBDAS.ArchivarDatosManual` (l. 220) | - `BorradoFilasHojasBDAS.BorrarFilasBDAS` (l. 39) |
| **Operaciones SQL** | 1. `UPDATE Ordenes` (Lease Lock)<br>2. `SELECT DNI, Nombre`<br>3. `DELETE FROM Ordenes_Historico`<br>4. `INSERT INTO Ordenes_Historico SELECT *`<br>5. `DELETE FROM Ordenes` | 1. `SELECT TOP 1` (Lock check)<br>2. `DELETE FROM Ordenes_Historico WHERE NumOrden IN (...)`<br>3. `INSERT INTO Ordenes_Historico SELECT * WHERE NumOrden IN (...)`<br>4. `DELETE FROM Ordenes WHERE NumOrden IN (...)` | 1. `SELECT TOP 1` (Lock check)<br>2. `DELETE FROM Ordenes_Historico WHERE NumOrden IN (SELECT ...)`<br>3. `INSERT INTO Ordenes_Historico SELECT *`<br>4. `DELETE FROM Ordenes WHERE Residencia = ...` | Delega el borrado/archivado en base de datos en `EliminarOrdenesLoteBD`. En Excel ejecuta limpieza y compactación `Shift:=xlUp` en hoja residencia. | 1. `ContarHistoricoPurgableRGPD`<br>2. `DELETE FROM Ordenes_Historico WHERE FechaSalida <= DateAdd('yyyy', -N, Date())`<br>3. `DELETE FROM BDAS_Historico_Snapshot ...`<br>4. `InsertarLog` |
| **Límites Transaccionales** | Transacción explícita ADODB (`cn.BeginTrans`, `cn.CommitTrans`, `cn.RollbackTrans`). | Transacción explícita ADODB (`cn.BeginTrans`, `cn.CommitTrans`, `cn.RollbackTrans`). | Transacción explícita ADODB (`cn.BeginTrans`, `cn.CommitTrans`, `cn.RollbackTrans`). | No maneja transacción ADODB directa; delega en `EliminarOrdenesLoteBD`. Controla transitoriamente el entorno Excel (`ScreenUpdating`, `Calculation`, `EnableEvents`). | No maneja transacción explícita; ejecuta sentencias atómicas vía `ExecuteNonQuery`. |
| **Comprobación Filas Afectadas** | Sí: `recsDeleted` en `DELETE FROM Ordenes`. Si `recsDeleted <= 0`, hace `RollbackTrans` y retorna `False`. | Sí: `recsAfectados` acumulado por chunks de hasta 200 órdenes. | Sí: `recsAfectados` en `DELETE FROM Ordenes`. Retorna total borrado. | Retorna `totalCandidatos` detectados y procesados. | Retorna `totalPurgables` si `ExecuteNonQuery` tiene éxito. |
| **Condición de Archivado vs Eliminación** | **Solo evalúa DNI y Nombre**:<br>`tieneDatosReales = (Len(Trim(DNI)) > 0 Or Len(Trim(Nombre)) > 0)`.<br>Si es `True`: archiva (previo borrado histórico).<br>Si es `False`: elimina sin archivar. | **Solo evalúa DNI y Nombre**:<br>`((DNI IS NOT NULL AND Trim(DNI) <> '') OR (Nombre IS NOT NULL AND Trim(Nombre) <> ''))`.<br>Registros con DNI o Nombre se copian a histórico; el resto se purgan sin archivar. | **Solo evalúa DNI y Nombre**:<br>`((DNI IS NOT NULL AND Trim(DNI) <> '') OR (Nombre IS NOT NULL AND Trim(Nombre) <> ''))`.<br>Registros con DNI o Nombre se copian a histórico; el resto se purgan sin archivar. | Filtra órdenes cuya `FechaSalida` supere `diasRetencion` (default 90 días) respecto a la fecha actual mediante `ObtenerFilasYOrdenesArchivablesSede`. Luego aplica la condición de `EliminarOrdenesLoteBD`. | Filtra únicamente por antigüedad legal RGPD:<br>`FechaSalida <= DateAdd('yyyy', -anosRetencion, Date())`. No evalúa volumen de datos cumplimentados. |
| **Tratamiento de Errores** | `On Error GoTo ErrEliminar`. En error ejecuta `RollbackTrans`, cierra conexión y retorna `False`. Idempotente si no existe (`EliminarOrdenBD = True`). | `On Error GoTo ErrLote`. En error cierra `rsLock`, ejecuta `RollbackTrans`, cierra conexión y retorna `False`. Soporta encolado offline en SICO-N2. | `On Error GoTo ErrTotalSede`. En error ejecuta `RollbackTrans`, cierra conexión y retorna `False`. Soporta encolado offline en SICO-N2. | `On Error GoTo ErrorHandler`. En error restaura `ScreenUpdating`, `Calculation` y eventos de Excel, muestra `MsgBox` y retorna `0`. | `On Error Resume Next`. Retorna `0` ante fallo de ejecución SQL. |

### 2.7. Comprobación Específica del Cálculo Actual de `tieneDatosReales` y Representación de Columnas

#### A. Cálculo Real en Código de `tieneDatosReales`
En la implementación vigente, la determinación de si un expediente se archiva o se elimina definitivamente se realiza **exclusivamente a partir de dos campos de texto (`DNI` y `Nombre`)**:

```vba
' modDatabase.bas, líneas 4365-4372:
Set rsCheck = cn.Execute("SELECT DNI, Nombre FROM Ordenes WHERE NumOrden = " & numOrden & _
                         " AND Residencia = '" & EscaparSQL(claveRes) & "'")
If Not rsCheck.EOF Then
    If Len(Trim$(CStr(Nz(rsCheck("DNI").Value, "")))) > 0 Or _
       Len(Trim$(CStr(Nz(rsCheck("Nombre").Value, "")))) > 0 Then
        tieneDatosReales = True
    End If
End If
```

- **Hecho acreditado**: No existe ninguna evaluación de 15 columnas ni recuento de campos completados. Si un registro tiene `DNI` o `Nombre` cumplimentado, se archiva; si ambos están vacíos, se elimina definitivamente sin crear histórico.
- **Hecho acreditado en lotes y sede**: En `EliminarOrdenesLoteBD` (línea 4519) y `EliminarTodasOrdenesSedeBD` (línea 4628 y 4633), la condición SQL equivalente es idéntica: `((DNI IS NOT NULL AND Trim(DNI) <> '') OR (Nombre IS NOT NULL AND Trim(Nombre) <> ''))`.

#### B. Representación de Filas y Columnas en las Tres Residencias
A partir de la inspección de `SincronizarFilaHaciaAccess` (líneas 1751, 1898-1960) y `SincronizarHojaDesdeAccess` (líneas 3302-3520), se constata la correspondencia física exacta entre las celdas de Excel y los campos de Access:

| Columna Excel | Letra | Campo Access en GIJÓN | Campo Access en SOTO | Campo Access en OVIEDO | Tipo de Contenido |
| :---: | :---: | :--- | :--- | :--- | :--- |
| 1 | **A** | `NumOrden` | `NumOrden` | `NumOrden` | Dato (Identificador)
| 2 | **B** | `FechaPeticion` | `FechaPeticion` | `FechaPeticion` | Dato (Fecha)
| 3 | **C** | `NumFactura` | `NumFactura` | `NumFactura` | Dato (Número)
| 4 | **D** | `Finalidad` | `Finalidad` | `Finalidad` | Dato (Texto)
| 5 | **E** | `Empleo` | `Empleo` | `Empleo` | Dato (Texto)
| 6 | **F** | `Situacion` | `Situacion` | `Situacion` | Dato (Texto)
| 7 | **G** | `Evaluacion` | `Evaluacion` | `Evaluacion` | Dato (Texto)
| 8 | **H** | `Comision` | `Turno` | `Comision` | Dato (Texto específico sede)
| 9 | **I** | `DNI` | `DNI` | `DNI` | Dato (Texto)
| 10 | **J** | `Rango` | `Rango` | `Rango` | Dato (Texto)
| 11 | **K** | `Nombre` | `Nombre` | `Nombre` | Dato (Texto)
| 12 | **L** | `FechaEntrada` | `FechaEntrada` | `FechaEntrada` | Dato (Fecha)
| 13 | **M** | `FechaSalida` | `FechaSalida` | `FechaSalida` | Dato (Fecha)
| 14 | **N** | `DiasUso` | `DiasUso` | `DiasUso` | Dato (Número)
| 15 | **O** | `PAX` | `PAX` | `PAX` | Dato (Número)
| 16 | **P** | `Resolucion` | `Resolucion` | `Resolucion` | Dato (Texto)
| 17 | **Q** | `NumHabIndividuales` | *(Fórmula: PRECIO APTO)* | `NumHabIndividuales` | Dato en GIJÓN/OVIEDO; Fórmula en SOTO
| 18 | **R** | `NumHabDobles` | *(Fórmula: SUPLE OCUPAN)* | `NumHabDobles` | Dato en GIJÓN/OVIEDO; Fórmula en SOTO
| 19 | **S** | `HabitacionesAsignadas` | `HabitacionesAsignadas` | `CamaSuple` | Dato (Texto / Cama supletoria)
| 20 | **T** | `DtoFamNum` | `DtoFamNum` | `HabitacionesAsignadas` | Dato (Texto)
| 21 | **U** | *(Fórmula: IMPORTE)* | *(Fórmula: IMPORTE)* | `DtoFamNum` | Fórmula en GIJÓN/SOTO; Dato en OVIEDO
| 22 | **V** | `Telefono` | `Telefono` | *(Fórmula: IMPORTE)* | Dato en GIJÓN/SOTO; Fórmula en OVIEDO
| 23 | **W** | `Direccion` | `Direccion` | `Telefono` | Dato (Texto)
| 24 | **X** | `CodigoPostal` | `CodigoPostal` | `Direccion` | Dato (Texto)
| 25 | **Y** | `Poblacion` | `Poblacion` | `CodigoPostal` | Dato (Texto)
| 26 | **Z** | `Provincia` | `Provincia` | `Poblacion` | Dato (Texto)
| 27 | **AA** | `EstadoPago` | `EstadoPago` | `Provincia` | Dato (Texto)
| 28 | **AB** | *(Fórmula: NOMBRE PDF)* | *(Fórmula: NOMBRE PDF)* | `EstadoPago` | Fórmula en GIJÓN/SOTO; Dato en OVIEDO
| 29 | **AC** | *(Fórmula: TEXTO EMAIL)* | *(Fórmula: TEXTO EMAIL)* | *(No usado)* | Fórmula en GIJÓN/SOTO
| 32 | **AF** | *(No usado)* | *(No usado)* | *(Fórmula: NOMBRE PDF)* | Fórmula en OVIEDO
| 33 | **AG** | *(No usado)* | *(No usado)* | *(Fórmula: TEXTO EMAIL)* | Fórmula en OVIEDO

---

## 3. EXTRACCIÓN Y ANÁLISIS DE LA NUMERACIÓN ONLINE (SECCIÓN 5)

### 3.1. Procedimiento `IntentarReservarOrdenBD` (`modDatabase.bas`, líneas 1281-1401)

```vba
1281: Public Function IntentarReservarOrdenBD( _
1282:     ByVal residencia As String, _
1283:     ByVal numOrdenSolicitado As Long, _
1284:     ByRef outUsuarioPropietario As String, _
1285:     ByRef outNombreHuesped As String, _
1286:     ByRef outSiguienteLibreBD As Long, _
1287:     Optional ByRef outSinRed As Boolean = False _
1288: ) As Boolean
1289: 
1290:     Dim cn As Object, rs As Object
1291:     Dim claveRes As String, usrActual As String, tokenActual As String
1292:     Dim enTrans As Boolean
1293: 
1294:     On Error GoTo ErrReserva
1295:     IntentarReservarOrdenBD = False
1296:     outUsuarioPropietario = ""
1297:     outNombreHuesped = ""
1298:     outSiguienteLibreBD = 0
1299:     outSinRed = False
1300: 
1301:     If numOrdenSolicitado <= 0 Then Exit Function
1302:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
1303:         outSinRed = True
1304:         IntentarReservarOrdenBD = False
1305:         Exit Function
1306:     End If
1307: 
1308:     Call AsegurarEsquemaConcurrenciaBD
1309: 
1310:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
1311:     usrActual = Trim$(ObtenerUsuarioActual())
1312:     If Len(usrActual) = 0 Then usrActual = "admin"
1313:     tokenActual = ObtenerTokenSesionInstancia()
1314: 
1315:     Set cn = CreateObject("ADODB.Connection")
1316:     cn.ConnectionTimeout = 5
1317:     cn.Open GetConnectionString()
1318: 
1319:     cn.BeginTrans
1320:     enTrans = True
1321: 
1322:     Set rs = cn.Execute("SELECT UsuarioCreacion, UsuarioModificacion, Nombre FROM Ordenes " & _
1323:                         "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden = " & numOrdenSolicitado)
1324: 
1325:     If Not rs.EOF Then
1326:         ' Ya existe en Access -> Leer propietario y calcular siguiente libre
1327:         outUsuarioPropietario = Trim$(CStr(Nz(rs("UsuarioModificacion").Value, "")))
1328:         If Len(outUsuarioPropietario) = 0 Then
1329:             outUsuarioPropietario = Trim$(CStr(Nz(rs("UsuarioCreacion").Value, "")))
1330:         End If
1331:         outUsuarioPropietario = FormatearPropietarioBloqueo(outUsuarioPropietario)
1332:         outNombreHuesped = Trim$(CStr(Nz(rs("Nombre").Value, "")))
1333:         rs.Close
1334:         Set rs = Nothing
1335: 
1336:         Set rs = cn.Execute("SELECT MAX(NumOrden) FROM Ordenes " & _
1337:                             "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden < 900001")
1338:         If Not rs.EOF And Not IsNull(rs(0).Value) Then
1339:             outSiguienteLibreBD = CLng(rs(0).Value) + 1
1340:         Else
1341:             outSiguienteLibreBD = numOrdenSolicitado + 1
1342:         End If
1343:         rs.Close
1344:         Set rs = Nothing
1345: 
1346:         cn.RollbackTrans
1347:         enTrans = False
1348:         cn.Close
1349:         Set cn = Nothing
1350:         IntentarReservarOrdenBD = False
1351:         Exit Function
1352:     End If
1353:     rs.Close
1354:     Set rs = Nothing
1355: 
1356:     ' Está libre -> Insertar y adjudicar atómicamente dentro de la transacción
1357:     Dim sqlIns As String
1358:     Dim recsIns As Long
1359:     sqlIns = "INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, UsuarioCreacion, " & _
1360:              "UsuarioModificacion, UltimaModificacion, Solapamiento, ConsentimientoRGPD) VALUES (" & _
1361:              numOrdenSolicitado & ", '" & EscaparSQL(claveRes) & "', " & _
1362:              FormatearFechaHoraSQL(Now) & ", '" & EscaparSQL(usrActual) & "', '" & _
1363:              EscaparSQL(tokenActual) & "', " & FormatearFechaHoraSQL(Now) & ", False, False)"
1364: 
1365:     cn.Execute sqlIns, recsIns
1366:     If recsIns <> 1 Then
1367:         cn.RollbackTrans
1368:         enTrans = False
1369:         cn.Close
1370:         Set cn = Nothing
1371:         IntentarReservarOrdenBD = False
1372:         Exit Function
1373:     End If
1374: 
1375:     cn.CommitTrans
1376:     enTrans = False
1377:     cn.Close
1378:     Set cn = Nothing
1379: 
1380:     IntentarReservarOrdenBD = True
1381:     Exit Function
1382: 
1383: ErrReserva:
1384:     On Error Resume Next
1385:     If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
1386:     If Not cn Is Nothing Then
1387:         If enTrans Then cn.RollbackTrans
1388:         If cn.State = adStateOpen Then cn.Close
1389:     End If
1390:     Set rs = Nothing
1391:     Set cn = Nothing
1392:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
1393:         outSinRed = True
1394:     Else
1395:         Dim dniDummy As String
1396:         Call ExisteOrdenEnBD(residencia, numOrdenSolicitado, outUsuarioPropietario, outNombreHuesped, dniDummy, outSiguienteLibreBD)
1397:         outUsuarioPropietario = FormatearPropietarioBloqueo(outUsuarioPropietario)
1398:     End If
1399:     IntentarReservarOrdenBD = False
1400:     On Error GoTo 0
1401: End Function

```

### 3.2. Procedimiento `ReservarSiguienteOrdenLibreBD` (`modDatabase.bas`, líneas 1403-1472)

```vba
1403: Public Function ReservarSiguienteOrdenLibreBD(ByVal residencia As String) As Long
1404:     Dim cn As Object, rs As Object
1405:     Dim claveRes As String, usrActual As String, tokenActual As String
1406:     Dim candidato As Long, maxIter As Integer
1407:     Dim insertado As Boolean
1408: 
1409:     On Error GoTo ErrReservaSig
1410:     ReservarSiguienteOrdenLibreBD = 0
1411: 
1412:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then Exit Function
1413:     Call AsegurarEsquemaConcurrenciaBD
1414: 
1415:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
1416:     usrActual = Trim$(ObtenerUsuarioActual())
1417:     If Len(usrActual) = 0 Then usrActual = "admin"
1418:     tokenActual = ObtenerTokenSesionInstancia()
1419: 
1420:     Set cn = CreateObject("ADODB.Connection")
1421:     cn.ConnectionTimeout = 5
1422:     cn.Open GetConnectionString()
1423: 
1424:     maxIter = 0
1425:     insertado = False
1426: 
1427:     Do While Not insertado And maxIter < 20
1428:         maxIter = maxIter + 1
1429:         Set rs = cn.Execute("SELECT MAX(NumOrden) FROM Ordenes " & _
1430:                             "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden < 900001")
1431:         If Not rs.EOF And Not IsNull(rs(0).Value) Then
1432:             candidato = CLng(rs(0).Value) + 1
1433:         Else
1434:             candidato = 1
1435:         End If
1436:         rs.Close
1437:         Set rs = Nothing
1438: 
1439:         On Error Resume Next
1440:         cn.BeginTrans
1441:         Dim sqlIns As String
1442:         sqlIns = "INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, UsuarioCreacion, " & _
1443:                  "UsuarioModificacion, UltimaModificacion, Solapamiento, ConsentimientoRGPD) VALUES (" & _
1444:                  candidato & ", '" & EscaparSQL(claveRes) & "', " & _
1445:                  FormatearFechaHoraSQL(Now) & ", '" & EscaparSQL(usrActual) & "', '" & _
1446:                  EscaparSQL(tokenActual) & "', " & FormatearFechaHoraSQL(Now) & ", False, False)"
1447:         cn.Execute sqlIns
1448:         If Err.Number = 0 Then
1449:             cn.CommitTrans
1450:             insertado = True
1451:         Else
1452:             cn.RollbackTrans
1453:             Err.Clear
1454:             candidato = 0
1455:         End If
1456:         On Error GoTo ErrReservaSig
1457:     Loop
1458: 
1459:     cn.Close
1460:     Set cn = Nothing
1461:     ReservarSiguienteOrdenLibreBD = candidato
1462:     Exit Function
1463: 
1464: ErrReservaSig:
1465:     On Error Resume Next
1466:     If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
1467:     If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
1468:     Set rs = Nothing
1469:     Set cn = Nothing
1470:     ReservarSiguienteOrdenLibreBD = 0
1471:     On Error GoTo 0
1472: End Function

```

### 3.3. Procedimiento `RenumerarOrdenExistenteBD` (`modDatabase.bas`, líneas 1475-1525)

```vba
1475: Public Function RenumerarOrdenExistenteBD( _
1476:     ByVal claveResidencia As String, _
1477:     ByVal numOrdenPrevio As Long, _
1478:     ByVal numOrdenNuevo As Long _
1479: ) As Boolean
1480:     Dim cn As Object
1481:     Dim sqlUpd As String
1482:     Dim recs As Long
1483:     Dim tokenActual As String
1484: 
1485:     On Error GoTo ErrRenumerar
1486:     RenumerarOrdenExistenteBD = False
1487:     If numOrdenPrevio <= 0 Or numOrdenNuevo <= 0 Or numOrdenPrevio = numOrdenNuevo Then Exit Function
1488:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then Exit Function
1489: 
1490:     tokenActual = ObtenerTokenSesionInstancia()
1491: 
1492:     Set cn = CreateObject("ADODB.Connection")
1493:     cn.ConnectionTimeout = 5
1494:     cn.Open GetConnectionString()
1495: 
1496:     cn.BeginTrans
1497:     sqlUpd = "UPDATE Ordenes SET [NumOrden] = " & numOrdenNuevo & ", " & _
1498:              "[UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "', " & _
1499:              "[UltimaModificacion] = " & FormatearFechaHoraSQL(Now) & _
1500:              " WHERE Residencia = '" & EscaparSQL(claveResidencia) & "' AND NumOrden = " & numOrdenPrevio
1501:     recs = 0
1502:     cn.Execute sqlUpd, recs
1503: 
1504:     If recs = 1 Then
1505:         cn.CommitTrans
1506:         RenumerarOrdenExistenteBD = True
1507:     Else
1508:         cn.RollbackTrans
1509:         RenumerarOrdenExistenteBD = False
1510:     End If
1511: 
1512:     cn.Close
1513:     Set cn = Nothing
1514:     Exit Function
1515: 
1516: ErrRenumerar:
1517:     On Error Resume Next
1518:     If Not cn Is Nothing Then
1519:         cn.RollbackTrans
1520:         If cn.State = adStateOpen Then cn.Close
1521:     End If
1522:     Set cn = Nothing
1523:     RenumerarOrdenExistenteBD = False
1524:     On Error GoTo 0
1525: End Function

```

### 3.4. Procedimiento `GestionarEntradaNumOrdenHoja` (`modDatabase.bas`, líneas 1527-1696)

```vba
1527: Public Function GestionarEntradaNumOrdenHoja( _
1528:     ByVal ws As Worksheet, _
1529:     ByVal Target As Range, _
1530:     ByVal targetRow As Long, _
1531:     ByRef ioValorAnteriorHoja As Variant, _
1532:     ByVal claveResidencia As String, _
1533:     ByVal nombreHojaResumen As String, _
1534:     ByVal colFactSum As String, _
1535:     ByVal colPagoSum As String, _
1536:     ByVal colPagoRes As String _
1537: ) As Boolean
1538: 
1539:     GestionarEntradaNumOrdenHoja = False
1540:     If ws Is Nothing Or Target Is Nothing Or targetRow < 2 Then Exit Function
1541: 
1542:     Dim celdaVal As Variant
1543:     celdaVal = Target.Value
1544:     If IsEmpty(celdaVal) Then Exit Function
1545:     If Not IsNumeric(celdaVal) Then Exit Function
1546: 
1547:     Dim numSolicitado As Long
1548:     numSolicitado = CLng(celdaVal)
1549:     If numSolicitado <= 0 Then Exit Function
1550: 
1551:     Dim teniamosNumeroPrevio As Boolean
1552:     Dim numPrevio As Long
1553:     numPrevio = 0
1554:     If Not IsEmpty(ioValorAnteriorHoja) Then
1555:         If IsNumeric(ioValorAnteriorHoja) Then
1556:             numPrevio = CLng(ioValorAnteriorHoja)
1557:             If numPrevio > 0 And numPrevio <> numSolicitado Then
1558:                 teniamosNumeroPrevio = True
1559:             End If
1560:         End If
1561:     End If
1562: 
1563:     Dim usrProp As String, nomHuesp As String, sigLibre As Long
1564: 
1565:     ' CASO A: Modificación / Renumeración de una orden que ya tenía número previo en esta fila
1566:     If teniamosNumeroPrevio Then
1567:         Dim ocupadoPorOtro As Boolean
1568:         Dim dniDummy As String
1569:         ocupadoPorOtro = ExisteOrdenEnBD(claveResidencia, numSolicitado, usrProp, nomHuesp, dniDummy, sigLibre)
1570: 
1571:         If ocupadoPorOtro Then
1572:             MsgBox "COLISIÓN AL RENUMERAR ORDEN:" & vbCrLf & vbCrLf & _
1573:                    "El Nº ORDEN " & numSolicitado & " ya está asignado en la base de datos (" & claveResidencia & "):" & vbCrLf & _
1574:                    "   Adjudicado a: " & IIf(Len(usrProp) > 0, usrProp, "(Otro operador)") & vbCrLf & _
1575:                    "   Huésped: " & IIf(Len(nomHuesp) > 0, nomHuesp, "(En grabación)") & vbCrLf & vbCrLf & _
1576:                    "Se restaurará el Nº ORDEN anterior (" & numPrevio & ") en esta fila.", _
1577:                    vbExclamation, "Colisión de Renumeración - " & claveResidencia
1578:             ws.Cells(targetRow, "A").Value = numPrevio
1579:             GestionarEntradaNumOrdenHoja = False
1580:             Exit Function
1581:         Else
1582:             If RenumerarOrdenExistenteBD(claveResidencia, numPrevio, numSolicitado) Then
1583:                 ioValorAnteriorHoja = numSolicitado
1584:                 Call MarcarSiPagadosEnResidencia.SincronizarOrdenYCampos( _
1585:                     ws.Name, "C", colPagoRes, nombreHojaResumen, colFactSum, colPagoSum, numSolicitado)
1586:                 Call SincronizarFilaHaciaAccess(ws, targetRow)
1587:                 GestionarEntradaNumOrdenHoja = True
1588:                 Exit Function
1589:             Else
1590:                 MsgBox "No se pudo actualizar el Nº ORDEN en la base de datos central. Se restaura el anterior.", _
1591:                        vbExclamation, "Error de Persistencia - " & claveResidencia
1592:                 ws.Cells(targetRow, "A").Value = numPrevio
1593:                 GestionarEntradaNumOrdenHoja = False
1594:                 Exit Function
1595:             End If
1596:         End If
1597:     End If
1598: 
1599:     ' CASO B: Alta nueva en una celda A que estaba vacía -> Reserva atómica en Access (o provisional 900xxx si Offline)
1600:     Dim reservadoOK As Boolean, sinRedReserva As Boolean
1601:     reservadoOK = IntentarReservarOrdenBD(claveResidencia, numSolicitado, usrProp, nomHuesp, sigLibre, sinRedReserva)
1602: 
1603:     If Not reservadoOK Then
1604:         ' Subcaso B.1: Sin conexión con Access -> Impedir adjudicar un Nº ORDEN oficial sin confirmación del servidor
1605:         If sinRedReserva Then
1606:             Dim numProv As Long
1607:             If ModuloSeguridadLocal.EsNumOrdenProvisional(numSolicitado) Then
1608:                 numProv = numSolicitado
1609:             Else
1610:                 numProv = ModuloSeguridadLocal.ObtenerSiguienteNumOrdenProvisional(ws)
1611:             End If
1612:             ws.Cells(targetRow, "A").Value = numProv
1613:             ioValorAnteriorHoja = numProv
1614:             Call ModuloSeguridadLocal.EncolarOperacionOffline( _
1615:                 claveResidencia, numProv, "PRERESERVA", "NumOrden", "", CStr(numProv))
1616:             MsgBox "MODO CONTINGENCIA SICO-N2 (SIN CONEXIÓN CON ACCESS):" & vbCrLf & vbCrLf & _
1617:                    "No es posible reservar en firme el Nº ORDEN oficial " & numSolicitado & " sin confirmación de la base de datos central." & vbCrLf & _
1618:                    "Se ha asignado el Nº ORDEN provisional " & numProv & ", que será convertido automáticamente en el siguiente Nº oficial libre al restablecerse la red.", _
1619:                    vbInformation, "Pre-Reserva Offline SICO-N2 - " & claveResidencia
1620:             ws.Cells(targetRow, 2).Select
1621:             GestionarEntradaNumOrdenHoja = False ' Encolado offline: no duplicar en LOG hasta conciliación FIFO
1622:             Exit Function
1623:         End If
1624: 
1625:         ' Subcaso B.2: COLISIÓN DETECTADA EN ACCESS:
1626:         ' 1. Limpiar inmediatamente la celda y resetear ValorAnterior
1627:         ws.Cells(targetRow, "A").ClearContents
1628:         ioValorAnteriorHoja = Empty
1629: 
1630:         ' 2. Sincronizar la hoja desde Access para que el operador vea los registros reales del servidor
1631:         Call SincronizarHojaDesdeAccess(claveResidencia, ws)
1632: 
1633:         ' 3. Ofrecer interactivamente al operador el siguiente Nº ORDEN libre disponible en la BD
1634:         Dim msgOferta As String
1635:         Dim respOferta As VbMsgBoxResult
1636: 
1637:         msgOferta = "ATENCIÓN - Nº ORDEN " & numSolicitado & " YA ASIGNADO EN LA BASE DE DATOS:" & vbCrLf & vbCrLf & _
1638:                     "   Adjudicado a: " & IIf(Len(usrProp) > 0, usrProp, "(Otro operador en edición)") & vbCrLf & _
1639:                     "   Huésped: " & IIf(Len(nomHuesp) > 0, nomHuesp, "(En grabación)") & vbCrLf & vbCrLf & _
1640:                     "No se puede utilizar el mismo Nº ORDEN en esta fila." & vbCrLf & _
1641:                     "El siguiente Nº ORDEN disponible en la base de datos es el:  " & sigLibre & vbCrLf & vbCrLf & _
1642:                     "¿Desea adjudicar y utilizar automáticamente el Nº ORDEN " & sigLibre & " ahora?"
1643: 
1644:         respOferta = MsgBox(msgOferta, vbYesNo + vbQuestion + vbDefaultButton1, _
1645:                             "Colisión de Nº ORDEN Evitada - " & claveResidencia)
1646: 
1647:         If respOferta = vbYes Then
1648:             Dim numAdjudicado As Long
1649:             numAdjudicado = ReservarSiguienteOrdenLibreBD(claveResidencia)
1650:             If numAdjudicado > 0 Then
1651:                 Dim filaLibre As Long
1652:                 Dim ultFilaActual As Long
1653:                 ultFilaActual = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
1654:                 If ultFilaActual < 2 And Trim$(CStr(ws.Cells(2, "A").Value)) = "" Then
1655:                     filaLibre = 2
1656:                 Else
1657:                     filaLibre = ultFilaActual + 1
1658:                 End If
1659: 
1660:                 If targetRow >= filaLibre And Trim$(CStr(ws.Cells(targetRow, "A").Value)) = "" Then
1661:                     filaLibre = targetRow
1662:                 End If
1663: 
1664:                 ws.Cells(filaLibre, "A").Value = numAdjudicado
1665:                 ioValorAnteriorHoja = numAdjudicado
1666: 
1667:                 Call MarcarSiPagadosEnResidencia.SincronizarOrdenYCampos( _
1668:                     ws.Name, "C", colPagoRes, nombreHojaResumen, colFactSum, colPagoSum, numAdjudicado)
1669:                 Call SincronizarFilaHaciaAccess(ws, filaLibre, True)
1670: 
1671:                 ws.Cells(filaLibre, 2).Select
1672:                 If numAdjudicado <> sigLibre Then
1673:                     MsgBox "Se le ha adjudicado y reservado en firme el Nº ORDEN " & numAdjudicado & _
1674:                            " en la fila " & filaLibre & ".", vbInformation, "Nº ORDEN Adjudicado"
1675:                 End If
1676:                 GestionarEntradaNumOrdenHoja = True
1677:             Else
1678:                 MsgBox "No se pudo confirmar la reserva del Nº ORDEN en la base de datos Access.", _
1679:                        vbExclamation, "Reserva No Confirmada - " & claveResidencia
1680:             End If
1681:         Else
1682:             Dim filaVaciaFinal As Long
1683:             filaVaciaFinal = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
1684:             If filaVaciaFinal < 2 Then filaVaciaFinal = 2
1685:             ioValorAnteriorHoja = Empty
1686:             ws.Cells(filaVaciaFinal, "A").Select
1687:         End If
1688:         Exit Function
1689:     End If
1690: 
1691:     ioValorAnteriorHoja = numSolicitado
1692:     Call MarcarSiPagadosEnResidencia.SincronizarOrdenYCampos( _
1693:         ws.Name, "C", colPagoRes, nombreHojaResumen, colFactSum, colPagoSum, numSolicitado)
1694:     Call SincronizarFilaHaciaAccess(ws, targetRow, True)
1695:     GestionarEntradaNumOrdenHoja = True
1696: End Function

```

### 3.5. Procedimiento `AsegurarOrdenExiste` (`modDatabase.bas`, líneas 1698-1732)

```vba
1698: Public Function AsegurarOrdenExiste(ByVal numOrden As Long, ByVal residencia As String) As Boolean
1699:     If numOrden <= 0 Then Exit Function
1700:     
1701:     Dim resTarget As String
1702:     resTarget = ObtenerClaveResidenciaDesdeHoja(residencia)
1703:     If resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO" Then Exit Function
1704:     
1705:     Call AsegurarEsquemaConcurrenciaBD
1706:     
1707:     Dim sqlCheck As String
1708:     Dim varCount As Variant
1709:     sqlCheck = "SELECT COUNT(*) FROM Ordenes WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "'"
1710:     varCount = ExecuteScalar(sqlCheck)
1711:     
1712:     If Not IsNull(varCount) And Not IsEmpty(varCount) Then
1713:         If CLng(varCount) > 0 Then
1714:             AsegurarOrdenExiste = True
1715:             Exit Function
1716:         End If
1717:     End If
1718:     
1719:     Dim sqlInsert As String
1720:     Dim userCreacion As String, tokenActual As String
1721:     userCreacion = Trim$(ObtenerUsuarioActual())
1722:     If Len(userCreacion) = 0 Then userCreacion = "admin"
1723:     tokenActual = ObtenerTokenSesionInstancia()
1724:     
1725:     sqlInsert = "INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, UsuarioCreacion, " & _
1726:                 "UsuarioModificacion, UltimaModificacion, Solapamiento, ConsentimientoRGPD) " & _
1727:                 "VALUES (" & numOrden & ", '" & EscaparSQL(resTarget) & "', " & _
1728:                 FormatearFechaHoraSQL(Now) & ", '" & EscaparSQL(userCreacion) & "', '" & _
1729:                 EscaparSQL(tokenActual) & "', " & FormatearFechaHoraSQL(Now) & ", False, False)"
1730:                 
1731:     AsegurarOrdenExiste = ExecuteNonQuery(sqlInsert)
1732: End Function

```

### 3.6. Procedimiento `ExisteOrdenEnBD` (`modDatabase.bas`, líneas 2812-2885)

```vba
2812: Public Function ExisteOrdenEnBD( _
2813:     ByVal residencia As String, _
2814:     ByVal numOrden As Long, _
2815:     Optional ByRef outUsuarioCreacion As String = "", _
2816:     Optional ByRef outNombreHuesped As String = "", _
2817:     Optional ByRef outDNI As String = "", _
2818:     Optional ByRef outSiguienteLibreBD As Long = 0 _
2819: ) As Boolean
2820: 
2821:     Dim cn As Object
2822:     Dim rs As Object
2823:     Dim sql As String
2824:     Dim claveRes As String
2825: 
2826:     On Error GoTo ErrExisteOrden
2827:     ExisteOrdenEnBD = False
2828:     outUsuarioCreacion = ""
2829:     outNombreHuesped = ""
2830:     outDNI = ""
2831:     outSiguienteLibreBD = 0
2832: 
2833:     If numOrden <= 0 Then Exit Function
2834:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then Exit Function
2835: 
2836:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
2837:     If Len(claveRes) = 0 Then claveRes = UCase$(Trim$(residencia))
2838: 
2839:     Set cn = CreateObject("ADODB.Connection")
2840:     cn.ConnectionTimeout = 5
2841:     cn.Open GetConnectionString()
2842: 
2843:     sql = "SELECT UsuarioCreacion, Nombre, DNI FROM Ordenes " & _
2844:           "WHERE Residencia = '" & EscaparSQL(claveRes) & "' " & _
2845:           "AND NumOrden = " & numOrden
2846: 
2847:     Set rs = cn.Execute(sql)
2848:     If Not rs.EOF Then
2849:         ExisteOrdenEnBD = True
2850:         outUsuarioCreacion = Trim$(CStr(Nz(rs("UsuarioCreacion").Value, "")))
2851:         outNombreHuesped = Trim$(CStr(Nz(rs("Nombre").Value, "")))
2852:         outDNI = Trim$(CStr(Nz(rs("DNI").Value, "")))
2853:     End If
2854:     rs.Close
2855:     Set rs = Nothing
2856: 
2857:     ' Si existe conflicto, calcular también el siguiente Nº ORDEN libre real en Access
2858:     If ExisteOrdenEnBD Then
2859:         Set rs = cn.Execute("SELECT MAX(NumOrden) FROM Ordenes " & _
2860:                             "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden < 900001")
2861:         If Not rs.EOF And Not IsNull(rs(0).Value) Then
2862:             outSiguienteLibreBD = CLng(rs(0).Value) + 1
2863:         Else
2864:             outSiguienteLibreBD = numOrden + 1
2865:         End If
2866:         rs.Close
2867:         Set rs = Nothing
2868:     End If
2869: 
2870:     cn.Close
2871:     Set cn = Nothing
2872:     Exit Function
2873: 
2874: ErrExisteOrden:
2875:     On Error Resume Next
2876:     If Not rs Is Nothing Then
2877:         If rs.State = adStateOpen Then rs.Close
2878:         Set rs = Nothing
2879:     End If
2880:     If Not cn Is Nothing Then
2881:         If cn.State = adStateOpen Then cn.Close
2882:         Set cn = Nothing
2883:     End If
2884:     On Error GoTo 0
2885: End Function

```

### 3.7. Jerarquía de Llamadas, Tablas Consultadas y Comprobaciones de Duplicidad

1. **Árbol de Llamadas Documentado**:
   - `Hoja2.Worksheet_Change` (GIJÓN) / `Hoja20.Worksheet_Change` (SOTO) / `Hoja24.Worksheet_Change` (OVIEDO)
     - $\rightarrow$ `modDatabase.GestionarEntradaNumOrdenHoja`
       - **Caso A (Modificación/Renumeración)**:
         - $\rightarrow$ `modDatabase.ExisteOrdenEnBD` (comprueba si el nuevo número solicitado ya existe en Access)
         - $\rightarrow$ Si existe: revierte a `numPrevio` y muestra `MsgBox` de colisión.
         - $\rightarrow$ Si no existe: llama a `modDatabase.RenumerarOrdenExistenteBD` (`UPDATE Ordenes SET NumOrden = ...`), sincroniza resumen y llama a `SincronizarFilaHaciaAccess`.
       - **Caso B (Nueva Asignación / Celda A vacía)**:
         - $\rightarrow$ `modDatabase.IntentarReservarOrdenBD`
           - Si no hay red: llama a `ModuloSeguridadLocal.ObtenerSiguienteNumOrdenProvisional` (asigna $\ge 900001$) y encola `PRERESERVA`.
           - Si hay colisión en Access: limpia celda A, refresca hoja vía `SincronizarHojaDesdeAccess` y ofrece interactivamente `sigLibre`.
           - Si el operador acepta `sigLibre`: llama a `modDatabase.ReservarSiguienteOrdenLibreBD` (`SELECT MAX(NumOrden) + 1` con bucle de reintento transaccional).
           - Si la reserva tiene éxito: llama a `SincronizarFilaHaciaAccess(ws, fila, True)`.
2. **Tablas Consultadas**:
   - **Únicamente `Ordenes`**. Ninguno de estos procedimientos consulta jamás `Ordenes_Historico`.
3. **Comprobaciones de Duplicidad**:
   - `ExisteOrdenEnBD`: `SELECT UsuarioCreacion, Nombre, DNI FROM Ordenes WHERE Residencia = '...' AND NumOrden = ...`
   - `IntentarReservarOrdenBD`: `SELECT UsuarioCreacion, UsuarioModificacion, Nombre FROM Ordenes WHERE Residencia = '...' AND NumOrden = ...`
   - `ReservarSiguienteOrdenLibreBD`: `SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '...' AND NumOrden < 900001`
   - **Hecho acreditado**: No se comprueba el año del expediente en ninguna sentencia SQL. Si un número de orden fue archivado a `Ordenes_Historico` o pertenecía a un ejercicio anterior, para todos los procedimientos de numeración online figura como **completamente libre** y se vuelve a adjudicar de inmediato.

---

## 4. EXTRACCIÓN Y ANÁLISIS DEL Nº FACTURA (SECCIÓN 6)

### 4.1. Procedimientos Principales Extraídos

#### A. `AsignarFacturaPorColumna` (`AsignarNumFactura.bas`, líneas 14-232)
```vba
  14: Public Sub AsignarFacturaPorColumna(ws As Worksheet, _
  15:                                     ByVal Target As Range, _
  16:                                     ByVal columnaCheck As String, _
  17:                                     Optional ByVal columnaFactura As String = "C")
  18: 
  19:     Dim celda As Range
  20:     Dim nuevoNumero As Long
  21:     Dim numeroOrden As Variant
  22:     Dim respuesta As VbMsgBoxResult
  23:     Dim nombreHojaFactura As String
  24:     Dim hojaFactura As Worksheet
  25:     Dim valorResolucion As String
  26:     Dim mensajeAsignacion As String
  27:     
  28:     ' OPTIMIZACIÓN v2.0: Variables para sincronización
  29:     Dim nombreHojaResidencia As String
  30:     Dim nombreHojaResumen As String
  31:     Dim columnaPagadoResidencia As String
  32:     Dim columnaPagadoResumen As String
  33:     Dim columnaFacturaResumen As String
  34:     Dim resCode As String
  35: 
  36:     ' ==============================================================================
  37:     ' Determinar código de residencia por CodeName (inmune a nombres de pestaña y codificación)
  38:     ' con fallback por Name mediante modDatabase.ObtenerClaveResidenciaDesdeHoja
  39:     ' ==============================================================================
  40:     Select Case ws.CodeName
  41:         Case "Hoja2"
  42:             resCode = "GIJON"
  43:         Case "Hoja20"
  44:             resCode = "SOTO"
  45:         Case "Hoja24"
  46:             resCode = "OVIEDO"
  47:         Case Else
  48:             resCode = modDatabase.ObtenerClaveResidenciaDesdeHoja(ws.Name)
  49:     End Select
  50: 
  51:     ' Determinar las hojas y columnas correspondientes según la residencia
  52:     Select Case resCode
  53:         Case "GIJON"
  54:             nombreHojaFactura = "FACTURA GIJÓN"
  55:             nombreHojaResidencia = "RESIDENCIA GIJÓN"
  56:             nombreHojaResumen = "RESUMEN GIJÓN"
  57:             columnaPagadoResidencia = "AA"
  58:             columnaPagadoResumen = "M"
  59:             columnaFacturaResumen = "N"
  60:             
  61:         Case "SOTO"
  62:             nombreHojaFactura = "FACTURA SOTO"
  63:             nombreHojaResidencia = "RESIDENCIA SOTO"
  64:             nombreHojaResumen = "RESUMEN SOTO"
  65:             columnaPagadoResidencia = "AA"
  66:             columnaPagadoResumen = "L"
  67:             columnaFacturaResumen = "M"
  68:             
  69:         Case "OVIEDO"
  70:             nombreHojaFactura = "FACTURA OVIEDO"
  71:             nombreHojaResidencia = "RESIDENCIA OVIEDO"
  72:             nombreHojaResumen = "RESUMEN OVIEDO"
  73:             columnaPagadoResidencia = "AB"
  74:             columnaPagadoResumen = "M"
  75:             columnaFacturaResumen = "N"
  76:             
  77:         Case Else
  78:             MsgBox "Esta hoja no está habilitada para el proceso de facturación automática.", vbExclamation
  79:             Exit Sub
  80:     End Select
  81: 
  82:     ' Intentar obtener la hoja de factura
  83:     On Error Resume Next
  84:     Set hojaFactura = ws.Parent.Worksheets(nombreHojaFactura)
  85:     If hojaFactura Is Nothing And resCode = "GIJON" Then
  86:         Set hojaFactura = ws.Parent.Worksheets("FACTURA GIJON")
  87:     End If
  88:     On Error GoTo 0
  89: 
  90:     If hojaFactura Is Nothing Then
  91:         MsgBox "No existe la hoja de factura llamada '" & nombreHojaFactura & "'.", vbCritical
  92:         Exit Sub
  93:     End If
  94: 
  95:     ' ' Comprobar si el cambio afecta a la columna de chequeo
  96:     If Not Intersect(Target, ws.Columns(columnaCheck)) Is Nothing Then
  97:         On Error GoTo CleanExit_Fact
  98:         Application.enableEvents = False
  99: 
 100:         For Each celda In Intersect(Target, ws.Columns(columnaCheck))
 101:             ' Usamos la validación centralizada de pagos para SMS, Bizum, Transferencia, Tarjeta, Efectivo y SI
 102:             If MarcarSiPagadosEnResidencia.EsEstadoPagado(celda.Value) Then
 103:                 ' Validar columna RESOLUCION (P = columna 16)
 104:                 valorResolucion = UCase$(Trim$(CStr(ws.Cells(celda.Row, 16).Value)))
 105:                 
 106:                 If valorResolucion = "SI" Or valorResolucion = "CONCEDIDA" Or valorResolucion = "REEVALUADA" Then
 107:                     ' Capturar el Nº ORDEN
 108:                     numeroOrden = ws.Cells(celda.Row, 1).Value
 109: 
 110:                     ' === [SICO-N2: FACTURA PROFORMA EN MODO OFFLINE Y ENCOLADO FIFO] ===
 111:                     If ModuloSeguridadLocal.EstaEnModoContingencia() Or Not ModuloSeguridadLocal.SondaRedDisponible(modDatabase.DB_PATH) Then
 112:                         Dim pregProforma As VbMsgBoxResult
 113:                         pregProforma = MsgBox( _
 114:                             "MODO FUERA DE LÍNEA (SICO-N2):" & vbCrLf & vbCrLf & _
 115:                             "El servidor central no está disponible en este momento." & vbCrLf & _
 116:                             "¿Desea generar el documento con la marca 'PROFORMA / PENDIENTE DE NUMERACIÓN OFICIAL' " & _
 117:                             "para el Nº ORDEN " & numeroOrden & " y encolar la solicitud en FIFO para asignarle " & _
 118:                             "su número oficial definitivo automáticamente al recuperar la conexión?", _
 119:                             vbYesNo + vbQuestion, "Emisión Proforma Offline - SICO-N2")
 120:                             
 121:                         If pregProforma = vbYes Then
 122:                             ws.Cells(celda.Row, columnaFactura).Value = "PROFORMA"
 123:                             
 124:                             If IsNumeric(numeroOrden) And CLng(numeroOrden) > 0 Then
 125:                                 Call ModuloSeguridadLocal.EncolarOperacionOffline( _
 126:                                     resCode, CLng(numeroOrden), "UPDATE_CAMPO", "EstadoPago", "", CStr(celda.Value))
 127:                                 Call ModuloSeguridadLocal.EncolarOperacionOffline( _
 128:                                     resCode, CLng(numeroOrden), "ASIGNAR_FACTURA_DIFERIDA", "NumFactura", "PROFORMA", "SOLICITAR_OFICIAL")
 129:                             End If
 130:                             
 131:                             If Trim$(CStr(numeroOrden)) <> "" Then
 132:                                 Call MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen( _
 133:                                     nombreHojaResidencia, columnaFactura, columnaPagadoResidencia, _
 134:                                     nombreHojaResumen, columnaFacturaResumen, columnaPagadoResumen, numeroOrden)
 135:                             End If
 136:                             
 137:                             respuesta = MsgBox("Documento registrado como 'PROFORMA / PENDIENTE DE NUMERACIÓN OFICIAL' " & _
 138:                                                "y encolado en el búfer FIFO." & vbCrLf & vbCrLf & _
 139:                                                "¿Desea abrir e imprimir el documento proforma?", _
 140:                                                vbYesNo + vbQuestion, "Imprimir Proforma")
 141:                             If respuesta = vbYes Then
 142:                                 Call modDatabase.MostrarYActivarHojaSegura(nombreHojaFactura)
 143:                                 hojaFactura.Range("D1").Value = numeroOrden
 144:                             End If
 145:                         End If
 146:                         GoTo SiguienteCeldaFactura
 147:                     End If
 148:                     ' === [FIN BLOQUE SICO-N2 PROFORMA OFFLINE] ===
 149:                     
 150:                     ' Previsualizar el próximo número sin incrementar aún el contador físico en Access
 151:                     Dim numPropuesto As Long
 152:                     Dim varUlt As Variant
 153:                     varUlt = modDatabase.ExecuteScalar("SELECT UltimoNumero FROM ContadorFacturas WHERE Residencia = '" & _
 154:                                                        modDatabase.EscaparSQL(resCode) & "' AND Ejercicio = " & Year(Now))
 155:                     If IsNumeric(varUlt) And Not IsNull(varUlt) Then
 156:                         numPropuesto = CLng(varUlt) + 1
 157:                     Else
 158:                         numPropuesto = 1
 159:                     End If
 160:                     mensajeAsignacion = "Asignación de factura única desde Access para " & resCode
 161:                     
 162:                     ' Detectar si existen huecos en la serie de facturación
 163:                     Dim textoHuecos As String
 164:                     textoHuecos = DetectarHuecosEnSerie(ws, columnaFactura, numPropuesto)
 165:                     
 166:                     Dim mensajePrompt As String
 167:                     mensajePrompt = "Se propone asignar al Nº ORDEN " & numeroOrden & _
 168:                                     " la factura Nº " & numPropuesto & vbCrLf & vbCrLf & _
 169:                                     mensajeAsignacion
 170:                     
 171:                     If Len(textoHuecos) > 0 Then
 172:                         mensajePrompt = mensajePrompt & vbCrLf & vbCrLf & _
 173:                                         "ADVERTENCIA. En la serie, faltan las facturas Nº:" & vbCrLf & _
 174:                                         textoHuecos & vbCrLf & _
 175:                                         "Valore si corresponde mantener este salto o rellenar huecos."
 176:                     End If
 177:                     
 178:                     mensajePrompt = mensajePrompt & vbCrLf & vbCrLf & _
 179:                                     "¿Desea adjudicar este número de factura?"
 180:                     
 181:                     ' Preguntar al usuario ANTES de incrementar el contador atómico en Access
 182:                     Dim preguntaAsignar As VbMsgBoxResult
 183:                     preguntaAsignar = MsgBox(mensajePrompt, _
 184:                                              vbYesNo + vbQuestion, "Confirmar Adjudicación de Factura")
 185: 
 186:                     If preguntaAsignar = vbYes Then
 187:                         ' === Ahora sí: Obtener e incrementar número atómico desde Access ===
 188:                         nuevoNumero = modDatabase.ObtenerSiguienteNumFactura(resCode, Year(Now))
 189:                         If nuevoNumero <= 0 Then GoTo CleanExit_Fact
 190: 
 191:                         ' Asignar el número de factura en Access (fuente de verdad) Y en la celda local
 192:                         If IsNumeric(numeroOrden) And CLng(numeroOrden) > 0 Then
 193:                             modDatabase.ActualizarOrden CLng(numeroOrden), "NumFactura", CStr(nuevoNumero), resCode
 194:                         End If
 195:                         ws.Cells(celda.Row, columnaFactura).Value = nuevoNumero
 196:                         
 197:                         ' ========================================================================
 198:                         ' OPTIMIZACIÓN v2.0: SINCRONIZAR AUTOMÁTICAMENTE CON RESUMEN
 199:                         ' Sincroniza C (Nº FACTURA) y AA/AB (PAGADO) con el RESUMEN
 200:                         ' ========================================================================
 201:                         If Trim$(CStr(numeroOrden)) <> "" Then
 202:                             Call MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen( _
 203:                                 nombreHojaResidencia, _
 204:                                 columnaFactura, _
 205:                                 columnaPagadoResidencia, _
 206:                                 nombreHojaResumen, _
 207:                                 columnaFacturaResumen, _
 208:                                 columnaPagadoResumen, _
 209:                                 numeroOrden)
 210:                         End If
 211:                         ' ========================================================================
 212:                         
 213:                         ' Preguntar si desea imprimir la factura
 214:                         respuesta = MsgBox("Nº de factura " & nuevoNumero & " asignado correctamente." & vbCrLf & vbCrLf & _
 215:                                            "¿Desea imprimir la factura?", vbYesNo + vbQuestion, "Imprimir factura")
 216:                         If respuesta = vbYes Then
 217:                             Call modDatabase.MostrarYActivarHojaSegura(nombreHojaFactura)
 218:                             hojaFactura.Range("D1").Value = numeroOrden
 219:                         End If
 220:                     End If
 221:                 Else
 222:                     MsgBox "No es posible generar una factura con la resolución adoptada (" & ws.Cells(celda.Row, 16).Value & ").", vbExclamation, "Factura no generada"
 223:                     celda.Value = ""
 224:                 End If
 225:             End If
 226: SiguienteCeldaFactura:
 227:         Next celda
 228: 
 229: CleanExit_Fact:
 230:         Application.enableEvents = True
 231:     End If
 232: End Sub

```

#### B. `ProcesarCambioManualFactura` (`AsignarNumFactura.bas`, líneas 561-641)
```vba
 561: Public Sub ProcesarCambioManualFactura(ws As Worksheet, Target As Range, ByVal targetRow As Long, Optional ByVal columnaFactura As String = "C")
 562:     Dim numOrden As Variant
 563:     Dim nuevoValor As Variant
 564:     Dim resCode As String
 565:     Dim numFacturaLong As Long
 566:     Dim ultimaFila As Long
 567:     Dim rangoCol As Range
 568:     
 569:     If Target.Cells.CountLarge > 1 Or targetRow < 2 Then Exit Sub
 570:     
 571:     ' Capturar el Nº ORDEN
 572:     numOrden = ws.Cells(targetRow, 1).Value
 573:     If IsEmpty(numOrden) Or Not IsNumeric(numOrden) Or CLng(numOrden) <= 0 Then
 574:         If Not IsEmpty(Target.Value) Then
 575:             MsgBox "Debe existir un Nº ORDEN válido antes de asignar un número de factura.", vbExclamation, "Sin Nº ORDEN"
 576:             Application.enableEvents = False
 577:             Target.ClearContents
 578:             Application.enableEvents = True
 579:         End If
 580:         Exit Sub
 581:     End If
 582:     
 583:     ' Determinar código de residencia
 584:     Select Case ws.CodeName
 585:         Case "Hoja2":  resCode = "GIJON"
 586:         Case "Hoja20": resCode = "SOTO"
 587:         Case "Hoja24": resCode = "OVIEDO"
 588:         Case Else:     resCode = modDatabase.ObtenerClaveResidenciaDesdeHoja(ws.Name)
 589:     End Select
 590:     
 591:     nuevoValor = Target.Value
 592:     
 593:     ' Caso 1: Celda vaciada (borrado manual de número de factura)
 594:     If IsEmpty(nuevoValor) Or Trim(CStr(nuevoValor)) = "" Then
 595:         modDatabase.ActualizarOrden CLng(numOrden), "NumFactura", "", resCode
 596: 
 597:         Exit Sub
 598:     End If
 599:     
 600:     ' Caso 2: Validación de valor numérico
 601:     If Not IsNumeric(nuevoValor) Then
 602:         MsgBox "El número de factura debe ser un valor numérico entero positivo.", vbExclamation, "Valor no válido"
 603:         Application.enableEvents = False
 604:         Target.ClearContents
 605:         Application.enableEvents = True
 606:         Exit Sub
 607:     End If
 608:     
 609:     numFacturaLong = CLng(nuevoValor)
 610:     If numFacturaLong <= 0 Then
 611:         MsgBox "El número de factura debe ser mayor que cero.", vbExclamation, "Valor no válido"
 612:         Application.enableEvents = False
 613:         Target.ClearContents
 614:         Application.enableEvents = True
 615:         Exit Sub
 616:     End If
 617:     
 618:     ' 1. Comprobación local de duplicados en la hoja Excel activa
 619:     ultimaFila = ws.Cells(ws.Rows.Count, columnaFactura).End(xlUp).Row
 620:     Set rangoCol = ws.Range(columnaFactura & "1:" & columnaFactura & ultimaFila)
 621:     Call DuplicadosOrdenyFactura.VerificarDuplicados(Target, rangoCol, "NUM. FACT.")
 622:     
 623:     If IsEmpty(Target.Value) Then Exit Sub
 624:     
 625:     ' 2. Comprobación anti-colisión en Access (por si otro usuario ya lo asignó)
 626:     If modDatabase.ExisteFacturaEnBD(resCode, numFacturaLong, CLng(numOrden)) Then
 627:         MsgBox "El Nº de factura " & numFacturaLong & " YA ESTÁ ASIGNADO en la base de datos para " & resCode & "." & vbCrLf & vbCrLf & _
 628:                "Operación cancelada para evitar duplicidad entre puestos.", vbCritical, "Factura Duplicada en Access"
 629:         Application.enableEvents = False
 630:         Target.ClearContents
 631:         Application.enableEvents = True
 632:         Exit Sub
 633:     End If
 634:     
 635:     ' 3. Advertencia informativa de huecos en la serie
 636:     Call AsignarNumFactura.ValidarFacturaManual(ws, numFacturaLong, columnaFactura)
 637:     
 638:     ' 4. Persistir inmediatamente en Access y alinear ContadorFacturas si el Nº manual adelanta la serie
 639:     modDatabase.ActualizarOrden CLng(numOrden), "NumFactura", CStr(numFacturaLong), resCode
 640:     modDatabase.SincronizarContadorFacturaManual resCode, numFacturaLong, Year(Now)
 641: End Sub

```

#### C. `ValidarFacturaManual` (`AsignarNumFactura.bas`, líneas 649-769)
```vba
 649: Public Sub ValidarFacturaManual(ws As Worksheet, numeroEscrito As Long, columnaFactura As String)
 650:     Dim numeros() As Long
 651:     Dim bloques As Collection
 652:     Dim bloque As Collection
 653:     Dim i As Long, ultimaFila As Long
 654:     Dim valor As Variant
 655:     Dim contador As Long
 656:     Dim minBloque As Long, maxBloque As Long
 657:     Dim huecos As String, infoOtrosBloques As String
 658:     Dim numHueco As Long
 659:     Dim encontrado As Boolean
 660:     Dim numB As Variant
 661:     Dim esNuevoBloque As Boolean
 662: 
 663:     ' === 1. ESCANEAR Números EXISTENTES (excluyendo el recién escrito) ===
 664:     ultimaFila = ws.Cells(ws.Rows.Count, columnaFactura).End(xlUp).Row
 665:     ReDim numeros(1 To ultimaFila)
 666:     contador = 0
 667: 
 668:     For i = 2 To ultimaFila
 669:         valor = ws.Cells(i, columnaFactura).Value
 670:         If IsNumeric(valor) Then
 671:             If CLng(valor) > 0 And CLng(valor) <> numeroEscrito Then
 672:                 contador = contador + 1
 673:                 numeros(contador) = CLng(valor)
 674:             End If
 675:         End If
 676:     Next i
 677: 
 678:     ' Si no hay otras facturas, no hay nada que validar
 679:     If contador = 0 Then Exit Sub
 680: 
 681:     ReDim Preserve numeros(1 To contador)
 682: 
 683:     ' === 2. ORDENAR Y AGRUPAR EN BLOQUES ===
 684:     Call OrdenarArray(numeros)
 685:     Set bloques = AgruparEnBloques(numeros)
 686: 
 687:     ' === 3. COMPROBAR SI EL número ESCRITO INICIA UN NUEVO BLOQUE ===
 688:     ' Un número inicia nuevo bloque si el mínimo de todos los bloques
 689:     ' existentes es MÁS de 100 veces mayor que el número escrito
 690:     Dim minimoGlobal As Long
 691:     minimoGlobal = numeros(1) ' Ya están ordenados, el primero es el mínimo global
 692:     esNuevoBloque = (minimoGlobal > numeroEscrito * 100)
 693: 
 694:     ' === 4A. SI ES NUEVO BLOQUE: solo informar huecos en bloques existentes ===
 695:     If esNuevoBloque Then
 696:         infoOtrosBloques = ""
 697:         For Each bloque In bloques
 698:             minBloque = bloque(1)
 699:             maxBloque = bloque(bloque.Count)
 700:             huecos = ""
 701:             For numHueco = minBloque To maxBloque
 702:                 encontrado = False
 703:                 For Each numB In bloque
 704:                     If CLng(numB) = numHueco Then
 705:                         encontrado = True
 706:                         Exit For
 707:                     End If
 708:                 Next numB
 709:                 If Not encontrado Then
 710:                     If huecos = "" Then
 711:                         huecos = CStr(numHueco)
 712:                     Else
 713:                         huecos = huecos & ", " & CStr(numHueco)
 714:                     End If
 715:                 End If
 716:             Next numHueco
 717:             If huecos <> "" Then
 718:                 If infoOtrosBloques = "" Then
 719:                     infoOtrosBloques = "**Huecos en bloques anteriores:" & vbCrLf
 720:                 End If
 721:                 infoOtrosBloques = infoOtrosBloques & "  Bloque [" & minBloque & "-" & maxBloque & "]: falta " & huecos & vbCrLf
 722:             End If
 723:         Next bloque
 724: 
 725:         If infoOtrosBloques <> "" Then
 726:             MsgBox "Nº de factura " & numeroEscrito & " inicia una nueva serie." & vbCrLf & vbCrLf & _
 727:                    infoOtrosBloques, vbInformation, "Nueva serie de facturas"
 728:         End If
 729:         Exit Sub
 730:     End If
 731: 
 732:     ' === 4B. SI NO ES NUEVO BLOQUE: comportamiento original con huecos ===
 733:     Dim bloquePertenece As Collection
 734:     Set bloquePertenece = Nothing
 735: 
 736:     For Each bloque In bloques
 737:         minBloque = bloque(1)
 738:         maxBloque = bloque(bloque.Count)
 739:         If (numeroEscrito >= minBloque And numeroEscrito <= maxBloque) Or _
 740:            (numeroEscrito > maxBloque And numeroEscrito - maxBloque <= 10) Then
 741:             Set bloquePertenece = bloque
 742:             Exit For
 743:         End If
 744:     Next bloque
 745: 
 746:     If bloquePertenece Is Nothing Then Exit Sub
 747: 
 748:     maxBloque = bloquePertenece(bloquePertenece.Count)
 749: 
 750:     If numeroEscrito > maxBloque Then
 751:         huecos = ""
 752:         For numHueco = maxBloque + 1 To numeroEscrito - 1
 753:             If huecos = "" Then
 754:                 huecos = CStr(numHueco)
 755:             Else
 756:                 huecos = huecos & ", " & CStr(numHueco)
 757:             End If
 758:         Next numHueco
 759: 
 760:         If huecos <> "" Then
 761:             MsgBox "ADVERTENCIA: Números de factura faltantes" & vbCrLf & vbCrLf & _
 762:                    "Ha introducido el número: " & numeroEscrito & vbCrLf & _
 763:                    "Nº factura máxima anterior del bloque: " & maxBloque & vbCrLf & vbCrLf & _
 764:                    "Números faltantes: " & huecos & vbCrLf & vbCrLf & _
 765:                    "Valore rellenar estos huecos para mantener la secuencia numérica.", _
 766:                    vbExclamation, "Huecos en Numeración de Facturas"
 767:         End If
 768:     End If
 769: End Sub

```

#### D. `ObtenerSiguienteNumFactura` (`modDatabase.bas`, líneas 2678-2778)
```vba
2678: Public Function ObtenerSiguienteNumFactura( _
2679:     ByVal residencia As String, _
2680:     ByVal ejercicio As Integer _
2681: ) As Long
2682:     Dim cn As Object
2683:     Dim rs As Object
2684:     Dim nuevoNum As Long
2685:     Dim maxEnOrdenes As Long
2686:     Dim intentos As Integer
2687:     Dim exito As Boolean
2688:     Dim ultErrorDesc As String
2689:     Dim claveRes As String
2690:     
2691:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
2692:     If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
2693:     
2694:     intentos = 5
2695:     exito = False
2696:     nuevoNum = 0
2697:     
2698:     Do While intentos > 0 And Not exito
2699:         On Error Resume Next
2700:         Err.Clear
2701:         
2702:         Set cn = CreateObject("ADODB.Connection")
2703:         cn.ConnectionTimeout = 5
2704:         cn.Open GetConnectionString()
2705:         
2706:         If Err.Number = 0 Then
2707:             cn.BeginTrans
2708:             
2709:             ' 1. Asegurar que ContadorFacturas no haya quedado por detrás de una factura manual
2710:             maxEnOrdenes = 0
2711:             Set rs = cn.Execute("SELECT MAX(NumFactura) FROM Ordenes WHERE Residencia = '" & EscaparSQL(claveRes) & "'")
2712:             If Err.Number = 0 And Not rs Is Nothing Then
2713:                 If Not rs.EOF And Not IsNull(rs(0).Value) Then
2714:                     maxEnOrdenes = CLng(rs(0).Value)
2715:                 End If
2716:                 rs.Close
2717:                 Set rs = Nothing
2718:             End If
2719:             Err.Clear
2720:             
2721:             If maxEnOrdenes > 0 Then
2722:                 cn.Execute "UPDATE ContadorFacturas SET UltimoNumero = " & maxEnOrdenes & _
2723:                            " WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND Ejercicio = " & ejercicio & _
2724:                            " AND UltimoNumero < " & maxEnOrdenes
2725:             End If
2726:             
2727:             ' 2. Incrementar el contador atómicamente
2728:             cn.Execute "UPDATE ContadorFacturas SET UltimoNumero = UltimoNumero + 1 " & _
2729:                        "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND Ejercicio = " & ejercicio
2730:             
2731:             ' 3. Leer el nuevo valor asignado
2732:             If Err.Number = 0 Then
2733:                 Set rs = cn.Execute("SELECT UltimoNumero FROM ContadorFacturas " & _
2734:                                     "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND Ejercicio = " & ejercicio)
2735:                 If Err.Number = 0 And Not rs Is Nothing Then
2736:                     If Not rs.EOF And Not IsNull(rs(0).Value) Then
2737:                         nuevoNum = CLng(rs(0).Value)
2738:                     End If
2739:                     rs.Close
2740:                     Set rs = Nothing
2741:                 End If
2742:             End If
2743:             
2744:             If Err.Number = 0 And nuevoNum > 0 Then
2745:                 cn.CommitTrans
2746:                 If Err.Number = 0 Then
2747:                     exito = True
2748:                 Else
2749:                     ultErrorDesc = Err.Description
2750:                     cn.RollbackTrans
2751:                 End If
2752:             Else
2753:                 ultErrorDesc = Err.Description
2754:                 cn.RollbackTrans
2755:             End If
2756:             
2757:             cn.Close
2758:             Set cn = Nothing
2759:         Else
2760:             ultErrorDesc = Err.Description
2761:         End If
2762:         
2763:         On Error GoTo 0
2764:         
2765:         If Not exito Then
2766:             intentos = intentos - 1
2767:             If intentos > 0 Then PausaMS 60
2768:         End If
2769:     Loop
2770:     
2771:     If exito Then
2772:         ObtenerSiguienteNumFactura = nuevoNum
2773:     Else
2774:         MsgBox "Error al obtener número de factura tras varios intentos:" & vbCrLf & ultErrorDesc, _
2775:                vbCritical, "Error de Facturación"
2776:         ObtenerSiguienteNumFactura = 0
2777:     End If
2778: End Function

```

#### E. `SincronizarContadorFacturaManual` (`modDatabase.bas`, líneas 2782-2801)
```vba
2782: Public Sub SincronizarContadorFacturaManual( _
2783:     ByVal residencia As String, _
2784:     ByVal numFacturaManual As Long, _
2785:     Optional ByVal ejercicio As Integer = 0 _
2786: )
2787:     On Error Resume Next
2788:     If numFacturaManual <= 0 Then Exit Sub
2789:     If ejercicio <= 0 Then ejercicio = Year(Now)
2790:     
2791:     Dim claveRes As String
2792:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
2793:     If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
2794:     
2795:     Dim sql As String
2796:     sql = "UPDATE ContadorFacturas SET UltimoNumero = " & numFacturaManual & _
2797:           " WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND Ejercicio = " & ejercicio & _
2798:           " AND UltimoNumero < " & numFacturaManual
2799:     Call ExecuteNonQuery(sql)
2800:     On Error GoTo 0
2801: End Sub

```

#### F. `ExisteFacturaEnBD` (`modDatabase.bas`, líneas 2889-2916)
```vba
2889: Public Function ExisteFacturaEnBD( _
2890:     ByVal residencia As String, _
2891:     ByVal numFactura As Long, _
2892:     Optional ByVal excluirNumOrden As Long = 0) As Boolean
2893:     
2894:     Dim sql As String
2895:     Dim resCount As Variant
2896:     Dim claveRes As String
2897:     
2898:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
2899:     If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
2900:     
2901:     sql = "SELECT COUNT(*) FROM Ordenes " & _
2902:           "WHERE Residencia = '" & EscaparSQL(claveRes) & "' " & _
2903:           "AND NumFactura = " & numFactura
2904:     
2905:     If excluirNumOrden > 0 Then
2906:         sql = sql & " AND NumOrden <> " & excluirNumOrden
2907:     End If
2908:     
2909:     resCount = ExecuteScalar(sql)
2910:     
2911:     If IsNull(resCount) Then
2912:         ExisteFacturaEnBD = False
2913:     Else
2914:         ExisteFacturaEnBD = (CLng(resCount) > 0)
2915:     End If
2916: End Function

```

#### G. `RefrescarColumnaFacturasDesdeAccess` (`modDatabase.bas`, líneas 2920-3010)
```vba
2920: Public Sub RefrescarColumnaFacturasDesdeAccess(ws As Worksheet, ByVal residencia As String)
2921:     Dim cn As Object
2922:     Dim rs As Object
2923:     Dim sql As String
2924:     Dim dictFacturas As Object
2925:     Dim ultimaFila As Long
2926:     Dim r As Long
2927:     Dim numOrd As Variant
2928:     Dim arrA As Variant
2929:     Dim arrC As Variant
2930:     Dim huboCambios As Boolean
2931:     Dim claveRes As String
2932:     
2933:     On Error GoTo ErrorHandler
2934:     
2935:     ' === [SICO-N2: Salida instantánea si estamos offline al activar pestaña] ===
2936:     If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then Exit Sub
2937:     
2938:     claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
2939:     If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
2940:     
2941:     ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
2942:     If ultimaFila < 2 Then Exit Sub
2943:     
2944:     ' 1. Consultar únicamente óórdenes con factura de esta residencia
2945:     sql = "SELECT NumOrden, NumFactura FROM Ordenes " & _
2946:           "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumFactura > 0"
2947:     
2948:     Set cn = CreateObject("ADODB.Connection")
2949:     cn.Open GetConnectionString()
2950:     Set rs = cn.Execute(sql)
2951:     
2952:     Set dictFacturas = CreateObject("Scripting.Dictionary")
2953:     Do While Not rs.EOF
2954:         If Not IsNull(rs("NumOrden").Value) And Not IsNull(rs("NumFactura").Value) Then
2955:             dictFacturas(CLng(rs("NumOrden").Value)) = CLng(rs("NumFactura").Value)
2956:         End If
2957:         rs.MoveNext
2958:     Loop
2959:     rs.Close
2960:     cn.Close
2961:     Set rs = Nothing
2962:     Set cn = Nothing
2963:     
2964:     ' 2. Comparación en memoria
2965:     If ultimaFila = 2 Then
2966:         numOrd = ws.Cells(2, "A").Value
2967:         If IsNumeric(numOrd) And CLng(numOrd) > 0 Then
2968:             If dictFacturas.Exists(CLng(numOrd)) Then
2969:                 If ws.Cells(2, "C").Value <> dictFacturas(CLng(numOrd)) Then
2970:                     ws.Cells(2, "C").Value = dictFacturas(CLng(numOrd))
2971:                 End If
2972:             End If
2973:         End If
2974:         Exit Sub
2975:     End If
2976:     
2977:     arrA = ws.Range("A2:A" & ultimaFila).Value
2978:     arrC = ws.Range("C2:C" & ultimaFila).Value
2979:     huboCambios = False
2980:     
2981:     For r = 1 To UBound(arrA, 1)
2982:         numOrd = arrA(r, 1)
2983:         If IsNumeric(numOrd) And Not IsEmpty(numOrd) Then
2984:             If CLng(numOrd) > 0 Then
2985:                 If dictFacturas.Exists(CLng(numOrd)) Then
2986:                     If arrC(r, 1) <> dictFacturas(CLng(numOrd)) Then
2987:                         arrC(r, 1) = dictFacturas(CLng(numOrd))
2988:                         huboCambios = True
2989:                     End If
2990:                 End If
2991:             End If
2992:         End If
2993:     Next r
2994:     
2995:     ' 3. Volcado en bloque solo si hubo cambios
2996:     If huboCambios Then
2997:         Application.enableEvents = False
2998:         ws.Range("C2:C" & ultimaFila).Value = arrC
2999:         Application.enableEvents = True
3000:     End If
3001:     
3002:     Exit Sub
3003:     
3004: ErrorHandler:
3005:     If Not rs Is Nothing Then Set rs = Nothing
3006:     If Not cn Is Nothing Then
3007:         If cn.State = adStateOpen Then cn.Close
3008:         Set cn = Nothing
3009:     End If
3010: End Sub

```

### 4.2. Análisis Técnico de la Numeración de Facturas

1. **Asignación Automática**:
   - Se dispara en `Worksheet_Change` al modificar la columna de pago (`AA` en GIJÓN/SOTO, `AB` en OVIEDO) si el valor introducido es reconocido como pagado (`MarcarSiPagadosEnResidencia.EsEstadoPagado`) y la resolución es concedida (`Resolucion` = SI, CONCEDIDA o REEVALUADA).
   - Si no hay red: asigna la marca textual `"PROFORMA"` en la celda y encola en SICO-N2 la operación `ASIGNAR_FACTURA_DIFERIDA`.
   - Si hay red: lee de la tabla `ContadorFacturas` el valor para `Ejercicio = Year(Now)`, previsualiza el número, avisa de posibles huecos en la serie de la hoja mediante `DetectarHuecosEnSerie`, y tras confirmación interactiva, incrementa atómicamente el contador en Access llamando a `ObtenerSiguienteNumFactura(resCode, Year(Now))`.
   - Escribe el número en la celda de la Columna C, actualiza el registro en Access vía `modDatabase.ActualizarOrden CLng(numeroOrden), "NumFactura", CStr(nuevoNumero), resCode`, y sincroniza con la columna de factura de la hoja RESUMEN.
2. **Asignación Manual**:
   - Se dispara en `Worksheet_Change` cuando el usuario edita directamente cualquier celda de la Columna C (`Nº FACTURA`).
   - Si la celda se vacía: borra el número de factura en Access ejecutando `ActualizarOrden(numOrden, "NumFactura", "", resCode)` (pone `NumFactura = Null`).
   - Si se introduce un entero positivo:
     1. Valida duplicados en la hoja local con `DuplicadosOrdenyFactura.VerificarDuplicados`.
     2. Valida colisión en Access mediante `modDatabase.ExisteFacturaEnBD(resCode, numFacturaLong, CLng(numOrden))`.
     3. Si ya existe en Access: cancela inmediatamente la operación, borra la celda y muestra aviso crítico de duplicidad multiusuario.
     4. Detecta huecos en la serie local con `ValidarFacturaManual`.
     5. Persiste en Access vía `ActualizarOrden` y alinea el contador físico en `ContadorFacturas` mediante `SincronizarContadorFacturaManual` para evitar que futuras asignaciones automáticas colisionen con este número adelantado.
3. **Columnas y Celdas de Excel Implicadas**:
   - En las tres residencias (GIJÓN, SOTO y OVIEDO): **Columna C (columna 3)**.
   - En las hojas RESUMEN: Columna **N** (GIJÓN), Columna **M** (SOTO), Columna **N** (OVIEDO).
4. **Campos de Access Correspondientes**:
   - Tabla `Ordenes`: Campo `NumFactura` (Tipo `dbLong`, 4 bytes, numérico entero largo).
   - Tabla `Ordenes_Historico`: Campo `NumFactura` (Tipo `dbLong`, 4 bytes).
   - Tabla auxiliar `ContadorFacturas`:
     - `Residencia` (`dbText`, 20)
     - `Ejercicio` (`dbLong`, 4)
     - `UltimoNumero` (`dbLong`, 4)
5. **Determinación del Año de Facturación (Ausencia de Implementación Estructural)**:
   - **Hecho acreditado**: En el código VBA, el año de facturación se obtiene exclusivamente en tiempo de ejecución mediante la función de reloj del sistema `Year(Now)` o `Year(CDate(fechaHoraLocal))` en el vaciado FIFO.
   - **Hecho acreditado en Access**: En las tablas `Ordenes` y `Ordenes_Historico` **NO EXISTE NINGÚN CAMPO** para almacenar el año de facturación, ni la fecha de factura, ni el ejercicio fiscal. Los únicos campos de fecha son `FechaPeticion`, `FechaEntrada`, `FechaSalida`, `FechaCreacion` y `UltimaModificacion`.
   - **Hecho acreditado en consultas**: `ExisteFacturaEnBD` ejecuta:
     `SELECT COUNT(*) FROM Ordenes WHERE Residencia = '...' AND NumFactura = ...`
     No filtra por ningún año ni fecha. Tampoco consulta `Ordenes_Historico`. En consecuencia, si una orden con factura es archivada a `Ordenes_Historico`, su número de factura deja de existir para `ExisteFacturaEnBD`, y podría volver a adjudicarse a otra orden activa en la misma sede.

---

## 5. EXTRACCIÓN Y ANÁLISIS DE LA RECONCILIACIÓN OFFLINE (SECCIÓN 7)

### 5.1. Verificación Previa de Identidad de Procedimientos
Se ha realizado una búsqueda exhaustiva en todos los módulos `.bas` y `.cls`. Se acredita documentalmente que el procedimiento denominado en enunciados auditoriales como `ProcesarColaSincronizacionLocal` corresponde exactamente a la función pública **`SincronizarColaOfflineHaciaAccess`** implementada en `ModuloSeguridadLocal.bas` (líneas 905-1116).

### 5.2. Procedimientos Extraídos Íntegramente

#### A. `ObtenerSiguienteNumOrdenProvisional` y `EsNumOrdenProvisional` (`ModuloSeguridadLocal.bas`, líneas 598-619)
```vba
 598: Public Function ObtenerSiguienteNumOrdenProvisional(ByVal wsResidencia As Worksheet) As Long
 599:     Dim maxProv As Long
 600:     maxProv = MIN_ORDEN_PROVISIONAL - 1
 601:     
 602:     If Not wsResidencia Is Nothing Then
 603:         Dim ultFila As Long, r As Long, v As Variant
 604:         ultFila = wsResidencia.Cells(wsResidencia.Rows.Count, 1).End(xlUp).Row
 605:         For r = 2 To ultFila
 606:             v = wsResidencia.Cells(r, 1).Value
 607:             If IsNumeric(v) And Not IsEmpty(v) Then
 608:                 If CLng(v) >= MIN_ORDEN_PROVISIONAL And CLng(v) > maxProv Then
 609:                     maxProv = CLng(v)
 610:                 End If
 611:             End If
 612:         Next r
 613:     End If
 614:     ObtenerSiguienteNumOrdenProvisional = maxProv + 1
 615: End Function
 616: 
 617: Public Function EsNumOrdenProvisional(ByVal numOrden As Long) As Boolean
 618:     EsNumOrdenProvisional = (numOrden >= MIN_ORDEN_PROVISIONAL)
 619: End Function

```

#### B. `ResolverNumOrdenProvisionalEnBD` (`ModuloSeguridadLocal.bas`, líneas 1364-1402)
```vba
1364: Private Function ResolverNumOrdenProvisionalEnBD( _
1365:     ByVal resCode As String, _
1366:     ByVal numOrdenProv As Long, _
1367:     ByVal wsCola As Worksheet, _
1368:     ByVal filaActualCola As Long, _
1369:     ByVal ultFilaCola As Long _
1370: ) As Long
1371:     On Error GoTo ErrResProv
1372:     Dim sqlMax As String, vMax As Variant, nuevoNum As Long
1373:     sqlMax = "SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '" & modDatabase.EscaparSQL(resCode) & "' AND NumOrden < " & MIN_ORDEN_PROVISIONAL
1374:     vMax = modDatabase.ExecuteScalar(sqlMax)
1375:     If IsNumeric(vMax) And Not IsNull(vMax) Then
1376:         nuevoNum = CLng(vMax) + 1
1377:     Else
1378:         nuevoNum = 1
1379:     End If
1380:     
1381:     If modDatabase.AsegurarOrdenExiste(nuevoNum, resCode) Then
1382:         Dim wsSede As Worksheet
1383:         Set wsSede = modDatabase.ObtenerHojaSegura(modDatabase.ObtenerNombreHoja(resCode))
1384:         If Not wsSede Is Nothing Then
1385:             Dim fHoja As Long
1386:             fHoja = BuscarFilaPorNumOrden(wsSede, numOrdenProv)
1387:             If fHoja > 0 Then wsSede.Cells(fHoja, 1).Value = nuevoNum
1388:         End If
1389:         
1390:         Dim k As Long
1391:         For k = filaActualCola To ultFilaCola
1392:             If UCase$(Trim$(CStr(wsCola.Cells(k, eColCola.colResidencia).Value))) = resCode And _
1393:                CLng(val(wsCola.Cells(k, eColCola.colNumOrden).Value)) = numOrdenProv Then
1394:                 wsCola.Cells(k, eColCola.colNumOrden).Value = nuevoNum
1395:             End If
1396:         Next k
1397:         ResolverNumOrdenProvisionalEnBD = nuevoNum
1398:     End If
1399:     Exit Function
1400: ErrResProv:
1401:     ResolverNumOrdenProvisionalEnBD = 0
1402: End Function

```

#### C. `SincronizarColaOfflineHaciaAccess` (`ModuloSeguridadLocal.bas`, líneas 905-1116)
```vba
 905: Public Function SincronizarColaOfflineHaciaAccess() As Boolean
 906:     On Error GoTo errSync
 907:     If m_SincronizandoEnCurso Then Exit Function
 908:     
 909:     Dim evOrig As Boolean, scrOrig As Boolean
 910:     evOrig = Application.enableEvents
 911:     scrOrig = Application.screenUpdating
 912:     
 913:     m_SincronizandoEnCurso = True
 914:     
 915:     ' Verificar conexión real con Access antes de iniciar el vaciado
 916:     If Not modDatabase.ComprobarConexion() Then
 917:         m_SincronizandoEnCurso = False
 918:         m_EnModoContingencia = True
 919:         m_UltimoFalloTick = Timer
 920:         SincronizarColaOfflineHaciaAccess = False
 921:         Exit Function
 922:     End If
 923:     
 924:     Dim wsCola As Worksheet
 925:     Set wsCola = ObtenerOCrearHojaCola()
 926:     Call RehidratarColaDesdeDiscoWALSiNecesario(wsCola)
 927:     
 928:     Dim ultFila As Long
 929:     ultFila = wsCola.Cells(wsCola.Rows.Count, eColCola.colIDTransaccion).End(xlUp).Row
 930:     If ultFila < 2 Then
 931:         m_SincronizandoEnCurso = False
 932:         m_EnModoContingencia = False
 933:         modDatabase.ActualizarBarraEstado
 934:         SincronizarColaOfflineHaciaAccess = True
 935:         Exit Function
 936:     End If
 937:     
 938:     m_EnModoContingencia = False
 939:     Application.enableEvents = False
 940:     Application.screenUpdating = False
 941:     Application.StatusBar = "SINCRONIZANDO COLA FIFO CON SERVIDOR CENTRAL..."
 942:     
 943:     DesprotegerHojaCola wsCola
 944:     
 945:     If ultFila > 2 Then
 946:         wsCola.Range(wsCola.Cells(1, 1), wsCola.Cells(ultFila, eColCola.colCRC)).Sort _
 947:             Key1:=wsCola.Cells(2, eColCola.colIDTransaccion), Order1:=xlAscending, Header:=xlYes
 948:     End If
 949:     
 950:     Dim r As Long
 951:     Dim estadoItem As String
 952:     Dim resItem As String, usrItem As String, tipoOp As String, campoMod As String
 953:     Dim valAnt As String, valNuevo As String, fechaLocalStr As String
 954:     Dim numOrdItem As Long
 955:     Dim totalConciliados As Long, totalConflictos As Long, totalErroresRed As Long
 956:     Dim detalleConflictos As String
 957:     
 958:     totalConciliados = 0
 959:     totalConflictos = 0
 960:     totalErroresRed = 0
 961:     detalleConflictos = ""
 962:     
 963:     For r = 2 To ultFila
 964:         estadoItem = UCase$(Trim$(CStr(wsCola.Cells(r, eColCola.colEstadoSync).Value)))
 965:         
 966:         If estadoItem = "PENDIENTE" Or estadoItem = "ERROR" Then
 967:             wsCola.Cells(r, eColCola.colEstadoSync).Value = "SINCRONIZANDO"
 968:             
 969:             fechaLocalStr = Trim$(CStr(wsCola.Cells(r, eColCola.colFechaHoraLocal).Value))
 970:             usrItem = Trim$(CStr(wsCola.Cells(r, eColCola.colUsuario).Value))
 971:             resItem = Trim$(CStr(wsCola.Cells(r, eColCola.colResidencia).Value))
 972:             numOrdItem = CLng(val(wsCola.Cells(r, eColCola.colNumOrden).Value))
 973:             tipoOp = UCase$(Trim$(CStr(wsCola.Cells(r, eColCola.colTipoOperacion).Value)))
 974:             campoMod = Trim$(CStr(wsCola.Cells(r, eColCola.colCampoModificado).Value))
 975:             valAnt = CStr(wsCola.Cells(r, eColCola.colValorAnterior).Value)
 976:             valNuevo = CStr(wsCola.Cells(r, eColCola.colValorNuevo).Value)
 977:             
 978:             Dim okItem As Boolean
 979:             Dim esConflictoHab As Boolean
 980:             Dim msgConflicto As String
 981:             okItem = False
 982:             esConflictoHab = False
 983:             msgConflicto = ""
 984:             
 985:             Select Case tipoOp
 986:                 Case "LOG_OFFLINE"
 987:                     okItem = InsertarLogHistoricoFIFO(fechaLocalStr, usrItem, campoMod, valNuevo)
 988:                     
 989:                 Case "ELIMINAR_ORDEN"
 990:                     If FueExpedienteModificadoPosteriormenteEnBD(resItem, numOrdItem, fechaLocalStr, msgConflicto) Then
 991:                         esConflictoHab = True
 992:                     Else
 993:                         okItem = modDatabase.EliminarOrdenBD(resItem, numOrdItem)
 994:                     End If
 995:                     
 996:                 Case "PRERESERVA", "SYNC_FILA_COMPLETA"
 997:                     Dim eraProvisional As Boolean
 998:                     eraProvisional = EsNumOrdenProvisional(numOrdItem)
 999:                     If eraProvisional Then
1000:                         Dim nuevoOrdReal As Long
1001:                         nuevoOrdReal = ResolverNumOrdenProvisionalEnBD(resItem, numOrdItem, wsCola, r, ultFila)
1002:                         If nuevoOrdReal > 0 Then numOrdItem = nuevoOrdReal
1003:                     End If
1004:                     If Not eraProvisional And FueExpedienteModificadoPosteriormenteEnBD(resItem, numOrdItem, fechaLocalStr, msgConflicto) Then
1005:                         esConflictoHab = True
1006:                     Else
1007:                         Dim wsHojaSede As Worksheet
1008:                         Set wsHojaSede = modDatabase.ObtenerHojaSegura(modDatabase.ObtenerNombreHoja(resItem))
1009:                         Dim filaEnHoja As Long
1010:                         filaEnHoja = BuscarFilaPorNumOrden(wsHojaSede, numOrdItem)
1011:                         If filaEnHoja > 0 Then
1012:                             Dim habFila As String
1013:                             habFila = ObtenerHabitacionDeFila(wsHojaSede, filaEnHoja, resItem)
1014:                             If Len(habFila) > 0 Then
1015:                                 If ExisteConflictoHabitacionServerWins(resItem, numOrdItem, habFila, msgConflicto) Then
1016:                                     esConflictoHab = True
1017:                                     LimpiarHabitacionEnConflictoFila wsHojaSede, filaEnHoja, resItem
1018:                                 End If
1019:                             End If
1020:                             okItem = modDatabase.SincronizarFilaHaciaAccess(wsHojaSede, filaEnHoja, (tipoOp = "PRERESERVA" Or eraProvisional))
1021:                         Else
1022:                             okItem = True
1023:                         End If
1024:                     End If
1025:                     
1026:                 Case "ASIGNAR_FACTURA_DIFERIDA"
1027:                     okItem = ConsolidarFacturaDiferidaFIFO(resItem, numOrdItem, fechaLocalStr, usrItem)
1028:                     
1029:                 Case "UPDATE_CAMPO"
1030:                     If StrComp(campoMod, "HabitacionesAsignadas", vbTextCompare) = 0 And Len(Trim$(valNuevo)) > 0 Then
1031:                         If ExisteConflictoHabitacionServerWins(resItem, numOrdItem, valNuevo, msgConflicto) Then
1032:                             esConflictoHab = True
1033:                             RevertirHabitacionEnHojaExcel resItem, numOrdItem, valAnt
1034:                         End If
1035:                     End If
1036:                     If Not esConflictoHab Then
1037:                         If ExisteConflictoPreImageEnServidor(resItem, numOrdItem, campoMod, valAnt, valNuevo, msgConflicto) Then
1038:                             esConflictoHab = True
1039:                         End If
1040:                     End If
1041:                     
1042:                     If Not esConflictoHab Then
1043:                         okItem = modDatabase.ActualizarOrden(numOrdItem, campoMod, valNuevo, resItem)
1044:                         If okItem Then
1045:                             Call InsertarLogHistoricoFIFO(fechaLocalStr, usrItem, "SYNC_FIFO_" & UCase$(campoMod), _
1046:                                 "Consolidado offline [" & resItem & "] Orden " & numOrdItem & ": " & campoMod & _
1047:                                 " ('" & valAnt & "' -> '" & valNuevo & "')")
1048:                         End If
1049:                     End If
1050:             End Select
1051:             
1052:             If esConflictoHab Then
1053:                 wsCola.Cells(r, eColCola.colEstadoSync).Value = "CONFLICTO_HAB"
1054:                 totalConflictos = totalConflictos + 1
1055:                 detalleConflictos = detalleConflictos & "  - [" & resItem & "] Orden Nº " & numOrdItem & _
1056:                                     " (Hab. '" & valNuevo & "'): " & msgConflicto & vbCrLf
1057:                 Call InsertarLogHistoricoFIFO(fechaLocalStr, usrItem, "CONFLICTO_FIFO_SERVER_WINS", _
1058:                     "Prevalece servidor en [" & resItem & "] Orden " & numOrdItem & ": " & msgConflicto)
1059:             ElseIf okItem Then
1060:                 wsCola.Cells(r, eColCola.colEstadoSync).Value = "CONCILIADO"
1061:                 totalConciliados = totalConciliados + 1
1062:             Else
1063:                 wsCola.Cells(r, eColCola.colEstadoSync).Value = "ERROR"
1064:                 wsCola.Cells(r, eColCola.colReintentos).Value = CLng(val(wsCola.Cells(r, eColCola.colReintentos).Value)) + 1
1065:                 totalErroresRed = totalErroresRed + 1
1066:                 If Not modDatabase.ComprobarConexion() Then
1067:                     m_EnModoContingencia = True
1068:                     m_UltimoFalloTick = Timer
1069:                     Exit For
1070:                 End If
1071:             End If
1072:         End If
1073:     Next r
1074:     
1075:     For r = ultFila To 2 Step -1
1076:         If UCase$(Trim$(CStr(wsCola.Cells(r, eColCola.colEstadoSync).Value))) = "CONCILIADO" Then
1077:             wsCola.Rows(r).Delete
1078:         End If
1079:     Next r
1080:     
1081:     ProtegerHojaCola wsCola
1082:     Call VolcarColaHaciaDiscoWAL(wsCola)
1083:     
1084:     Application.enableEvents = evOrig
1085:     Application.screenUpdating = scrOrig
1086:     m_SincronizandoEnCurso = False
1087:     modDatabase.ActualizarBarraEstado
1088:     
1089:     If totalConciliados > 0 Or totalConflictos > 0 Then
1090:         Dim msgResumen As String
1091:         msgResumen = "Sincronización FIFO completada con el servidor central:" & vbCrLf & vbCrLf & _
1092:                      "  - Operaciones consolidadas en Access: " & totalConciliados & vbCrLf
1093:         If totalConflictos > 0 Then
1094:             msgResumen = msgResumen & vbCrLf & _
1095:                          "ATENCIÓN - CONFLICTOS DE HABITACIÓN (Regla Server Wins):" & vbCrLf & _
1096:                          "Las siguientes habitaciones asignadas fuera de línea ya habían sido ocupadas en el servidor por otro puesto. " & _
1097:                          "Prevalece la reserva del servidor; por favor, reasigne habitación a:" & vbCrLf & _
1098:                          detalleConflictos
1099:             MsgBox msgResumen, vbExclamation, "Reconciliación SICO-N2 con Conflictos"
1100:         Else
1101:             MsgBox msgResumen, vbInformation, "Reconciliación SICO-N2 Completada"
1102:         End If
1103:     End If
1104:     
1105:     SincronizarColaOfflineHaciaAccess = (totalErroresRed = 0)
1106:     Exit Function
1107: 
1108: errSync:
1109:     On Error Resume Next
1110:     ProtegerHojaCola wsCola
1111:     Application.enableEvents = evOrig
1112:     Application.screenUpdating = scrOrig
1113:     m_SincronizandoEnCurso = False
1114:     modDatabase.ActualizarBarraEstado
1115:     SincronizarColaOfflineHaciaAccess = False
1116: End Function

```

#### D. `ConsolidarFacturaDiferidaFIFO` (`ModuloSeguridadLocal.bas`, líneas 1122-1187)
```vba
1122: Private Function ConsolidarFacturaDiferidaFIFO( _
1123:     ByVal resCode As String, _
1124:     ByVal numOrden As Long, _
1125:     ByVal fechaHoraLocal As String, _
1126:     ByVal usuario As String _
1127: ) As Boolean
1128:     On Error GoTo ErrFactFIFO
1129:     
1130:     Dim anioEjercicio As Integer
1131:     If IsDate(fechaHoraLocal) Then
1132:         anioEjercicio = Year(CDate(fechaHoraLocal))
1133:     Else
1134:         anioEjercicio = Year(Now)
1135:     End If
1136:     
1137:     Dim numFactOficial As Long
1138:     numFactOficial = modDatabase.ObtenerSiguienteNumFactura(resCode, anioEjercicio)
1139:     If numFactOficial <= 0 Then
1140:         ConsolidarFacturaDiferidaFIFO = False
1141:         Exit Function
1142:     End If
1143:     
1144:     If Not modDatabase.ActualizarOrden(numOrden, "NumFactura", CStr(numFactOficial), resCode) Then
1145:         ConsolidarFacturaDiferidaFIFO = False
1146:         Exit Function
1147:     End If
1148:     
1149:     Dim nombreHojaRes As String, nombreHojaSum As String
1150:     Dim colPagRes As String, colPagSum As String, colFactSum As String
1151:     
1152:     Select Case UCase$(Trim$(resCode))
1153:         Case "GIJON"
1154:             nombreHojaRes = "RESIDENCIA GIJÓN"
1155:             nombreHojaSum = "RESUMEN GIJÓN"
1156:             colPagRes = "AA": colPagSum = "M": colFactSum = "N"
1157:         Case "SOTO"
1158:             nombreHojaRes = "RESIDENCIA SOTO"
1159:             nombreHojaSum = "RESUMEN SOTO"
1160:             colPagRes = "AA": colPagSum = "L": colFactSum = "M"
1161:         Case "OVIEDO"
1162:             nombreHojaRes = "RESIDENCIA OVIEDO"
1163:             nombreHojaSum = "RESUMEN OVIEDO"
1164:             colPagRes = "AB": colPagSum = "M": colFactSum = "N"
1165:     End Select
1166:     
1167:     Dim wsRes As Worksheet
1168:     Set wsRes = modDatabase.ObtenerHojaSegura(nombreHojaRes)
1169:     If Not wsRes Is Nothing Then
1170:         Dim filaOrd As Long
1171:         filaOrd = BuscarFilaPorNumOrden(wsRes, numOrden)
1172:         If filaOrd > 0 Then
1173:             wsRes.Cells(filaOrd, "C").Value = numFactOficial
1174:             Call MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen( _
1175:                 wsRes.Name, "C", colPagRes, nombreHojaSum, colFactSum, colPagSum, numOrden)
1176:         End If
1177:     End If
1178:     
1179:     Call InsertarLogHistoricoFIFO(fechaHoraLocal, usuario, "FACTURA_DIFERIDA_FIFO", _
1180:         "Convertida PROFORMA a Factura Oficial Nº " & numFactOficial & " en [" & resCode & "] para Nº ORDEN " & numOrden)
1181:         
1182:     ConsolidarFacturaDiferidaFIFO = True
1183:     Exit Function
1184: 
1185: ErrFactFIFO:
1186:     ConsolidarFacturaDiferidaFIFO = False
1187: End Function

```

#### E. Estructura de la Cola WAL y Coalescing (`ModuloSeguridadLocal.bas`, líneas 38-52, 330-388)
```vba
  37: ' Columnas de SYS_COLA_SYNC (Especificación PROP-03 + Integridad)
  38: Private Enum eColCola
  39:     colIDTransaccion = 1     ' A: ID_Transaccion (Long secuencial FIFO)
  40:     colFechaHoraLocal = 2    ' B: FechaHoraLocal (yyyy-mm-dd hh:nn:ss)
  41:     colUsuario = 3           ' C: Usuario
  42:     colResidencia = 4        ' D: Residencia (GIJON, SOTO, OVIEDO)
  43:     colNumOrden = 5          ' E: NumOrden (Definitivo o Provisional 900xxx)
  44:     colTipoOperacion = 6     ' F: TipoOperacion (UPDATE_CAMPO, ASIGNAR_FACTURA_DIFERIDA, PRERESERVA, ELIMINAR_ORDEN, LOG_OFFLINE)
  45:     colCampoModificado = 7   ' G: CampoModificado (EstadoPago, HabitacionesAsignadas, etc.)
  46:     colValorAnterior = 8     ' H: ValorAnterior (Pre-Image para auditoría/rollback)
  47:     colValorNuevo = 9        ' I: ValorNuevo
  48:     colEstadoSync = 10       ' J: EstadoSync (PENDIENTE, SINCRONIZANDO, CONCILIADO, CONFLICTO_HAB, ERROR)
  49:     colReintentos = 11       ' K: Reintentos
  50:     colUUID = 12             ' L: UUID de Idempotencia
  51:     colCRC = 13              ' M: CRC16 de Integridad
  52: End Enum

' ...
 333:     ' COMPACTACIÓN INTELIGENTE (COALESCING)
 334:     Dim r As Long
 335:     Dim filaDestino As Long
 336:     Dim idTrans As Long
 337:     Dim uuidOp As String
 338:     Dim fechaStr As String
 339:     fechaStr = Format$(Now, "yyyy-mm-dd hh:nn:ss")
 340:     filaDestino = 0
 341:     
 342:     If UCase$(tipoOperacion) = "UPDATE_CAMPO" Or UCase$(tipoOperacion) = "ASIGNAR_FACTURA_DIFERIDA" Then
 343:         For r = 2 To ultFila
 344:             If UCase$(Trim$(CStr(wsCola.Cells(r, eColCola.colEstadoSync).Value))) = "PENDIENTE" Then
 345:                 If UCase$(Trim$(CStr(wsCola.Cells(r, eColCola.colResidencia).Value))) = resNorm And _
 346:                    CLng(val(wsCola.Cells(r, eColCola.colNumOrden).Value)) = numOrden And _
 347:                    StrComp(Trim$(CStr(wsCola.Cells(r, eColCola.colCampoModificado).Value)), Trim$(campoModificado), vbTextCompare) = 0 And _
 348:                    StrComp(Trim$(CStr(wsCola.Cells(r, eColCola.colTipoOperacion).Value)), Trim$(tipoOperacion), vbTextCompare) = 0 Then
 349:                     filaDestino = r
 350:                     idTrans = CLng(val(wsCola.Cells(r, eColCola.colIDTransaccion).Value))
 351:                     uuidOp = CStr(wsCola.Cells(r, eColCola.colUUID).Value)
 352:                     valorAnterior = CStr(wsCola.Cells(r, eColCola.colValorAnterior).Value)
 353:                     Exit For
 354:                 End If
 355:             End If
 356:         Next r
 357:     End If
 358:     
 359:     If filaDestino = 0 Then
 360:         filaDestino = ultFila + 1
 361:         If ultFila < 2 Then
 362:             idTrans = 1
 363:         Else
 364:             idTrans = CLng(val(wsCola.Cells(ultFila, eColCola.colIDTransaccion).Value)) + 1
 365:         End If
 366:         uuidOp = GenerarUUID(resNorm, idTrans)
 367:         
 368:         wsCola.Cells(filaDestino, eColCola.colIDTransaccion).Value = idTrans
 369:         wsCola.Cells(filaDestino, eColCola.colFechaHoraLocal).Value = fechaStr
 370:         wsCola.Cells(filaDestino, eColCola.colUsuario).Value = usuarioAct
 371:         wsCola.Cells(filaDestino, eColCola.colResidencia).Value = resNorm
 372:         wsCola.Cells(filaDestino, eColCola.colNumOrden).Value = numOrden
 373:         wsCola.Cells(filaDestino, eColCola.colTipoOperacion).Value = tipoOperacion
 374:         wsCola.Cells(filaDestino, eColCola.colCampoModificado).Value = campoModificado
 375:         wsCola.Cells(filaDestino, eColCola.colValorAnterior).Value = "'" & valorAnterior
 376:         wsCola.Cells(filaDestino, eColCola.colEstadoSync).Value = "PENDIENTE"
 377:         wsCola.Cells(filaDestino, eColCola.colReintentos).Value = 0
 378:         wsCola.Cells(filaDestino, eColCola.colUUID).Value = uuidOp
 379:     End If
 380:     
 381:     wsCola.Cells(filaDestino, eColCola.colValorNuevo).Value = "'" & valorNuevo
 382:     wsCola.Cells(filaDestino, eColCola.colFechaHoraLocal).Value = fechaStr
 383:     
 384:     Dim rawCRC As String
 385:     rawCRC = idTrans & "|" & uuidOp & "|" & resNorm & "|" & numOrden & "|" & tipoOperacion & "|" & campoModificado & "|" & valorNuevo
 386:     wsCola.Cells(filaDestino, eColCola.colCRC).Value = CalcularCRC16(rawCRC)
 387:     
 388:     ProtegerHojaCola wsCola

```

### 5.3. Reconstrucción Documental del Ciclo de Reconciliación Offline SICO-N2

1. **Generación del Número Provisional**:
   - Se genera mediante `ObtenerSiguienteNumOrdenProvisional(wsResidencia)`.
   - Escanea la columna A de la hoja de la residencia activa, localizando el valor numérico máximo que sea $\ge 900001$ (`MIN_ORDEN_PROVISIONAL`). Si no hay ninguno, asigna $900001$; si hay otros, asigna `maxProv + 1`.
2. **Relación con la Solicitud Original**:
   - La fila física de Excel conserva el número provisional ($900xxx$) en la columna A.
   - En la cola oculta `SYS_COLA_SYNC` y en el archivo en disco (`sico_wal_queue.dat`), todas las operaciones encoladas para esa fila registran en `colNumOrden` el valor provisional $900xxx$, asociadas al código de residencia (`colResidencia`), el usuario local y un UUID único (`colUUID`).
3. **Creación de la Reserva Definitiva en Access**:
   - Al restablecerse la red, `SincronizarColaOfflineHaciaAccess` procesa la cola en orden FIFO (`colIDTransaccion ASC`).
   - Al encontrar una operación de tipo `PRERESERVA` o `SYNC_FILA_COMPLETA` con un número provisional, invoca a `ResolverNumOrdenProvisionalEnBD(resItem, numOrdItem, wsCola, r, ultFila)`.
   - `ResolverNumOrdenProvisionalEnBD` consulta en Access: `SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '...' AND NumOrden < 900001`.
   - Toma `nuevoNum = MAX + 1` y crea la reserva en Access ejecutando `modDatabase.AsegurarOrdenExiste(nuevoNum, resCode)`.
4. **Actualización de la Hoja y de la Cola**:
   - Inmediatamente después de asegurar la orden en Access, `ResolverNumOrdenProvisionalEnBD` busca la fila en la hoja de Excel mediante `BuscarFilaPorNumOrden` y reemplaza el $900xxx$ por `nuevoNum` en la celda `(fHoja, 1)`.
   - A continuación, recorre el resto de la cola desde la fila actual hasta el final (`For k = filaActualCola To ultFilaCola`) y actualiza `colNumOrden` en todas las operaciones pendientes de esa misma orden provisional, sustituyendo $900xxx$ por `nuevoNum`.
   - Después, `SincronizarColaOfflineHaciaAccess` ejecuta `SincronizarFilaHaciaAccess` con el número definitivo recién asignado.
5. **Tratamiento de Errores y Reintentos**:
   - Si una operación falla por error de red: se incrementa `colReintentos` (`CLng(val) + 1`), se marca el estado como `"ERROR"`, se activa de nuevo el modo contingencia (`m_EnModoContingencia = True`), se anota `m_UltimoFalloTick = Timer` para cumplir el cooldown de 60 segundos, y se interrumpe el bucle FIFO (`Exit For`) para preservar el orden secuencial.
   - Las operaciones exitosas se marcan como `"CONCILIADO"`.
   - Al final de la pasada, todas las filas con estado `"CONCILIADO"` se eliminan físicamente de `SYS_COLA_SYNC` y se guarda la cola depurada en disco WAL (`VolcarColaHaciaDiscoWAL`).
6. **Prevención de Procesamiento Duplicado (Idempotencia y Coalescing)**:
   - En el momento de encolar (`EncolarOperacionOffline`), si la operación es de tipo `UPDATE_CAMPO` o `ASIGNAR_FACTURA_DIFERIDA`, el sistema realiza **compactación inteligente (*coalescing*)**: busca si ya existe una operación previa en estado `PENDIENTE` para el mismo `(Residencia, NumOrden, CampoModificado, TipoOperacion)`. Si existe, reutiliza esa fila, conserva el `ValorAnterior` original y actualiza únicamente el `ValorNuevo` y el CRC16, evitando que múltiples ediciones sobre el mismo campo generen duplicados en cola.
   - Cada entrada de la cola posee un `UUID` único (`colUUID`) y un `CRC16` (`colCRC`).
   - La eliminación de filas `CONCILIADO` garantiza que una operación confirmada por Access nunca se vuelva a ejecutar en una reconexión posterior.
7. **Sincronización Offline de Cambios en Nº FACTURA**:
   - **Factura automática en offline**: `AsignarFacturaPorColumna` graba el texto `"PROFORMA"` en la celda C y encola una operación de tipo `ASIGNAR_FACTURA_DIFERIDA`. Durante la reconciliación FIFO, `ConsolidarFacturaDiferidaFIFO` solicita a Access el siguiente número oficial mediante `ObtenerSiguienteNumFactura(resCode, anioEjercicio)`, actualiza la base de datos central (`ActualizarOrden`), actualiza la celda C de la hoja con el número oficial obtenido y sincroniza con el RESUMEN.
   - **Edición manual de factura en offline**: `ProcesarCambioManualFactura` llama a `modDatabase.ActualizarOrden(numOrden, "NumFactura", CStr(numFacturaLong), resCode)`. Si no hay red, `ActualizarOrden` detecta la caída y llama a `ModuloSeguridadLocal.EncolarOperacionOffline(resTarget, numOrden, "UPDATE_CAMPO", "NumFactura", "", valor)`. Durante la reconciliación FIFO, esta operación se vuelca a Access como `UPDATE_CAMPO`.

---

## 6. ESQUEMA COMPLETO DE ACCESS (SECCIÓN 8)

Mediante inspección DAO de solo lectura sobre `H:\ResidenciaBD\Residencia_BE.accdb`, se han extraído todos los campos e índices de `Ordenes` y `Ordenes_Historico`.

### 6.1. Definición Física de Campos en `Ordenes` y `Ordenes_Historico`

| Pos | Nombre Campo | Tipo DAO | Código | Tamaño | Requerido (Ord) | Requerido (Hist) | AllowZeroLen | Autonumérico | Atributos DAO |
| :---: | :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| 0 | `Id` | `dbLong` | 4 | 4 | False | False | False | True | 17 |
| 1 | `NumOrden` | `dbLong` | 4 | 4 | True | False | False | False | 1 |
| 2 | `NumFactura` | `dbLong` | 4 | 4 | False | False | False | False | 1 |
| 3 | `Residencia` | `dbText` | 10 | 20 | True | False | False | False | 2 |
| 4 | `FechaPeticion` | `dbDate` | 8 | 8 | False | False | False | False | 1 |
| 5 | `DNI` | `dbText` | 10 | 20 | False | False | False | False | 2 |
| 6 | `Nombre` | `dbText` | 10 | 100 | False | False | False | False | 2 |
| 7 | `Apellidos` | `dbText` | 10 | 100 | False | False | False | False | 2 |
| 8 | `TipoHuesped` | `dbText` | 10 | 30 | False | False | False | False | 2 |
| 9 | `NumHabIndividuales` | `dbLong` | 4 | 4 | False | False | False | False | 1 |
| 10 | `NumHabDobles` | `dbLong` | 4 | 4 | False | False | False | False | 1 |
| 11 | `FechaEntrada` | `dbDate` | 8 | 8 | False | False | False | False | 1 |
| 12 | `FechaSalida` | `dbDate` | 8 | 8 | False | False | False | False | 1 |
| 13 | `Resolucion` | `dbText` | 10 | 30 | False | False | False | False | 2 |
| 14 | `HabitacionesAsignadas` | `dbText` | 10 | 200 | False | False | False | False | 2 |
| 15 | `EstadoPago` | `dbText` | 10 | 30 | False | False | False | False | 2 |
| 16 | `Solapamiento` | `dbBoolean` | 1 | 1 | False | False | False | False | 1 |
| 17 | `Observaciones` | `dbMemo` | 12 | 0 | False | False | False | False | 2 |
| 18 | `Telefono` | `dbText` | 10 | 20 | False | False | False | False | 2 |
| 19 | `Email` | `dbText` | 10 | 100 | False | False | False | False | 2 |
| 20 | `Direccion` | `dbText` | 10 | 200 | False | False | False | False | 2 |
| 21 | `CodigoPostal` | `dbText` | 10 | 10 | False | False | False | False | 2 |
| 22 | `Poblacion` | `dbText` | 10 | 100 | False | False | False | False | 2 |
| 23 | `Provincia` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 24 | `ConsentimientoRGPD` | `dbBoolean` | 1 | 1 | False | False | False | False | 1 |
| 25 | `PagadoResidencia` | `dbText` | 10 | 30 | False | False | False | False | 2 |
| 26 | `FechaCreacion` | `dbDate` | 8 | 8 | False | False | False | False | 1 |
| 27 | `UsuarioCreacion` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 28 | `Finalidad` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 29 | `Empleo` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 30 | `Situacion` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 31 | `Evaluacion` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 32 | `Comision` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 33 | `Turno` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 34 | `Rango` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 35 | `DiasUso` | `dbLong` | 4 | 4 | False | False | False | False | 1 |
| 36 | `PAX` | `dbLong` | 4 | 4 | False | False | False | False | 1 |
| 37 | `CamaSuple` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 38 | `DtoFamNum` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 39 | `Importe` | `dbDouble` | 7 | 8 | False | False | False | False | 1 |
| 40 | `FechaGrabacion` | `dbText` | 10 | 255 | False | False | False | False | 2 |
| 41 | `UsuarioModificacion` | `dbText` | 10 | 50 | False | False | False | False | 2 |
| 42 | `UltimaModificacion` | `dbDate` | 8 | 8 | False | False | False | False | 1 |

### 6.2. Índices de las Tablas `Ordenes` y `Ordenes_Historico`

#### A. Índices en `Ordenes` (Total: 7 índices)

| Nombre del Índice | Campos Integrantes | Unique | Primary | Required | IgnoreNulls |
| :--- | :--- | :---: | :---: | :---: | :---: |
| `idx_Ordenes_DNI` | `DNI` | False | False | False | False |
| `idx_Ordenes_FechaEntrada` | `FechaEntrada` | False | False | False | False |
| `idx_Ordenes_FechaSalida` | `FechaSalida` | False | False | False | False |
| `idx_Ordenes_Residencia` | `Residencia` | False | False | False | False |
| `idx_Ordenes_Resolucion` | `Resolucion` | False | False | False | False |
| `idx_Residencia_NumOrden` | `Residencia, NumOrden` | True | False | False | False |
| `Index_50734CA2_0946_4D32` | `Id` | True | True | True | False |

#### B. Índices en `Ordenes_Historico` (Total: 1 índice)

| Nombre del Índice | Campos Integrantes | Unique | Primary | Required | IgnoreNulls |
| :--- | :--- | :---: | :---: | :---: | :---: |
| `idx_Hist_Res_NumOrden` | `Residencia, NumOrden` | False | False | False | False |

### 6.3. Verificaciones Estructurales Expresas (Puntos 1 a 7)

1. **Propiedades del campo `Id` en ambas tablas**:
   - En `Ordenes`: Tipo `dbLong` (4 bytes), Autonumérico (`dbAutoIncrField = 16`, atributos = 17). Posee clave primaria única.
   - En `Ordenes_Historico`: Tipo `dbLong` (4 bytes), Autonumérico (`dbAutoIncrField = 16`, atributos = 17). **NO es clave primaria ni posee índice único**.
2. **Existencia de Clave Primaria**:
   - En `Ordenes`: **SÍ existe clave primaria** sobre el campo `Id` (`Index_50734CA2_0946_4D32`, `Primary = True`, `Unique = True`). Además posee índice único de negocio sobre `(Residencia, NumOrden)` (`idx_Residencia_NumOrden`, `Unique = True`).
   - En `Ordenes_Historico`: **NO EXISTE CLAVE PRIMARIA** (`Primary = False` en toda la tabla). Su único índice es `idx_Hist_Res_NumOrden` sobre `(Residencia, NumOrden)`, el cual es **no único (`Unique = False`)**.
3. **Compatibilidad Estructural de los Campos entre Ambas Tablas**:
   - Ambas tablas contienen exactamente 43 campos.
   - Los 43 campos tienen el mismo nombre, el mismo tipo DAO y el mismo tamaño en ambas tablas.
   - La única discrepancia en atributos es la propiedad `Required`: en `Ordenes`, `NumOrden` y `Residencia` tienen `Required = True`; en `Ordenes_Historico`, tienen `Required = False`.
4. **Correspondencia del Orden Físico de los Campos**:
   - La posición ordinal (`OrdinalPosition`) del 0 al 42 coincide al 100% de manera idéntica entre `Ordenes` y `Ordenes_Historico`.
5. **Compatibilidad Efectiva de `INSERT INTO Ordenes_Historico SELECT * FROM Ordenes`**:
   - **Sintácticamente compatible**: Al tener exactamente el mismo número de campos (43), en el mismo orden físico y con los mismos tipos de datos, la instrucción ejecuta sin error de sintaxis en el motor Access Jet/ACE.
   - **Comportamiento en `Id`**: Dado que `Id` es el campo 0 en ambas tablas, el motor Access copia el valor numérico explícito de `Id` de `Ordenes` hacia `Ordenes_Historico`. Como `Ordenes_Historico` no tiene restricción de clave primaria ni índice único sobre `Id`, la inserción se realiza con éxito.
   - **Impacto Destructivo Demostrado**: En todos los procedimientos actuales (`EliminarOrdenBD`, `EliminarOrdenesLoteBD`, `EliminarTodasOrdenesSedeBD`), la sentencia `INSERT INTO Ordenes_Historico SELECT *` está **precedida por un `DELETE FROM Ordenes_Historico WHERE Residencia = '...' AND NumOrden = ...`**. Por tanto, si un número de orden ya había sido archivado en el pasado (de otro huésped o de otro año), el archivado actual **elimina físicamente el expediente histórico previo** antes de insertar el nuevo.
6. **Existencia de Campos para Identificar el Año del Expediente y el Año de Facturación**:
   - **Año del Expediente (`NumOrden`)**: **NO EXISTE**. En ninguna de las dos tablas existe campo `Año`, `Ano`, `Ejercicio` o similar para discriminar el ejercicio al que pertenece el número de orden.
   - **Año de Facturación (`NumFactura`)**: **NO EXISTE**. En ninguna de las dos tablas existe campo `AñoFactura`, `AnoFactura`, `EjercicioFactura` o `FechaFactura`. El único lugar donde existe el campo `Ejercicio` es en la tabla independiente `ContadorFacturas` (que solo guarda el último número entero asignado).
7. **Longitud y Tipo de los Campos Clave**:
   - `NumOrden`: Tipo DAO `dbLong` (4 bytes, Entero Largo). Rango: $-2.147.483.648$ a $2.147.483.647$.
   - `NumFactura`: Tipo DAO `dbLong` (4 bytes, Entero Largo). Rango: $-2.147.483.648$ a $2.147.483.647$.
   - `Residencia`: Tipo DAO `dbText` (Texto corto), Tamaño = 20 caracteres.

---

## 7. MAPA DE DEPENDENCIAS Y CATÁLOGO DE PROCEDIMIENTOS (SECCIÓN 9)

A continuación se relacionan todos los procedimientos inspeccionados que intervienen en el ciclo de vida del Nº ORDEN, Nº FACTURA, archivado, eliminación y reconciliación:

| Componente / Módulo | Procedimiento / Firma | Líneas | Llamadores Directos | Dependencias Directas |
| :--- | :--- | :---: | :--- | :--- |
| `modDatabase.bas` | `Public Function EliminarOrdenBD(ByVal residencia As String, ByVal numOrden As Long) As Boolean` | 4261-4420 | - `BorradoNumOrden.GestionarBorradoOrden` (l. 58)<br>- `ModuloSeguridadLocal.SincronizarColaOfflineHaciaAccess` (l. 993) | `SondaRedDisponible`, `AsegurarEsquemaConcurrenciaBD`, `GetConnectionString`, `ObtenerTokenSesionInstancia`, `ObtenerUsuarioActual`, `EscaparSQL`, `FormatearFechaHoraSQL`, `FormatearPropietarioBloqueo`, `SincronizarHojaDesdeAccess` |
| `modDatabase.bas` | `Public Function EliminarOrdenesLoteBD(ByVal residencia As String, ByVal colOrdenes As Collection, Optional ByRef outTotalEliminadas As Long = 0) As Boolean` | 4424-4552 | - `BorradoSolicitudes.BorrarSolicitudesPorOrden` (l. 389)<br>- `modDatabase.ArchivarOrdenesAntiguasBD` (l. 4811) | `ObtenerClaveResidenciaDesdeHoja`, `SondaRedDisponible`, `EncolarBorradosLoteOffline`, `AsegurarEsquemaConcurrenciaBD`, `ObtenerTokenSesionInstancia`, `ObtenerUsuarioActual`, `GetConnectionString`, `EscaparSQL`, `FormatearPropietarioBloqueo` |
| `modDatabase.bas` | `Public Function EliminarTodasOrdenesSedeBD(ByVal residencia As String, Optional ByRef outTotalEliminadas As Long = 0, Optional ByVal colOrdenesOffline As Collection = Nothing) As Boolean` | 4556-4659 | - `BorradoDatosAlojamiento.BorrarDatosAlojamientoOptimizado` (l. 150) | `ObtenerClaveResidenciaDesdeHoja`, `SondaRedDisponible`, `EncolarBorradosLoteOffline`, `AsegurarEsquemaConcurrenciaBD`, `ObtenerTokenSesionInstancia`, `ObtenerUsuarioActual`, `GetConnectionString`, `EscaparSQL`, `FormatearPropietarioBloqueo` |
| `modDatabase.bas` | `Public Function ArchivarOrdenesAntiguasBD(Optional ByVal residencia As String = "", Optional ByVal diasRetencion As Long = 90) As Long` | 4784-4950 | - `ModuloArchivadoDatosBDAS.ArchivarDatosAntiguosAlAbrir` (l. 109-111)<br>- `ModuloArchivadoDatosBDAS.ArchivarDatosManual` (l. 220) | `ObtenerClaveResidenciaDesdeHoja`, `ObtenerFilasYOrdenesArchivablesSede`, `EliminarOrdenesLoteBD`, `ObtenerHojaSegura`, `ObtenerNombreHoja`, `ObtenerPasswordHojas`, `InsertarLog` |
| `modDatabase.bas` | `Public Function PurgarHistoricoRGPD(Optional ByVal anosRetencion As Long = 3) As Long` | 4973-5003 | - `BorradoFilasHojasBDAS.BorrarFilasBDAS` (l. 39) | `ContarHistoricoPurgableRGPD`, `ExecuteNonQuery`, `InsertarLog` |
| `BorradoNumOrden.bas` | `Public Sub GestionarBorradoOrden(ByVal ws As Worksheet, ByVal fila As Long, ByVal numOrdenAnterior As Variant, ByVal nombreResumen As String, ByVal ultimaColumna As String)` | 20-119 | - `Hoja2.Worksheet_Change` (GIJÓN)<br>- `Hoja20.Worksheet_Change` (SOTO)<br>- `Hoja24.Worksheet_Change` (OVIEDO) | `modDatabase.ObtenerClaveResidenciaDesdeHoja`, `modDatabase.EliminarOrdenBD`, `ModuloConfigSegura.ObtenerPasswordHojas`, `BorradoNumOrden.BorrarFilaEnResumen_Central` |
| `BorradoNumOrden.bas` | `Private Sub BorrarFilaEnResumen_Central(ByVal numOrden As Variant, ByVal nombreHojaResumen As String, ByVal pwdHojas As String)` | 127-184 | - `BorradoNumOrden.GestionarBorradoOrden` (l. 104) | `modDatabase.ObtenerHojaSegura`, Excel Range Find / EntireRow Delete |
| `BorradoSolicitudes.bas` | `Public Sub BorrarSolicitudesPorOrden()` | 175-488 | - `BorradoSolicitudes.BorradoParcialCeldas_Click` (Ribbon) | `modDatabase.ObtenerClaveResidenciaDesdeHoja`, `modDatabase.EliminarOrdenesLoteBD`, `ModuloConfigSegura.ObtenerPasswordHojas` |
| `BorradoDatosAlojamiento.bas` | `Public Sub BorrarDatosAlojamientoOptimizado(control As IRibbonControl)` | 21-244 | - Ribbon Button `btnBorrarDatosAlojamiento` | `modDatabase.ObtenerClaveResidenciaDesdeHoja`, `modDatabase.EliminarTodasOrdenesSedeBD`, `ModuloConfigSegura.ObtenerPasswordHojas` |
| `modDatabase.bas` | `Public Function IntentarReservarOrdenBD(ByVal residencia As String, ByVal numOrdenSolicitado As Long, ByRef outUsuarioPropietario As String, ByRef outNombreHuesped As String, ByRef outSiguienteLibreBD As Long, Optional ByRef outSinRed As Boolean = False) As Boolean` | 1281-1401 | - `modDatabase.GestionarEntradaNumOrdenHoja` (l. 1601) | `SondaRedDisponible`, `AsegurarEsquemaConcurrenciaBD`, `ObtenerClaveResidenciaDesdeHoja`, `ObtenerUsuarioActual`, `ObtenerTokenSesionInstancia`, `GetConnectionString`, `EscaparSQL`, `FormatearFechaHoraSQL`, `FormatearPropietarioBloqueo`, `ExisteOrdenEnBD` |
| `modDatabase.bas` | `Public Function ReservarSiguienteOrdenLibreBD(ByVal residencia As String) As Long` | 1403-1472 | - `modDatabase.GestionarEntradaNumOrdenHoja` (l. 1649) | `SondaRedDisponible`, `AsegurarEsquemaConcurrenciaBD`, `ObtenerClaveResidenciaDesdeHoja`, `ObtenerUsuarioActual`, `ObtenerTokenSesionInstancia`, `GetConnectionString`, `EscaparSQL`, `FormatearFechaHoraSQL` |
| `modDatabase.bas` | `Public Function RenumerarOrdenExistenteBD(ByVal claveResidencia As String, ByVal numOrdenPrevio As Long, ByVal numOrdenNuevo As Long) As Boolean` | 1475-1525 | - `modDatabase.GestionarEntradaNumOrdenHoja` (l. 1582) | `SondaRedDisponible`, `ObtenerTokenSesionInstancia`, `GetConnectionString`, `EscaparSQL`, `FormatearFechaHoraSQL` |
| `modDatabase.bas` | `Public Function GestionarEntradaNumOrdenHoja(ByVal ws As Worksheet, ByVal Target As Range, ByVal targetRow As Long, ByRef ioValorAnteriorHoja As Variant, ByVal claveResidencia As String, ByVal nombreHojaResumen As String, ByVal colFactSum As String, ByVal colPagoSum As String, ByVal colPagoRes As String) As Boolean` | 1527-1696 | - `Hoja2.Worksheet_Change` (l. 153)<br>- `Hoja20.Worksheet_Change` (l. 138)<br>- `Hoja24.Worksheet_Change` (l. 155) | `ExisteOrdenEnBD`, `RenumerarOrdenExistenteBD`, `MarcarSiPagadosEnResidencia.SincronizarOrdenYCampos`, `SincronizarFilaHaciaAccess`, `IntentarReservarOrdenBD`, `ModuloSeguridadLocal.EsNumOrdenProvisional`, `ModuloSeguridadLocal.ObtenerSiguienteNumOrdenProvisional`, `ModuloSeguridadLocal.EncolarOperacionOffline`, `SincronizarHojaDesdeAccess`, `ReservarSiguienteOrdenLibreBD` |
| `modDatabase.bas` | `Public Function AsegurarOrdenExiste(ByVal numOrden As Long, ByVal residencia As String) As Boolean` | 1698-1732 | - `ModuloSeguridadLocal.ResolverNumOrdenProvisionalEnBD` (l. 1381) | `ObtenerClaveResidenciaDesdeHoja`, `AsegurarEsquemaConcurrenciaBD`, `ExecuteScalar`, `ObtenerUsuarioActual`, `ObtenerTokenSesionInstancia`, `EscaparSQL`, `FormatearFechaHoraSQL`, `ExecuteNonQuery` |
| `modDatabase.bas` | `Public Function ExisteOrdenEnBD(ByVal residencia As String, ByVal numOrden As Long, Optional ByRef outUsuarioCreacion As String = "", Optional ByRef outNombreHuesped As String = "", Optional ByRef outDNI As String = "", Optional ByRef outSiguienteLibreBD As Long = 0) As Boolean` | 2812-2885 | - `modDatabase.GestionarEntradaNumOrdenHoja` (l. 1569)<br>- `modDatabase.IntentarReservarOrdenBD` (l. 1396)<br>- `DuplicadosOrdenyFactura.VerificarDuplicados` (l. 351) | `SondaRedDisponible`, `ObtenerClaveResidenciaDesdeHoja`, `GetConnectionString`, `EscaparSQL` |
| `AsignarNumFactura.bas` | `Public Sub AsignarFacturaPorColumna(ws As Worksheet, ByVal Target As Range, ByVal columnaCheck As String, Optional ByVal columnaFactura As String = "C")` | 14-232 | - `Hoja2.Worksheet_Change` (l. 250)<br>- `Hoja20.Worksheet_Change` (l. 210)<br>- `Hoja24.Worksheet_Change` (l. 235) | `modDatabase.ObtenerClaveResidenciaDesdeHoja`, `MarcarSiPagadosEnResidencia.EsEstadoPagado`, `ModuloSeguridadLocal.EstaEnModoContingencia`, `ModuloSeguridadLocal.SondaRedDisponible`, `ModuloSeguridadLocal.EncolarOperacionOffline`, `modDatabase.ExecuteScalar`, `DetectarHuecosEnSerie`, `modDatabase.ObtenerSiguienteNumFactura`, `modDatabase.ActualizarOrden`, `MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen`, `modDatabase.MostrarYActivarHojaSegura` |
| `AsignarNumFactura.bas` | `Public Sub ProcesarCambioManualFactura(ws As Worksheet, Target As Range, ByVal targetRow As Long, Optional ByVal columnaFactura As String = "C")` | 561-641 | - `Hoja2.Worksheet_Change` (l. 195)<br>- `Hoja20.Worksheet_Change` (l. 165)<br>- `Hoja24.Worksheet_Change` (l. 185) | `modDatabase.ObtenerClaveResidenciaDesdeHoja`, `modDatabase.ActualizarOrden`, `DuplicadosOrdenyFactura.VerificarDuplicados`, `modDatabase.ExisteFacturaEnBD`, `AsignarNumFactura.ValidarFacturaManual`, `modDatabase.SincronizarContadorFacturaManual` |
| `AsignarNumFactura.bas` | `Public Sub ValidarFacturaManual(ws As Worksheet, numeroEscrito As Long, columnaFactura As String)` | 649-769 | - `AsignarNumFactura.ProcesarCambioManualFactura` (l. 636) | `OrdenarArray`, `AgruparEnBloques` |
| `modDatabase.bas` | `Public Function ObtenerSiguienteNumFactura(ByVal residencia As String, ByVal ejercicio As Integer) As Long` | 2678-2778 | - `AsignarNumFactura.AsignarFacturaPorColumna` (l. 188)<br>- `ModuloSeguridadLocal.ConsolidarFacturaDiferidaFIFO` (l. 1138) | `ObtenerClaveResidenciaDesdeHoja`, `GetConnectionString`, `EscaparSQL`, ADODB Connection / BeginTrans / CommitTrans / RollbackTrans sobre `ContadorFacturas` |
| `modDatabase.bas` | `Public Sub SincronizarContadorFacturaManual(ByVal residencia As String, ByVal numFacturaManual As Long, Optional ByVal ejercicio As Integer = 0)` | 2782-2801 | - `AsignarNumFactura.ProcesarCambioManualFactura` (l. 640) | `ObtenerClaveResidenciaDesdeHoja`, `EscaparSQL`, `ExecuteNonQuery` sobre `ContadorFacturas` |
| `modDatabase.bas` | `Public Function ExisteFacturaEnBD(ByVal residencia As String, ByVal numFactura As Long, Optional ByVal excluirNumOrden As Long = 0) As Boolean` | 2889-2916 | - `AsignarNumFactura.ProcesarCambioManualFactura` (l. 626) | `ObtenerClaveResidenciaDesdeHoja`, `EscaparSQL`, `ExecuteScalar` sobre `Ordenes` |
| `modDatabase.bas` | `Public Sub RefrescarColumnaFacturasDesdeAccess(ws As Worksheet, ByVal residencia As String)` | 2920-3010 | - `Hoja2.Worksheet_Activate`<br>- `Hoja20.Worksheet_Activate`<br>- `Hoja24.Worksheet_Activate` | `SondaRedDisponible`, `ObtenerClaveResidenciaDesdeHoja`, `GetConnectionString`, ADODB Connection sobre `Ordenes` |
| `DuplicadosOrdenyFactura.bas` | `Public Sub VerificarDuplicados(ByVal Target As Range, ByVal rangoColumna As Range, ByVal nombreColumna As String)` | 249-387 | - `AsignarNumFactura.ProcesarCambioManualFactura` (l. 621)<br>- Worksheet_Change de residencias | `modDatabase.ExisteOrdenEnBD`, Excel Application.CountIf |
| `ModuloSeguridadLocal.bas` | `Public Function ObtenerSiguienteNumOrdenProvisional(ByVal wsResidencia As Worksheet) As Long` | 598-615 | - `modDatabase.GestionarEntradaNumOrdenHoja` (l. 1610) | Lectura de celda A en `wsResidencia` contra `MIN_ORDEN_PROVISIONAL` (900001) |
| `ModuloSeguridadLocal.bas` | `Public Function EsNumOrdenProvisional(ByVal numOrden As Long) As Boolean` | 617-619 | - `modDatabase.GestionarEntradaNumOrdenHoja` (l. 1607)<br>- `ModuloSeguridadLocal.SincronizarColaOfflineHaciaAccess` (l. 998) | Comparación numérica $\ge 900001$ |
| `ModuloSeguridadLocal.bas` | `Private Function ResolverNumOrdenProvisionalEnBD(ByVal resCode As String, ByVal numOrdenProv As Long, ByVal wsCola As Worksheet, ByVal filaActualCola As Long, ByVal ultFilaCola As Long) As Long` | 1364-1402 | - `ModuloSeguridadLocal.SincronizarColaOfflineHaciaAccess` (l. 1001) | `modDatabase.EscaparSQL`, `modDatabase.ExecuteScalar`, `modDatabase.AsegurarOrdenExiste`, `modDatabase.ObtenerHojaSegura`, `BuscarFilaPorNumOrden` |
| `ModuloSeguridadLocal.bas` | `Public Function SincronizarColaOfflineHaciaAccess() As Boolean` | 905-1116 | - `modDatabase.SincronizarHojaDesdeAccess` (l. 3251)<br>- `ModuloSeguridadLocal.ComprobarRetornoRedYSincronizar` (l. 125) | `ObtenerOCrearHojaCola`, `DesprotegerHojaCola`, `ProtegerHojaCola`, `InsertarLogHistoricoFIFO`, `FueExpedienteModificadoPosteriormenteEnBD`, `modDatabase.EliminarOrdenBD`, `ResolverNumOrdenProvisionalEnBD`, `BuscarFilaPorNumOrden`, `ObtenerHabitacionDeFila`, `ExisteConflictoHabitacionServerWins`, `LimpiarHabitacionEnConflictoFila`, `modDatabase.SincronizarFilaHaciaAccess`, `ConsolidarFacturaDiferidaFIFO`, `RevertirHabitacionEnHojaExcel`, `ExisteConflictoPreImageEnServidor`, `modDatabase.ActualizarOrden`, `VolcarColaHaciaDiscoWAL` |
| `ModuloSeguridadLocal.bas` | `Private Function ConsolidarFacturaDiferidaFIFO(ByVal resCode As String, ByVal numOrden As Long, ByVal fechaHoraLocal As String, ByVal usuario As String) As Boolean` | 1122-1187 | - `ModuloSeguridadLocal.SincronizarColaOfflineHaciaAccess` (l. 1027) | `modDatabase.ObtenerSiguienteNumFactura`, `modDatabase.ActualizarOrden`, `modDatabase.ObtenerHojaSegura`, `BuscarFilaPorNumOrden`, `MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen`, `InsertarLogHistoricoFIFO` |
| `ModuloSeguridadLocal.bas` | `Public Function EncolarOperacionOffline(ByVal residencia As String, ByVal numOrden As Long, ByVal tipoOperacion As String, ByVal campoModificado As String, ByVal valorAnterior As String, ByVal valorNuevo As String) As Boolean` | 307-399 | - `modDatabase.GestionarEntradaNumOrdenHoja` (l. 1614)<br>- `modDatabase.SincronizarFilaHaciaAccess` (l. 1768)<br>- `modDatabase.ActualizarOrden` (l. 2545)<br>- `AsignarNumFactura.AsignarFacturaPorColumna` (l. 125, 127) | `ObtenerOCrearHojaCola`, `DesprotegerHojaCola`, `ProtegerHojaCola`, `modDatabase.ObtenerClaveResidenciaDesdeHoja`, `GenerarUUID`, `CalcularCRC16`, `VolcarColaHaciaDiscoWAL` |
| `ModuloSeguridadLocal.bas` | `Public Function EncolarBorradosLoteOffline(ByVal resCode As String, ByVal colOrdenes As Collection) As Boolean` | 440-566 | - `modDatabase.EliminarOrdenesLoteBD` (l. 4450)<br>- `modDatabase.EliminarTodasOrdenesSedeBD` (l. 4576) | `ObtenerOCrearHojaCola`, `DesprotegerHojaCola`, `ProtegerHojaCola`, `GenerarUUID`, `CalcularCRC16`, `VolcarColaHaciaDiscoWAL` |
| `ModuloSeguridadLocal.bas` | `Private Sub VolcarColaHaciaDiscoWAL(ByVal wsCola As Worksheet)` | 1460-1497 | - `ModuloSeguridadLocal.EncolarOperacionOffline` (l. 391)<br>- `ModuloSeguridadLocal.SincronizarColaOfflineHaciaAccess` (l. 1082) | `ObtenerRutaArchivoWAL` (`%LOCALAPPDATA%\ResidenciaApp\sico_wal_queue.dat`), Visual Basic File I/O (`Open For Output / Print #`) |
| `ModuloSeguridadLocal.bas` | `Public Sub RehidratarColaDesdeDiscoWALSiNecesario(ByVal wsCola As Worksheet)` | 1499-1569 | - `ModuloSeguridadLocal.ObtenerOCrearHojaCola` (l. 299) | `ObtenerRutaArchivoWAL`, Visual Basic File I/O (`Open For Input / Line Input #`), `RestaurarPipe`, `CalcularCRC16` |

---

## 8. COMPARACIÓN: REGLAS FUNCIONALES VS MECANISMOS ACTUALMENTE IMPLEMENTADOS (SECCIÓN 10)

Sin emitir juicios de valor ni proponer soluciones, se contrastan objetivamente las reglas funcionales fijadas por la Dirección Arquitectónica con el comportamiento real del código extraído:

| Regla Funcional Decidida | Mecanismo Actualmente Implementado en el Código | Hecho Objetivo / Diferencia Identificada |
| :--- | :--- | :--- |
| **1. Unicidad del Nº ORDEN**:<br>La combinación `(Residencia, Año, NumOrden)` debe ser única entre activos e históricos. El año forma parte de la identidad de numeración. | En `Ordenes`, el índice es `(Residencia, NumOrden)` con `Unique = True`. En `Ordenes_Historico`, el índice es `(Residencia, NumOrden)` con `Unique = False`.<br>No existe campo de año ni en `Ordenes` ni en `Ordenes_Historico`.<br>`ExisteOrdenEnBD`, `IntentarReservarOrdenBD` y `ReservarSiguienteOrdenLibreBD` consultan exclusivamente `Ordenes`. | **Diferencia crítica**:<br>1. El sistema no almacena ni evalúa el año del expediente.<br>2. El sistema no comprueba `Ordenes_Historico` al reservar o validar un Nº ORDEN.<br>3. Un número archivado en el histórico queda disponible para ser reutilizado de inmediato en la tabla activa. |
| **2. Conservación del Histórico**:<br>Si se archiva un expediente, su número sigue ocupado durante ese ejercicio. Si se elimina definitivamente sin histórico, puede reutilizarse. No se puede sustituir ni eliminar un expediente anterior por coincidir en número. | En `EliminarOrdenBD`, `EliminarOrdenesLoteBD` y `EliminarTodasOrdenesSedeBD`, antes de insertar en `Ordenes_Historico`, se ejecuta de forma invariable:<br>`DELETE FROM Ordenes_Historico WHERE Residencia = '...' AND NumOrden = ...` (o `IN (...)`). | **Diferencia crítica**:<br>El código actual destruye activamente cualquier expediente histórico previo que comparta el mismo número de orden y sede, sobrescribiéndolo con el nuevo registro archivado. |
| **3. Umbral de Archivado (15 columnas)**:<br>Con $\ge 15$ columnas cumplimentadas (contando columna A): se archiva y se borra de activa.<br>Con $< 15$ columnas: se elimina definitivamente sin crear histórico. | En `EliminarOrdenBD` se evalúa únicamente:<br>`tieneDatosReales = (Len(Trim(DNI)) > 0 Or Len(Trim(Nombre)) > 0)`.<br>En lotes y sede se evalúa en SQL:<br>`((DNI IS NOT NULL AND Trim(DNI) <> '') OR (Nombre IS NOT NULL AND Trim(Nombre) <> ''))`. | **Diferencia crítica**:<br>No se cuentan 15 columnas ni se consideran los otros 26 campos de datos. Si un registro solo tiene DNI o solo Nombre (2 columnas rellenas), se archiva. Si tiene 14 campos rellenos pero DNI y Nombre vacíos, se elimina definitivamente sin archivar. |
| **4. Unicidad del Nº FACTURA**:<br>La combinación `(Residencia, AñoFactura, NumFactura)` debe ser única entre activos e históricos. El año aplicable es el ejercicio de facturación. | El número de factura se almacena en el campo `NumFactura` de `Ordenes` y `Ordenes_Historico`.<br>No existe campo de año de factura en ninguna de las dos tablas.<br>`ExisteFacturaEnBD` consulta exclusivamente `SELECT COUNT(*) FROM Ordenes WHERE Residencia = '...' AND NumFactura = ...`.<br>La tabla auxiliar `ContadorFacturas` sí posee columna `Ejercicio`, pero solo para mantener el secuencial. | **Diferencia crítica**:<br>1. No se almacena el año de facturación en las tablas de datos.<br>2. `ExisteFacturaEnBD` no valida contra `Ordenes_Historico`. Si una factura se archiva a histórico, queda invisible para la validación anti-colisión activa. |
| **5. Reconciliación Offline (SICO-N2)**:<br>Conservación de unicidad en modo contingencia y asignación de identificadores provisionales. | `ObtenerSiguienteNumOrdenProvisional` genera números $\ge 900001$. En la reconexión, `ResolverNumOrdenProvisionalEnBD` consulta `MAX(NumOrden) < 900001` en `Ordenes` y asegura la orden. | **Hecho acreditado**:<br>La resolución provisional consulta únicamente `Ordenes`. No consulta `Ordenes_Historico`. Si el número correlativo asignado existía en histórico de un año anterior, se genera solapamiento entre activa e histórico. |
| **6. Transacciones y Concurrencia**:<br>Garantía de integridad transaccional y protección contra ediciones simultáneas (*Lease Lock*). | `EliminarOrdenBD`, `EliminarOrdenesLoteBD` y `EliminarTodasOrdenesSedeBD` emplean `BeginTrans`/`CommitTrans`/`RollbackTrans` sobre ADODB.<br>`SincronizarFilaHaciaAccess` adquiere Lease Lock atómico dentro de transacción mediante `UPDATE Ordenes SET UsuarioModificacion = ... WHERE ... AND DateDiff >= 120 OR DateDiff < 0`. | **Hecho acreditado**:<br>El mecanismo de Lease Lock está implementado a nivel transaccional. Sin embargo, como se probó en la Orden 1N, el predicado `DateDiff < 0` permite que un desfase horario en un puesto usurpe un bloqueo legítimo. |

---

## 9. CONCLUSIONES Y RESUMEN ESTRUCTURADO (SECCIÓN 10)

### 9.1. Comportamiento Acreditado
1. **Existencia del borrado previo en histórico**: Queda probado de manera incontrovertible que `EliminarOrdenBD` (línea 4378), `EliminarOrdenesLoteBD` (línea 4513) y `EliminarTodasOrdenesSedeBD` (línea 4626) ejecutan sentencias `DELETE FROM Ordenes_Historico` inmediatamente antes de insertar el nuevo registro archivado mediante `INSERT INTO Ordenes_Historico SELECT * FROM Ordenes`. Este es el mecanismo técnico causante de la destrucción de expedientes históricos previos.
2. **Identidad de `Id` en ambas tablas**: Se confirma que el campo `Id` es autonumérico (`dbAutoIncrField`) en ambas tablas, pero `Ordenes_Historico` carece de clave primaria y de restricciones de unicidad. La sintaxis `INSERT INTO Ordenes_Historico SELECT * FROM Ordenes` funciona porque el número de columnas (43), su orden físico y sus tipos son 100% idénticos, copiando el `Id` original de `Ordenes` en `Ordenes_Historico`.
3. **Regla vigente de discriminación de datos**: El archivado ordinario no utiliza ningún umbral de 15 columnas. Actualmente evalúa única y exclusivamente si `DNI` o `Nombre` contienen algún carácter no blanco. Si ambos están vacíos, la orden se elimina sin archivar; si cualquiera tiene texto, se archiva (destruyendo el histórico previo si coincidía en número).
4. **Consultas a tabla activa exclusiva**: Todas las funciones de comprobación y asignación de números (`ExisteOrdenEnBD`, `IntentarReservarOrdenBD`, `ReservarSiguienteOrdenLibreBD`, `AsegurarOrdenExiste`, `ExisteFacturaEnBD`) ejecutan consultas dirigidas exclusivamente contra la tabla `Ordenes`. Ninguna de ellas interroga a `Ordenes_Historico`.
5. **Identidad del procedimiento de sincronización offline**: Se acredita que el procedimiento denominado en antecedentes como `ProcesarColaSincronizacionLocal` corresponde exactamente a `SincronizarColaOfflineHaciaAccess` en `ModuloSeguridadLocal.bas`.

### 9.2. Diferencias Respecto a las Reglas Funcionales
1. **Ausencia del año como clave de unicidad**: Las reglas funcionales establecen que la unicidad debe regir sobre `(Residencia, Año, NumOrden)` y `(Residencia, AñoFactura, NumFactura)`. Actualmente, ni `Ordenes` ni `Ordenes_Historico` disponen de campos de año, y los índices existentes operan exclusivamente sobre `(Residencia, NumOrden)`.
2. **Destrucción frente a preservación histórica**: La regla funcional exige conservar todos los expedientes archivados sin sobreescritura. El código vigente ejecuta borrado preventivo en `Ordenes_Historico` antes de archivar.
3. **Cálculo del umbral de archivado**: La regla funcional exige archivar con $\ge 15$ columnas cumplimentadas (contando la columna A) y eliminar definitivamente con $< 15$. El código vigente evalúa exclusivamente los campos `DNI` y `Nombre`.
4. **Independencia del año de facturación**: La regla funcional exige que el ejercicio de facturación sea independiente del año de la orden y garantice unicidad de factura. Actualmente, el ejercicio solo existe en la tabla auxiliar `ContadorFacturas` alimentado por el reloj del sistema `Year(Now)`, pero no queda registrado en ningún campo del expediente ni en activa ni en histórico.

### 9.3. Información que No Existe o No Ha Podido Verificarse
1. **Campo `Año` en órdenes**: No existe en el esquema de Access ni en los módulos VBA ningún campo que almacene el ejercicio o año de la orden.
2. **Campo `AñoFactura` en órdenes**: No existe en el esquema de Access ni en los módulos VBA ningún campo que almacene el año o ejercicio fiscal de la factura en los expedientes.
3. **Procedimiento `ProcesarColaSincronizacionLocal`**: No existe ningún procedimiento con ese nombre literal en los módulos del libro; su implementación real se denomina `SincronizarColaOfflineHaciaAccess`.

---
**Fin del Informe de Evidencias Técnicas — Orden 1Q.**
