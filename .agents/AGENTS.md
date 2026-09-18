# Reglas del Proyecto ResidenciaApp

- La ruta principal de este proyecto es `H:\ResidenciaApp`.
- Todas las lecturas, modificaciones, búsquedas y ejecuciones de comandos se realizarán de forma predeterminada sobre la ruta `H:\ResidenciaApp`.
- **Regla estricta de nombres**: Usar siempre el texto literal `GIJÓN` (o `GIJON` donde corresponda). Queda terminantemente prohibido el uso de `Chr(211)`, `ChrW(211)` o concatenaciones tipo `GIJ" & Chr(211) & "N`.
- **Protocolo de Recuperación Instantánea (Golden Master)**: Existe una réplica protegida en modo Solo Lectura en `H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm` y un commit etiquetado en git (`GOLDEN_MASTER_TEXTOS_OK`). Si en cualquier momento se introducen errores en el archivo en uso o en el maestro, se restaurará ejecutando el script `bdas-multiusuario\scripts\Restaurar-Golden-Master.ps1`, el cual restablece de inmediato la versión limpia sin errores ortográficos ni de acentuación.
- **Inmutabilidad del Golden Master**: La réplica Golden Master NUNCA se modificará ni sobrescribirá, salvo petición EXPRESA e inequívoca del usuario indicándolo textualmente.
- **Copias de Seguridad Incrementales Diarias**: En cada sesión o día de trabajo donde se realicen modificaciones sobre el libro Excel o la base de datos, se generará una copia de respaldo con fecha y hora (`BACKUP_AAAA-MM-DD_HHMM`) antes y después de los cambios, asegurando un historial completo de versiones de fácil reversión.

