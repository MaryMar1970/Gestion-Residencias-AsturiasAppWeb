# Reglas del Proyecto ResidenciaApp

- La ruta principal de este proyecto es `H:\ResidenciaApp`.
- Todas las lecturas, modificaciones, búsquedas y ejecuciones de comandos se realizarán de forma predeterminada sobre la ruta `H:\ResidenciaApp`.
- **Regla estricta de nombres**: Usar siempre el texto literal `GIJÓN` (o `GIJON` donde corresponda). Queda terminantemente prohibido el uso de `Chr(211)`, `ChrW(211)` o concatenaciones tipo `GIJ" & Chr(211) & "N`.
- **Regla de preservación de correcciones de texto y acentos**: El archivo `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` contiene las correcciones manuales de texto y acentuación en español realizadas por el usuario. Queda terminantemente prohibido sobrescribirlo con versiones anteriores (`_OLD.xlsm`, backups o repositorios desactualizados). Cualquier copia o despliegue futuro debe tomar como fuente base única este archivo actualizado o el maestro promovido en `H:\ResidenciaApp\bdas-multiusuario\`.

