Attribute VB_Name = "InicioGestionResidencias"
' Variable global para almacenar la referencia a la cinta de Office
Dim miRibbon As IRibbonUI

' =============================================================================
' PROCEDIMIENTO: RibbonOnLoad
' PROP�SITO:     Callback que se ejecuta autom�ticamente cuando la cinta se carga
' PAR�METRO:     ribbon - Objeto de la cinta proporcionado por Office
' USO EN XML:    onLoad="RibbonOnLoad"
' ACTUALIZACI�N: 2025-11-20 00:12 - Integraci�n con m�dulo RGPD
' =============================================================================
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    ' ? FLUJO: 1. Almacenar referencia global para uso futuro
    Set miRibbon = ribbon
    
    ' ? FLUJO: 2. Almacenar referencia tambi�n para m�dulo RGPD
    Set ModuloRibbonRGPD.ribbonRGPD = ribbon
    
    ' ? FLUJO: 3. Configurar manejo de errores temporal
    On Error Resume Next
    
    ' ? FLUJO: 4. Activar pesta�a personalizada al cargar
    ribbon.ActivateTab "gestionResidencia"
    ' Nota: Si la pesta�a no existe, On Error Resume Next evita que falle el programa
    
End Sub
