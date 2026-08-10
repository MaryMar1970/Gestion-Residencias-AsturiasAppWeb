Attribute VB_Name = "ModuloAjusteZoomHojas"
Option Explicit

'------------------------------------------------------------------------------
' M�dulo: ModuloAjusteZoomHojas
'
' Utilidad: Ajustar autom�ticamente el zoom de la ventana activa de Excel
'           para que el rango de columnas especificado de una hoja
'           (ej: A:N, A:M, D:Q, etc.) se muestre al m�ximo ancho posible sin barras horizontales.
'
' FUNCIONAMIENTO Y ROBUSTEZ:
' - Utiliza el m�todo m�s fiable de Excel: seleccionar el rango de columnas deseado
'   y luego usar "ActiveWindow.Zoom = True", que ajusta el zoom para que quepa todo a lo ancho visible.
' - NO se ve afectado por columnas ocultas; solo calcula las visibles.
' - No lanza errores visuales ni interrumpe el trabajo si se da cualquier fallo (p.ej. hoja oculta, sin ventana, par�metros incorrectos).
' - SIEMPRE deja la aplicaci�n en estado estable, con o sin selecci�n previa.
'
' FLEXIBILIDAD PARA FUTURO:
' - Permite f�cilmente adaptar a cualquier hoja y cualquier rango de columnas;
'   s�lo tienes que llamar a esta Sub desde el evento Worksheet_Activate de la hoja deseada,
'   pasando la columna de inicio y la columna de fin como n�meros (A=1, B=2,..., Z=26,...).
' - Ejemplo para una hoja "RESUMEN MADRID" que va de D a Q (columna 4 a 17):
'     Private Sub Worksheet_Activate()
'         Call ModuloAjusteZoomHojas.AjustarZoomAncho(Me, 4, 17)
'     End Sub
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Ajusta el zoom de la ventana activa para mostrar columnas desde colInicial hasta colFinal (A=1, B=2,...)
Public Sub AjustarZoomAncho(ws As Worksheet, colInicial As Long, colFinal As Long)
    On Error GoTo SalidaSinError

    Application.screenUpdating = False ' Previene parpadeos y transiciones raras de pantalla

    Dim colInicialLetra As String, colFinalLetra As String
    Dim rngSel As Range
    Dim wnd As Window

    ' ---- Validaciones robustas de argumentos ----
    If ws Is Nothing Then GoTo SalidaSinError ' No hay hoja, abortar
    If colInicial < 1 Or colFinal < colInicial Then GoTo SalidaSinError ' N�meros incoherentes

    ' ---- Convierte n�mero de columna a letra (A, B, ..., Z, AA, AB, ...) ----
    colInicialLetra = Split(ws.Cells(1, colInicial).Address(False, False), "1")(0)
    colFinalLetra = Split(ws.Cells(1, colFinal).Address(False, False), "1")(0)

    ' ---- Asegura que la ventana activa es v�lida ----
    Set wnd = Application.ActiveWindow
    If wnd Is Nothing Then GoTo SalidaSinError

    ' ---- Activa la hoja (necesario si saltas entre libros/hojas) ----
    ws.Activate

    ' ---- Selecciona solo la PRIMERA celda del rango destino (evita selecciones gigantes) ----
    ws.Cells(1, colInicial).Select

    ' ---- Selecciona el rango de columnas destino (solo cabecera, para zoom correcto) ----
    ws.Range(colInicialLetra & ":" & colFinalLetra).Select

    ' ---- Ejecuta el zoom �ptimo (el m�s fiable de Excel): ajusta al ancho visible de ventana s�lo las columnas visibles ----
    wnd.Zoom = True
    ' ### TRUCO: Selecciona SOLO la primera celda para limpiar el resaltado ###
    ws.Cells(1, colInicial).Select
    Application.screenUpdating = True ' Restablece el refresco de pantalla
    Exit Sub

SalidaSinError:
    Application.screenUpdating = True
    ' Si ocurre cualquier error, la Sub termina silenciosamente y Excel sigue funcionando normal
End Sub

'------------------------------------------------------------------------------
' *** GU�A PARA INCLUIR NUEVAS HOJAS ***
'
' 1. Ve al m�dulo de la hoja nueva, por ejemplo: Hoja "RESUMEN MADRID"
' 2. A�ade el siguiente c�digo:
'
'     Private Sub Worksheet_Activate()
'         ' Ajusta el zoom para mostrar, por ejemplo, de D a Q (columnas 4 a 17)
'         Call ModuloAjusteZoomHojas.AjustarZoomAncho(Me, 4, 17)
'     End Sub
'
' 3. �Listo! El zoom de esa hoja se adaptar� autom�ticamente cuando pulses su pesta�a.
'
' Recuerda: No tienes que tocar el m�dulo externo. S�lo a�ade el evento Worksheet_Activate en la hoja deseada.
'------------------------------------------------------------------------------
