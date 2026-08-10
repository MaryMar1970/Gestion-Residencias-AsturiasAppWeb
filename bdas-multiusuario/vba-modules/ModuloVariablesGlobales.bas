Attribute VB_Name = "ModuloVariablesGlobales"
' Introducir mecanismo de bloqueo temporal para evitar ejecuciones redundantes.
' Usaremos una variable p�blica que registre si una actualizaci�n ya fue ejecutada recientemente.
Public evitarActualizacion As Boolean
Public UltimoOrdenLiberado As Variant
Public UltimasHabitacionesLiberadas As String
Public UltimaHojaLiberada As String
