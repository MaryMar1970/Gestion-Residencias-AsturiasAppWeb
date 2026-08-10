VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResultadosSoto 
   Caption         =   "UserForm1"
   ClientHeight    =   8955
   ClientLeft      =   240
   ClientTop       =   900
   ClientWidth     =   13620
   OleObjectBlob   =   "ResultadosSoto.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "ResultadosSoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- C�digo en frmResultados ---
Private Sub lResultadosSoto_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    
    ' Validaci�n 1: Ruta base definida
    If Trim(gFolderPathSoto) = "" Then
        MsgBox "La ruta base no est� definida.", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Validaci�n 2: Elemento seleccionado en la lista
    If Trim(lResultadosSoto.Value) = "" Then
        MsgBox "No se ha seleccionado ning�n archivo.", vbExclamation, "Error"
        Exit Sub
    End If
    
    filePath = gFolderPathSoto & "\" & lResultadosSoto.Value
    
    ' Validaci�n 3: Archivo existe
    If Dir(filePath) = "" Then
        MsgBox "El archivo no se encontr�: " & filePath, vbExclamation, "Error de Archivo"
        Exit Sub
    End If
    
    ' Abrir archivo de forma segura (evitar espacios en rutas)
    Shell "cmd /c start """" """ & filePath & """", vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error inesperado: " & Err.Description, vbCritical, "Error en el Doble Click"
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub

