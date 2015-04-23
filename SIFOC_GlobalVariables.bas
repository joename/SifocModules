Attribute VB_Name = "SIFOC_GlobalVariables"
Option Explicit
Option Compare Database

'-------------------------------------------------------------------------------
'       SIFOC global variables
'-------------------------------------------------------------------------------
'>Fecha Version de la APLICACION
Public Const U_organizacion = "Institut de Formacio i Ocupació de Calvià"

'>Fecha Version de la APLICACION
Public Const U_applicationDate = "07/11/2014 12:00:00"

'> Usuario que emplea la base de datos(sólo se toca por formulario INICIO)
Global U_usuario As Long
Public U_idIfocUsuarioActivo As Long    'IfocUsuario de la sesión activa
Public U_idPersona As Long

Public U_fechaFin As Date
Public U_servicio As Long
Public U_idMotivoBaja As Long
Public U_observacion As String

'> Validación de formulario
Global U_clave As String

'> Pruebas
Global Const hola As Integer = 1

'-------------------------------------------------------------------------------
'       BASE DE DATOS
'-------------------------------------------------------------------------------
Global Const server As String = "server"
Global Const DDB1 As String = "user"
Global Const USER_DB1 As String = "user"
Global Const PASS_DB1 As String = "password"

'-------------------------------------------------------------------------------
'       Global functions
'-------------------------------------------------------------------------------
Public Function getOrganizacion() As String
    getOrganizacion = U_organizacion
End Function

'-------
Public Function aaaa()
    Dim d As Date
    
    d = "20/09/1970"
    Debug.Print DateDiff("d", d, now()) / 365
    
    aaaa = IIf(IsNull(d), "", "Edad: " & (Fix(DateDiff("d", d, now()) / 365)) & " años")
End Function
