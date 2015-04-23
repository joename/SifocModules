Attribute VB_Name = "SIFOC_Autoempleo"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Name:   putPersonaEnAutoempleoProyecto
'   Desc:   Añade una persona al proyecto de autoempleo
'   Param:  idAutoempleoProyecto, identificador de proyecto de autoempleo
'           idPersona, identificador de la persona donde añadimos la cita
'---------------------------------------------------------------------------
Public Function putPersonaEnAutoempleoProyecto(idAutoempleoProyecto As Long, _
                                               idPersona As Long) As Integer
    On Error GoTo Error
    
    Dim str As String
    str = " INSERT INTO r_AutoempleoProyectopersona (fkAutoempleoProyecto, fkPersona, fkIfocUsuario)" & _
          " VALUES (" & idAutoempleoProyecto & ", " & idPersona & ", " & U_idIfocUsuarioActivo & ");"

'debugando str
    CurrentDb.Execute str
    
    putPersonaEnAutoempleoProyecto = 0
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    putPersonaEnAutoempleoProyecto = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Name:   delPersonaEnAutoempleoProyecto
'   Desc:   Añade una persona al proyecto de autoempleo
'   Param:  idAutoempleoProyecto, identificador de proyecto de autoempleo
'           idPersona, identificador de la persona donde añadimos la cita
'---------------------------------------------------------------------------
Public Function delPersonaAutoempleoProyecto(idProyectoAutoempleo As Long, _
                                             idPersona As Long) As Integer
    On Error GoTo Error
    
    Dim str As String
    str = " DELETE fkProyectoAutoempleo, fkPersona" & _
          " FROM r_proyectoautoempleopersona" & _
          " WHERE fkProyectoAutoempleo=" & idProyectoAutoempleo & " AND fkPersona=" & idPersona & ";"

'debugando str
    CurrentDb.Execute str
    
    delPersonaAutoempleoProyecto = 0
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    delPersonaAutoempleoProyecto = -1
End Function

