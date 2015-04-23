Attribute VB_Name = "SIFOC_Empleo"
Option Explicit
Option Compare Database

'--------------------------------------------------------------------------------------------
'               Actualizamos subForms de GestionOferta
'--------------------------------------------------------------------------------------------
'Public Function cambioEstadoOfertaCorrecto(estadoNuevo As Integer, _
                                            estadoactual As Integer) As Boolean
'    If (estadoNuevo > estadoactual) Then
'        cambioEstadoOfertaCorrecto = True
'    Else
'        cambioEstadoOfertaCorrecto = False
'    End If
'End Function

Public Function test() As String
    Dim sql As String
    
    sql = " INSERT INTO t_ofertaestados (fkOferta, fkOfertaEstado, fecha, fkUsuarioIFOC) " & _
          " VALUES (6,1,#" & Format(Date & " " & Time, "mm/dd/yyyy hh:nn:ss") & "#,1);"

'    "  INSERT INTO t_persona" & _
'    " (nombre, apellido1, " & fieldApe2 & "dni, fkSexo, fechaNacimiento, fechaAlta, fkIfocUsuario)" & _
'    " VALUES ('" & Me.txt_nombre & "', '" & Me.txt_apellido1 & "', " & valueApe2 & "'" & Me.txt_dni & "', " & Me.cbx_sexo & ", #" & Format(Me.txt_fechaNacimiento, "mm/dd/yyyy") & "#, #" & Format(Date, "mm/dd/yyyy") & "#, " & usuarioIFOC() & ");"
    
    CurrentDb.Execute sql

End Function

