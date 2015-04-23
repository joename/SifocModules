Attribute VB_Name = "SIFOC_MejorasTemporales"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  18/03/2009
'   Name:   AltaEnFormacionAlumnosDeCurso
'   Descr:  Alta de alumnos en servicio de formación
'   Param:
'   Retur:
'---------------------------------------------------------------------------

Public Function AltaEnFormacionAlumnosDeCurso() As Integer
    Dim strSql As String
    Dim rs As ADODB.Recordset
    Dim counter As Integer
            
    strSql = " SELECT r_cursoalumno.fkCurso, r_cursoalumno.fkPersona, r_cursoalumno.fechaAlta, t_curso.fechaInicio, t_curso.fechaFin, r_cursoalumno.fkIfocUsuario" & _
             " FROM r_cursoalumno LEFT JOIN t_curso ON r_cursoalumno.fkCurso = t_curso.id" & _
             " WHERE not fechaInicio is null" & _
             " ORDER BY r_cursoalumno.fkCurso, r_cursoalumno.fkPersona;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    counter = 0
    While Not rs.EOF
        ActivaServicioFormacion rs!fkPersona, rs!fkUsuarioIFOC, rs!fechaInicio, rs!fechaFin, 9, rs!fkCurso 'Motivo 9= Finalización programa
        counter = counter + 1
        rs.MoveNext
    Wend
    
    rs.Close
    Set rs = Nothing
    
    AltaEnFormacionAlumnosDeCurso = counter
End Function

