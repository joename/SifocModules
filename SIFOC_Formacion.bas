Attribute VB_Name = "SIFOC_Formacion"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/1/2009 - Actualización:  21/1/2009
'   Name:   getFechaInsGrupoFormacion2
'   Desc:   Calcula la fecha de inscripcion del interes de una persona en
'           un grupo formativo
'   Param:  idPersona
'           idCurso
'   Retur:  devuelve dniNie correcto NNNNNNNL
'---------------------------------------------------------------------------
Public Function getFechaInsGrupoFormacion2(idPersona As Long, _
                                           idGrupoFormativo2 As Long) As Date
    getFechaInsGrupoFormacion2 = Nz(DLookup("fecha", "t_interesformacion", "[fkPersona]=" & idPersona & " AND fkGrupoFormacion2=" & idGrupoFormativo2), "01/01/1900")
End Function

'-----------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Name:   getNivelTitulacion
'   Desc:   obtiene el nivel de formación de la tabla a_nivelformacion
'   Param:  titulación de la que queremos el nivel de formación
'   Retur:  devuelve id(LONG) de la tabla a_nivel de formación que correspone
'           a la titulación solicitada
'-----------------------------------------------------------------------------
Public Function getNivelTitulacion(titulacion As Long) As Long
    Dim str As String
    Dim idNivel As Long
    Dim rs As ADODB.Recordset
    
    str = " SELECT t_titulacion.id as idTitulacion, a_nivelformacion.id as idNivelFormacion" & _
          " FROM t_titulacion" & _
          " INNER JOIN ((r_titulacionesciclostipo INNER JOIN a_nivelformacion ON r_titulacionesciclostipo.fkNivelFormacion = a_nivelformacion.id)" & _
          " INNER JOIN a_areaformacionregladaciclo ON r_titulacionesciclostipo.fkTitulacionesCiclo = a_areaformacionregladaciclo.id) ON t_titulacion.fkAreaFormacionRegladaCiclo = a_areaformacionregladaciclo.id" & _
          " WHERE t_titulacion.id = " & titulacion & ";"
    
    'Set rs = New ADODB.Recordset
    
    'rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    'If Not rs.EOF Then
    '    idNivel = rs!idNivelFormacion
    'End If
    
    'rs.Close
    'Set rs = Nothing
End Function

'----------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/1/2009 - Actualización:  21/1/2009
'   Name:   insFormacionOcupaional
'   Desc:   Inserta la formación ocupacional de la persona pasada por parámetro
'   Param:  *idPersona(long), identificador de persona
'           *idIfocUsuario(long), identificador usuario ifoc que introduce curso
'           *curso(String), nombre del curso
'           *anyoFin(date), fecha en que finalizó el curso
'           idCurso(long), identificación curso IFOC si se realizó aquí
'           curso(string), nombre del curso
'           horas(int), número de horas del curso
'           idCursoNivel(long), identificador del nivel del curso
'           idPais(long), identificador de país
'           idComunidadAutonoma(long), identificador de comunidad autónoma
'           idEstadoFormacion(long), identificador de formación
'           idAreaFormativa(long), identificador de área formativa
'           idCertificacion(long), identificación de certificación
'           accesoEmpleo(boolean), proporciona acceso al empleo
'           accesoCarneProfesional(boolean), proporciona acceso a carné profesional
'   Retur:   0(int), correcto OK
'           -1(int), error en la operación de inserción
'-----------------------------------------------------------------------------------
Public Function insFormacionOcupacional(idPersona As Long, _
                                        idIfocUsuario As Long, _
                                        nomCurso As String, _
                                        anyoFin As Date, _
                                        Optional idCurso As Long = 0, _
                                        Optional horas As Integer = 0, _
                                        Optional idCursoNivel As Long = 0, _
                                        Optional centro As String = "", _
                                        Optional idPais As Long = 0, _
                                        Optional idComunidadAutonoma As Long = 0, _
                                        Optional idEstadoFormacion As Long = 0, _
                                        Optional idAreaFormativa As Long = 0, _
                                        Optional idCertificacion As Long = 0, _
                                        Optional accesoEmpleo As Boolean = False, _
                                        Optional accesoCarneProfesional As Boolean = False) As Integer
    Dim strSql As String
    Dim strFields As String
    Dim strValues As String
On Error GoTo TratarError
    
    strFields = "(fkPersona" & _
                ", fkIfocUsuario" & _
                ", curso" & _
                ", fechaFin" & _
                IIf(horas = 0, "", ", horas") & _
                IIf(idCurso = 0, "", ", fkCurso") & _
                IIf(idCursoNivel = 0, "", ", fkCursoNivel") & _
                IIf(centro = "", "", ", centro") & _
                IIf(idPais = 0, "", ", fkPais") & _
                IIf(idComunidadAutonoma = 0, "", ", fkComunidadAutonoma") & _
                IIf(idEstadoFormacion = 0, "", ", fkEstadoFormacion") & _
                IIf(idAreaFormativa = 0, "", ", fkAreaFormativa") & _
                IIf(idCertificacion = 0, "", ", fkCertificacion") & _
                IIf(accesoEmpleo = False, "", ", accesoEmpleo") & _
                IIf(accesoCarneProfesional = False, "", ", accesoCarneProfesional") & _
                ")"
    
    strValues = "(" & idPersona & _
                ", " & idIfocUsuario & _
                ", '" & filterSQL(nomCurso) & "'" & _
                ", #" & Format(anyoFin, "mm/dd/yyyy hh:nn:ss") & "#" & _
                IIf(horas = 0, "", ", " & horas) & _
                IIf(idCurso = 0, "", ", " & idCurso) & _
                IIf(idCursoNivel = 0, "", ", " & idCursoNivel) & _
                IIf(centro = "", "", ", '" & filterSQL(centro) & "'") & _
                IIf(idPais = 0, "", ", " & idPais) & _
                IIf(idComunidadAutonoma = 0, "", ", " & idComunidadAutonoma) & _
                IIf(idEstadoFormacion = 0, "", ", " & idEstadoFormacion) & _
                IIf(idAreaFormativa = 0, "", ", " & idAreaFormativa) & _
                IIf(idCertificacion = 0, "", ", " & idCertificacion) & _
                IIf(accesoEmpleo = False, "", ", -1") & _
                IIf(accesoCarneProfesional = False, "", ", -1") & _
                ")"
    
    strSql = " INSERT INTO t_formacionnoreglada" & _
             strFields & _
             " VALUES " & _
             strValues & ";"
    
'Debug.Print strSql
    u_db.Execute strSql
    
SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error (Inserción formación ocupacional): " & vbNewLine & _
            Err.description, , "Alert: Formación - insFormaciónOcupacional"
    insFormacionOcupacional = -1
End Function

Public Function tieneCursoComoFormacionOcupacional(idPersona As Long, _
                                                   idCurso As Long) As Boolean
    Dim rs As dao.Recordset
    Dim strSql As String
    
    strSql = " SELECT id, fkCurso" & _
             " FROM t_formacionnoreglada" & _
             " WHERE fkPersona = " & idPersona & " AND fkCurso =" & idCurso & ";"
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    
    'rs.Open strSQL, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        tieneCursoComoFormacionOcupacional = True
    Else
        tieneCursoComoFormacionOcupacional = False
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

Public Function delInteresFormativo(idPersona As Long, idGrupoFormacion2 As Long) As Integer
    Dim strSql As String
    Dim rs As dao.Recordset
    
On Error GoTo TratarError
    strSql = " DELETE " & _
             " FROM t_interesformacion" & _
             " WHERE fkPersona = " & idPersona & " AND fkGrupoFormacion2 = " & idGrupoFormacion2
    u_db.Execute strSql
    'Set rs = u_db.OpenRecordset(strSQL, dbOpenSnapshot)
    
    delInteresFormativo = 0
    
ExitError:
    Exit Function
TratarError:
    delInteresFormativo = -1
End Function

'..........................................................................................
'                       Probando
'..........................................................................................
Public Function traspasoMasivoFormacionOcupacional()
    Dim strSql As String
    Dim rs As ADODB.Recordset
    Dim counterError As Integer
    Dim counterOk As Integer
    Dim resultado As Integer
    
    strSql = " SELECT r_cursoalumno.fkPersona, t_curso.id AS fkCurso, t_curso.nombre, t_curso.fkCursoNivel, t_curso.horas, t_curso.fechaFin, t_curso.fkUsuarioIFOC, a_grupoformacion2.fkGrupoFormacion1 AS fkAreaFormativa, r_cursoalumno.fkCalificacionFinal" & _
             " FROM (a_grupoformacion2 LEFT JOIN a_areaformativa ON a_grupoformacion2.fkGrupoFormacion1 = a_areaformativa.id) RIGHT JOIN (r_cursoalumno RIGHT JOIN t_curso ON r_cursoalumno.fkCurso = t_curso.id) ON a_grupoformacion2.id = t_curso.fkGrupoFormacion2" & _
             " WHERE ((Not (r_cursoalumno.fkPersona) Is Null) AND (Not (t_curso.id) Is Null) AND (r_cursoalumno.fkCalificacionFinal)=1);"
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    counterError = 0
    counterOk = 0
    
    rs.MoveFirst
    Debug.Print rs.RecordCount
    While Not rs.EOF
        
        If Not IsNull(rs!fkPersona) Then
            'insFormacionOcupacional rs!fkPersona, IIf(Not IsNull(rs!fkUsuarioIFOC), rs!fkUsuarioIFOC, 1), rs!nombre, rs!fechaFin, rs!fkCurso, rs!horas, IIf(Not IsNull(rs!fkCursoNivel), rs!fkCursoNivel, 0), "IFOC", 724, 4, 1, rs!fkAreaFormativa
            resultado = insFormacionOcupacional( _
                                    rs!fkPersona, _
                                    IIf(Not IsNull(rs!fkUsuarioIFOC), rs!fkUsuarioIFOC, 1), _
                                    rs!nombre, _
                                    rs!fechaFin, _
                                    rs!fkCurso, _
                                    IIf(Not IsNull(rs!horas), rs!horas, 0), _
                                    IIf(Not IsNull(rs!fkCursoNivel), rs!fkCursoNivel, 0), _
                                    "IFOC", _
                                    724, _
                                    4, _
                                    1, _
                                    rs!fkAreaFormativa)
            If (resultado = 0) Then
                counterOk = counterOk + 1
            ElseIf (resultado = -1) Then
                counterError = counterError + 1
            End If
        Else
            counterError = counterError + 1
        End If
        rs.MoveNext
    Wend
    
    Debug.Print "OK: " & counterOk & " KO: " & counterError
    rs.Close
    Set rs = Nothing
    
End Function

