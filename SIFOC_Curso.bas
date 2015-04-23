Attribute VB_Name = "SIFOC_Curso"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  13/09/2011  Fecha act.:13/09/2011
'   Name:   isDatosCursoCompletos
'   Descr:  Comprueba que los datos del curso esten rellenos
'   Param:  persona, quien pasa aser alumno id persona
'           curso, id curso
'   Retur:   0, si el alta fue correcta
'           -1, si hubo un error
'---------------------------------------------------------------------------
Public Function isDatosCursoCompletos(idCurso As Long) As Boolean
On Error GoTo Error
    
    Dim isOK As Integer
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    isOK = True
    
    strSql = " SELECT fechaInicio, fechaFin, fkServicio, fkTipoCurso, nombre, fkGrupoFormacion2, fkIfocUsuarioTec, fkIfocUsuarioAux" & _
             " FROM t_curso" & _
             " WHERE id =" & idCurso
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        rs.MoveFirst
        'Campos básicos del curso
        If Not IsNull(rs!fechaInicio) _
             And Not IsNull(rs!fechaFin) _
             And Not IsNull(rs!fkServicio) _
             And Not IsNull(rs!fkTipoCurso) _
             And Not IsNull(rs!nombre) _
             And Not IsNull(rs!fkGrupoFormacion2) _
             And Not IsNull(rs!fkIfocUsuariotec) _
             And Not IsNull(rs!fkIfocUsuarioAux) Then
        Else
           isOK = False
        End If
    End If
        
    rs.Close 'cerramos consulta curso
    
    'Profesor del curso
    strSql = " SELECT fkCurso, fkFormador FROM r_cursoformador WHERE fkCurso =" & idCurso
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    If (rs.EOF) Then
        'Campos básicos del curso
        If rs.RecordCount = 0 Then
           isOK = False
        End If
    End If
    
    'Cerramos profes
    rs.Close
    Set rs = Nothing
    
    isDatosCursoCompletos = isOK
    Exit Function
    
Error:
    debugando "Error(isDatosCursoCompletos): " & Err.description
    isDatosCursoCompletos = False
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  20/09/2011 Fecha act.:20/09/2011
'   Name:   altaServicioDelCurso
'   Descr:  damos alta en servicio del curso a la persona
'   Param:  idCurso, identificador del curso
'           idPersona, identificador de persona
'           idIfocUsuario, identificador ifoc usuario
'   Retur:   0, si el alta fue correcta
'           -1, si hubo un error
'---------------------------------------------------------------------------
Public Function altaServicioDelCurso(idCurso As Long, idPersona As Long, idIfocUsuario As Long) As Integer
    Dim fechaI As Date
    Dim FECHAF As Date
    Dim idServicio As Long
    Dim nombre As String
    Dim numServicios As Integer
    Dim OBS As String
    Dim resultado As Integer
On Error GoTo Error
    
    resultado = 0
    fechaI = DLookup("fechaInicio", "t_curso", "[id]=" & idCurso)
    FECHAF = DLookup("fechaFin", "t_curso", "[id]=" & idCurso)
    idServicio = DLookup("fkServicio", "t_curso", "[id]=" & idCurso)
    nombre = DLookup("nombre", "t_curso", "[id]=" & idCurso)
    OBS = "Alta automática al pasar a ser alumno de curso " & idCurso & " - " & nombre
    
    numServicios = numServiciosUsuarioActivosEnFecha(fechaI, idServicio, idPersona)
    'numServiciosUsuarioActivosEnAmbitoEnFecha(FECHAI, 1, idPersona)
    
    If (numServicios = 0) Then
        'motivo baja, finaliza programa(9)
        resultado = altaServicioUsuario(idServicio, fechaI, idIfocUsuario, FECHAF, 9, , OBS, idPersona)
    Else
        resultado = -1
    End If
    
    altaServicioDelCurso = resultado
    
    Exit Function
Error:
    Debug.Print "Error(isDatosCursoCompletos): " & Err.description
    altaServicioDelCurso = -1
End Function


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  19/03/2009  Fecha act.:13/09/2011
'   Name:   estaPreseleccionadoParaCurso
'   Descr:  Miramos si una persona esta preseleccionada para un curso
'   Param:  persona, quien pasa aser alumno id persona
'           curso, id curso
'   Retur:  true, si esta preseleccionado
'           false, sino
'---------------------------------------------------------------------------
Public Function estaPreseleccionadoParaCurso(idPersona As Long, idCurso As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim strSql As String
    Dim respuesta As Boolean
    
    strSql = " SELECT r_cursopersona.fkPersona, r_cursopersona.fkCurso" & _
             " FROM r_cursopersona" & _
             " WHERE ((fkPersona=" & idPersona & ") AND (fkCurso=" & idCurso & "));"
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        respuesta = True
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    estaPreseleccionadoParaCurso = respuesta
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  19/09/2011  Fecha act.: 19/09/2011
'   Name:   delCursoPersona
'   Descr:  elimina persona de preseleccion de curso
'   Param:  persona, quien pasa aser alumno id persona
'           curso, id curso
'   Retur:   0, si el alta fue correcta
'           -1, si hubo un error
'---------------------------------------------------------------------------
Public Function delCursoPersona(idCurso As Long, _
                                idPersona As Long, _
                                idIfocUsuario As Long) As Integer
    Dim strSql As String
On Error GoTo TratarError
    
    strSql = " DELETE FROM r_cursopersona" & _
             " WHERE fkPersona = " & idPersona & " AND fkCurso = " & idCurso
    CurrentDb.Execute strSql
    
SalirTratarError:
    Exit Function
TratarError:
    delCursoPersona = -1
    debugando Err.description
    Resume SalirTratarError
End Function

'------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  09/07/2009 - Actualizacion: 15/03/2010
'   Name:   updCursoEstadoPersonaPreseleccion
'   Desc:   Actualiza el estado de la persona en la preseleccion del curso
'   Param:  idCurso
'           idPersona
'           idCursoEstadoPreseleccion
'           (o)idCursoMotivoNoSeleccion
'           (o)idTipoInscripcion
'           (o)observaciones
'           (o)valoracion
'           (o)viaAcceso
'   Return: 0 -> OK
'           -1 -> Error
'------------------------------------------------------------------------------
Public Function updCursoEstadoPersonaPreseleccion(idCurso As Long, _
                                                idPersona As Long, _
                                                idIfocUsuario As Long, _
                                                idCursoEstadoPreseleccion As Long, _
                                                Optional idCursoMotivoNoSeleccion As Long = 0, _
                                                Optional idTipoInscripcion As Integer = 0, _
                                                Optional observaciones As String = "", _
                                                Optional valoracion As String = "", _
                                                Optional viaacceso As String = "", _
                                                Optional fecha As Date = "01/01/1900 00:00:00") As Integer
    Dim strSql As String
    Dim fechaL As Date
    Dim OBS As String
 On Error GoTo TratarError
 
    OBS = filterSQL(observaciones)
 
    fechaL = IIf(fecha = "01/01/1900 00:00:00", "01/01/1900 00:00:00", fecha)
    
    'SQL para la actualización del estado del candidato en la tabla r_ofertacandidatos
    strSql = " UPDATE r_cursopersona" & _
             " SET fkIFOCUsuario = " & idIfocUsuario & _
             IIf(idCursoEstadoPreseleccion = 0, "", ", fkCursoEstadoPreseleccion = " & idCursoEstadoPreseleccion) & _
             IIf(idCursoMotivoNoSeleccion = 0, "", ", fkCursoMotivoNoSeleccion = " & idCursoMotivoNoSeleccion) & _
             IIf(idTipoInscripcion = 0, "", ", fkTipoInscripcion =" & idTipoInscripcion) & _
             IIf(observaciones = "", "", ", observaciones = '" & filterSQL(OBS) & "'") & _
             IIf(valoracion = "", "", ", valoracion = '" & valoracion & "'") & _
             IIf(viaacceso = "", "", ", viaacceso = '" & viaacceso & "'") & _
             IIf(fecha = "01/01/1900 00:00:00", "", ", fechaUpd = #" & Format(fechaL, "mm/dd/yyyy hh:mm:nn") & "#") & _
             " WHERE (((r_cursopersona.fkCurso) = " & idCurso & ") AND (r_cursopersona.fkPersona = " & idPersona & "));"
    CurrentDb.Execute strSql
    
    'SQL inserción del nuevo cambio de estado en el histórico de estados del candidato r_cursopersonahistorico
    If (idCursoEstadoPreseleccion <> 0) Then
        strSql = " INSERT INTO r_cursopersonahistorico (fkCurso" & _
                                                        ", fkPersona" & _
                                                        ", fkIFOCUsuario" & _
                                                        ", fkCursoEstadoPreseleccion" & _
                                                        ", fechaUpd" & _
                                                        ", observaciones)"
        strSql = strSql & " VALUES ( " & idCurso & _
                                  ", " & idPersona & _
                                  ", " & idIfocUsuario & _
                                  ", " & idCursoEstadoPreseleccion & _
                                  ", now()" & _
                                  ", '" & OBS & "');"
        CurrentDb.Execute strSql
    End If
    
    updCursoEstadoPersonaPreseleccion = 0
    
SalirTratarError:
    Exit Function
TratarError:
    Debug.Print Err.Source & " - " & Err.description
    updCursoEstadoPersonaPreseleccion = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  19/03/2009  Fecha act.:19/09/2011
'   Name:   insCursoAlumno
'   Descr:  Da de alta un alumno en un curso
'   Param:  persona, quien pasa aser alumno id persona
'           curso, id curso
'   Retur:   0, si el alta fue correcta
'           -1, si hubo un error
'---------------------------------------------------------------------------
Public Function insCursoAlumno(idPersona As Long, _
                               idCurso As Integer, _
                               idIfocUsuario As Long) As Integer
On Error GoTo Error
    
    Dim num As Integer
    Dim sql As String
    
    'Damos alta a persona en curso
    sql = " INSERT INTO r_cursoalumno (fkPersona, fkIfocUsuario, fkCurso, fechaAlta) " & _
          " VALUES (" & idPersona & ", " & idIfocUsuario & ", " & idCurso & ", now());"
    
'debugando sql
    CurrentDb.Execute sql
    
    insCursoAlumno = 0
    Exit Function
    
Error:
    debugando "Error(SIFOC_Personas-insCursoAlumno): " & Err.description
    insCursoAlumno = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  19/03/2009  Fecha act.:14/09/2011
'   Name:   cambioEstadoPermitido
'   Descr:  Cambio de estado de preseleccion en cruso
'   Param:  eActual, estado preseleccion actual
'           eNuevo, estado preseleccion nueva
'   Retur:  true, si es permitido
'           false, si no es permitido
'---------------------------------------------------------------------------
Public Function cambioEstadoCursoPermitido(eActual As Integer, _
                                           eNuevo As Integer) As Boolean
'    If (eActual = 0 And eNuevo = 0) Then 'si no hay cambio estado OK
'        cambioEstadoCursoPermitido = True
'    ElseIf (eActual < eNuevo) Then
'        cambioEstadoCursoPermitido = True
'    Else
'        cambioEstadoCursoPermitido = False
'    End If
    If eNuevo <> 0 Then
        cambioEstadoCursoPermitido = True
    Else
        cambioEstadoCursoPermitido = False
    End If
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  12/09/2011  Fecha act.:12/09/2011
'   Name:   getIdGrupoFormacion2Curso
'   Descr:  insertamos persona en preseleccion curso
'   Param:  idCurso,
'           idPersona
'   Retur:  0, si OK
'           -1, si KO
'---------------------------------------------------------------------------
Public Function getIdGrupoFormacion2Curso(idCurso As Long) As Long
    getIdGrupoFormacion2Curso = Nz(DLookup("fkGrupoFormacion2", "t_curso", "[id]=" & idCurso), 0)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  12/09/2011  Fecha act.:12/09/2011
'   Name:   insCursoPersonaPreseleccion
'   Descr:  insertamos persona en preseleccion curso
'   Param:  idCurso,
'           idPersona
'   Retur:  0, si OK
'           -1, si KO
'---------------------------------------------------------------------------
Public Function insCursoPersonaPreseleccion(idCurso As Long, _
                                            idPersona As Long, _
                                            idIfocUsuario As Long) As Integer
    Dim strSql As String
    Dim idGrupo2 As Long
    Dim fechaIGF2 As Date
On Error GoTo TratarError

    'Calcular la fecha del interes del grupo formativo
    idGrupo2 = Nz(DLookup("[fkGrupoFormacion2]", "t_curso", "[id]=" & idCurso), 0)
    fechaIGF2 = getFechaInsGrupoFormacion2(idPersona, idGrupo2)
    strSql = " INSERT INTO r_cursopersona (fkCurso, fkPersona, fkIfocUsuario, fechaI, fechaUpd)" & _
             " VALUES (" & idCurso & ", " & idPersona & ", " & idIfocUsuario & ", #" & Format(fechaIGF2, "yyyy/mm/dd hh:nn") & "#, now())"
             
    CurrentDb.Execute strSql
SalirTratarError:
    Exit Function
TratarError:
    insCursoPersonaPreseleccion = -1
    debugando Err.description
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  18/10/2013  Fecha act.:18/10/2013
'   Name:   insCursoPersonasPreseleccion
'   Descr:  insertamos persona en preseleccion curso
'   Param:  idCurso,
'           idPersona
'   Retur:  0, si OK
'           -1, si KO
'---------------------------------------------------------------------------
Public Function insCursoPersonasPreseleccion(idCurso As Long, _
                                             idPersonas As String, _
                                             idIfocUsuario As Long) As Integer
    Dim strSql As String
    Dim idGrupo2 As Long
    Dim fechaIGF2 As Date
    Dim numPersonas As Long
    Dim args As Variant
    Dim i As Integer
    Dim idPersona As Long
On Error GoTo TratarError

    'Miramos si todavia se pueden preinscribir personas al curso
    If Not sePuedePreinscribirCurso(Forms!GestionCurso!txt_id) Then
        MsgBox "Ya no se puede inscribir más personas en curso. Posibles causas:" & vbNewLine & _
                "- Ha pasado fecha límite de inscripción." & vbNewLine & _
                "- Ha pasado fecha inicio curso y fecha límite esta vacía.", _
                vbOKOnly, "Alert: SIFOC_Curso"
        
        Exit Function
    End If
    
    'Añadimos listado de ids persona
    numPersonas = countSubStrings(idPersonas, ",")
    args = Split(idPersonas, ",")
    
    For i = 0 To numPersonas - 1
        idPersona = args(i)
        If insCursoPersonaPreseleccion(idCurso, idPersona, idIfocUsuario) = -1 Then
            GoTo TratarError
        End If
    Next
    
'    If (idPersonas <> Null And idPersonas <> "") Then
'        If Not (estaPreseleccionadoParaCurso(Me.lst_personas, Forms("GestionCurso")!txt_id)) Then
'            insCursoPersonaPreseleccion Forms("GestionCurso")!txt_id, Me.lst_personas, idIfocUsuario
'            'Forms("GestionCurso").Controls("Subformulario_GestionCursoPreseleccion").Form.Requery
'        End If
'    End If
    
SalirTratarError:
    Exit Function
TratarError:
    insCursoPersonasPreseleccion = -1
    debugando Err.description
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  14/09/2011  Fecha act.:14/09/2011
'   Name:   isAllowedAccesCurso
'   Descr:  Indica si ifocusuario puede modificar, eliminar personas de curso
'   Param:  idCurso, identificador de curso
'           idPersona, identificador de persona
'           idIfocUsuario, identificador usuario ifoc
'   Retur:  true, puede acceder, modificar, eliminar
'           false, NO puede acceder, modificar, eliminar
'---------------------------------------------------------------------------
Public Function isAllowedAccessCurso(idCurso As Long, _
                                    idPersona As Long, _
                                    idIfocUsuario As Long) As Boolean
    Dim idTecCurso As Long
    Dim idAuxCurso As Long
    Dim access As Boolean
    access = False
    idTecCurso = DLookup("fkIfocUsuarioTec", "t_curso", "[id]=" & idCurso)
    idAuxCurso = DLookup("fkIfocUsuarioAux", "t_curso", "[id]=" & idCurso)
    
    If (idTecCurso = idIfocUsuario) Or (idAuxCurso = idIfocUsuario) Then
        access = True
    ElseIf isTRDePersona(idIfocUsuario, idPersona) Then
        access = True
    End If
    isAllowedAccessCurso = access
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  14/09/2011  Fecha act.:14/09/2011
'   Name:   sePuedeEliminarPreseleccion
'   Descr:  Indica si puede eliminarse persona de preselección
'           no ha realizado ningún cambio de estado
'   Param:  idCurso, identificador de curso
'           idPersona, identificador de persona
'   Retur:  true, se puede eliminar
'           false, NO se puede eliminar
'---------------------------------------------------------------------------
Public Function sePuedeEliminarPreseleccion(idCurso As Long, idPersona As Long) As Boolean
        Dim rs As ADODB.Recordset
    Dim strSql As String
    
    strSql = " SELECT fkCurso" & _
             " FROM r_cursopersonahistorico" & _
             " WHERE fkCurso =" & idCurso & " AND fkPersona=" & idPersona
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        sePuedeEliminarPreseleccion = False
    Else
        sePuedeEliminarPreseleccion = True
    End If
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  10/08/2011  Fecha act.:10/08/2011
'   Name:   isAccessAlumnoAllowed
'   Descr:  Indica si el trabajador puede ver datos de alumno en curso
'   Param:  idCurso, identificador de curso
'           idIfocUsuario, identificador de personal ifoc
'           idPersona, identificador de persona
'   Retur:  true, tiene acceso permitido
'           false, no tiene acceso permitido
'---------------------------------------------------------------------------
Public Function isAccessAlumnoAllowed(idCurso As Long, _
                                      idIfocUsuario As Long, _
                                      Optional idPersona As Long = 0) As Boolean
    Dim idTec As Long
    Dim idAux As Long
    
    'Permisos
    '1.Responsables
    '2.Tec y Aux del curso
    '3.TR de personas
    If ifocUsuarioIdNivel(idIfocUsuario) < 3 Then 'Responsable
        isAccessAlumnoAllowed = True
    ElseIf (Nz(DLookup("fkIfocUsuarioTec", "t_curso", "[id]=" & idCurso), 0) = idIfocUsuario) Or _
           (Nz(DLookup("fkIfocUsuarioAux", "t_curso", "[id]=" & idCurso), 0) = idIfocUsuario) Then
        isAccessAlumnoAllowed = True
    'ElseIf isTRDePersona(idIfocUsuario, idPersona) Then
    '    isAccessAlumnoAllowed = True
    'ElseIf isPersonaAlumno(idCurso, idPersona) Then '¿?
    '    idTec = Nz(DLookup("fkIfocUsuarioTec", "t_curso", "[id]=" & idCurso), 0)
    '    idAux = Nz(DLookup("fkIfocUsuarioAux", "t_curso", "[id]=" & idCurso), 0)
    '
    '    If idTec = idIfocUsuario _
    '            Or idAux = idIfocUsuario _
    '            Or esTRdeUsuario(1, idPersona, idIfocUsuario) Then
    '        isAccessAlumnoAllowed = True
    '    Else
    '        isAccessAlumnoAllowed = False
    '    End If
    Else
        isAccessAlumnoAllowed = False
    End If
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  14/09/2011  Fecha act.:14/09/2011
'   Name:   actualizaPreseleccionEstadoPersonaSMS
'   Descr:  Indica si el trabajador puede ver datos de alumno en curso
'   Param:  idCurso, identificador de curso
'           strPersonas, lista de personas separadas por ','
'           estado,
'   Retur:  true, tiene acceso permitido
'           false, no tiene acceso permitido
'---------------------------------------------------------------------------
'--------------------------------------------------------------------------------------------
'               Actualiza Estado Persona Preseleccionada para curso, por aviso de sms
'--------------------------------------------------------------------------------------------
Public Function actualizaPreseleccionEstadoPersonaSMS(idCurso As Long, _
                                                      strPersonas As String, _
                                                      idCursoEstadoPreseleccion As Long, _
                                                      idIfocUsuario As Long, _
                                                      Optional OBS As String = "", _
                                                      Optional fecha As Date = "01/01/1900 00:00:00")
    Dim numPersonas As Integer
    Dim i As Integer
    Dim persona As String
    Dim idPersona As Long
    Dim correcto As Integer
    
    correcto = 0
    numPersonas = cuentaSubstrings(strPersonas)
    
    For i = 0 To numPersonas - 1 Step 1
        persona = devuelveString(strPersonas, i + 1, ";")
        idPersona = IIf(IsNumeric(persona), CLng(persona), 0)
        
        'CursoEstadoPreseleccion= O - Avisado(1)
        If (updCursoEstadoPersonaPreseleccion(idCurso, idPersona, idIfocUsuario, idCursoEstadoPreseleccion, , , OBS, , , fecha) = -1) Then
            correcto = -1
        End If
    Next
    
    actualizaPreseleccionEstadoPersonaSMS = correcto
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Actual: Jose Manuel Sanchez
'   Fecha:  10/08/2011  Fecha act.:10/08/2011
'   Name:   isAccessAlumnoAllowed
'   Descr:  Indica si el trabajador puede ver datos de alumno en curso
'   Param:  idCurso, identificador de curso
'           idIfocUsuario, identificador de personal ifoc
'           idPersona, identificador de persona
'   Retur:  true, tiene acceso permitido
'           false, no tiene acceso permitido
'---------------------------------------------------------------------------
Public Function isPersonaAlumno(idCurso As Long, _
                                idPersona As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim strSql As String
    
    strSql = " SELECT fkCurso" & _
             " FROM r_cursoalumno" & _
             " WHERE fkCurso =" & idCurso & " AND fkPersona=" & idPersona
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        isPersonaAlumno = True
    Else
        isPersonaAlumno = False
    End If
    
End Function

'--------------------------------------------------------------------------------------------
'       Se puede preinscribir en curso -> Si o No
'--------------------------------------------------------------------------------------------
Public Function sePuedePreinscribirCurso(fkCurso As Long) As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim fechaLimite As Date
    Dim fechaFin As Date
    Dim respuesta As Boolean
    
    respuesta = False
    If (fkCurso > 0) Then
        str = " SELECT t_curso.id, t_curso.fechaInicio, t_curso.fechaFin, t_curso.fechaLimite" & _
              " FROM t_curso" & _
              " WHERE (t_curso.id=" & fkCurso & ");"
        
        Set rs = New ADODB.Recordset
        
        rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
        If Not rs.EOF Then
            rs.MoveFirst
            If IsDate(rs!fechaLimite) Then
            'si pasa la fecha limite no se pueden preinscribir alumnos
                If rs!fechaLimite >= Date Then
                    respuesta = True
                Else
                    respuesta = False
                End If
            ElseIf IsDate(rs!fechaInicio) Then
            'si el curso se ha iniciado no se pueden preinscribir alumnos
                If rs!fechaInicio >= Date Then
                    respuesta = True
                Else
                    respuesta = False
                End If
            End If
        Else 'no existe curso
            respuesta = False
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
    sePuedePreinscribirCurso = respuesta
End Function

'--------------------------------------------------------------------------------------------
'       Es usuario de formacion -> Si o No
'--------------------------------------------------------------------------------------------
'Public Function esUsuarioDeFormacion(fkPersona As Long) As Boolean
'    Dim str As String
'    Dim rs As ADODB.Recordset

'    str = " SELECT t_ifocusuario.fkPersona, t_ifocusuario.fkIfocUnidad" & _
          " FROM t_ifocusuario" & _
          " WHERE (t_ifocusuario.fkIfocUnidad=3)AND (t_ifocusuario.fkIfocSubArea=3) AND (t_ifocusuario.fkPersona=" & fkPersona & ");"

'    Set rs = New ADODB.Recordset
'    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
'    If Not (rs.EOF) Then
'        esUsuarioDeFormacion = True
'    Else
'        esUsuarioDeFormacion = False
'    End If
    
'    rs.Close
'    Set rs = Nothing
'End Function

'--------------------------------------------------------------------------------------------
'           Alta formador a curso
'--------------------------------------------------------------------------------------------
Public Function addFormadorCurso(idCurso As Integer, idFormador As Integer)
    Dim strSql As String
    Dim rs As ADODB.Recordset

    strSql = "r_CursoFormador"

    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
    'Insertamos nuevo formador a curso
    rs.AddNew

    rs!fkCurso = idCurso
    rs!fkFormador = idFormador
    rs!materia = "No especificada"

    rs.update

    rs.Close
    Set rs = Nothing
End Function

'--------------------------------------------------------------------------------------------
'       baja alumno a curso
'--------------------------------------------------------------------------------------------
Public Function delPreseleccionadoCurso(idCurso As Integer, idPersona As Integer)
    Dim strSql As String
    
    strSql = " DELETE * FROM r_cursopersona" & _
             " WHERE fkCurso=" & idCurso & " AND fkPersona=" & idPersona & ";"
    CurrentDb.Execute strSql
End Function

'--------------------------------------------------------------------------------------------
'       baja alumno a curso
'--------------------------------------------------------------------------------------------
Public Function delAlumnoCurso(idCurso As Integer, idPersona As Integer)
    Dim strSql As String
    
    strSql = "DELETE * FROM r_cursoalumno WHERE fkCurso=" & idCurso & " AND fkPersona=" & idPersona & ";"
    CurrentDb.Execute strSql
End Function

'--------------------------------------------------------------------------------------------
'       baja formador a curso
'--------------------------------------------------------------------------------------------
'Public Function delFormadorCurso(idCurso As Integer, idFormador As Integer)
'    Dim strSql As String
'
'    strSql = "DELETE * FROM R_CursoFormador WHERE fkCurso=" & idCurso & " AND fkFormador=" & idFormador & ";"
'    CurrentDb.Execute strSql
'End Function

Public Function delCursoFormador(idCursoFormador As Integer)
    Dim strSql As String
    
    strSql = "DELETE * FROM R_CursoFormador WHERE id=" & idCursoFormador
    CurrentDb.Execute strSql
End Function

'--------------------------------------------------------------------------------------------
'               Miramos si una persona esta preseleccionada para un curso
'--------------------------------------------------------------------------------------------
Public Function esAlumnoDeCurso(idPersona As Long, idCurso As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim strSql As String
    Dim respuesta As Boolean
    
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursoalumno" & _
             " WHERE ((fkPersona=" & idPersona & ") AND (fkCurso=" & idCurso & "));"
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        respuesta = True
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    esAlumnoDeCurso = respuesta
    
End Function

'--------------------------------------------------------------------------------------------
'           Calculamos Estadisticas de GestionCurso
'--------------------------------------------------------------------------------------------
Public Static Function getEstadisticasPreseleccion(idCurso As Long) As Variant
    'caracter separacion ";"
    Dim rsPreseleccion As ADODB.Recordset
    Dim strSql As String
    Dim estPreseleccion(5) As Integer
    
'(0)0 PRESELECCION Total Preselecciones
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursopersona" & _
             " WHERE ((fkCurso)=" & idCurso & ");"
             
    Set rsPreseleccion = New ADODB.Recordset
    rsPreseleccion.Open strSql, CurrentProject.Connection, adOpenStatic, adLockOptimistic
    
    If rsPreseleccion.RecordCount > 0 Then
        rsPreseleccion.MoveLast
        estPreseleccion(0) = rsPreseleccion.RecordCount
    Else
        estPreseleccion(0) = 0
    End If
    rsPreseleccion.Close
    
'(1)1 PRESELECCION S - Aceptado, pasan a ser alumnos
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursopersona" & _
             " WHERE (((fkCurso)=" & idCurso & ") AND ((fkCursoEstadoPreseleccion)=7))"
    'fkCursoEstadoPreseleccion = 7(aceptado)significa que pasa a ser alumno
    
    'utilizamos mismo recordset no importa instanciar nuevo objeto
    rsPreseleccion.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rsPreseleccion.RecordCount > 0 Then
        rsPreseleccion.MoveLast
        estPreseleccion(1) = rsPreseleccion.RecordCount
    Else
        estPreseleccion(1) = 0
    End If
    rsPreseleccion.Close
    
    
'(2)PRESELECCION confirmados
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursopersona" & _
             " WHERE (((fkCurso)=" & idCurso & ") AND (fkCursoEstadoPreseleccion=3))"
    
    Set rsPreseleccion = New ADODB.Recordset
    rsPreseleccion.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    
    If rsPreseleccion.RecordCount > 0 Then
        rsPreseleccion.MoveLast
        estPreseleccion(2) = rsPreseleccion.RecordCount
    Else
        estPreseleccion(2) = 0
    End If
    
    rsPreseleccion.Close
    
    
'(3)3 PRESELECCION En Reserva
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM R_CursoPersona" & _
             " WHERE (((fkCurso)=" & idCurso & ") AND (fkCursoEstadoPreseleccion=6))"
    'fkCursoEstadoPreseleccion = 6(reserva)significa que pasa a ser alumno
    
    Set rsPreseleccion = New ADODB.Recordset
    rsPreseleccion.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rsPreseleccion.RecordCount > 0 Then
        rsPreseleccion.MoveLast
        estPreseleccion(3) = rsPreseleccion.RecordCount
    Else
        estPreseleccion(3) = 0
    End If
    rsPreseleccion.Close
    
    'desvinculamos objeto recordset
    Set rsPreseleccion = Nothing
    
    getEstadisticasPreseleccion = estPreseleccion
End Function

Public Static Function getEstadisticasAlumnos(idCurso As Long) As Variant
    Dim rsAlumnos As ADODB.Recordset
    Dim strSql As String
    Dim estAlumnos(5) As Integer
    
'(1)0 ALUMNOS total alumnos
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursoalumno" & _
             " WHERE (((fkCurso)=" & idCurso & "));"

    Set rsAlumnos = New ADODB.Recordset
    rsAlumnos.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
    If rsAlumnos.RecordCount > 0 Then
        rsAlumnos.MoveLast
        estAlumnos(1) = rsAlumnos.RecordCount
    Else
        estAlumnos(1) = 0
    End If
    rsAlumnos.Close

'(2)1 ALUMNOS aptos
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursoalumno" & _
             " WHERE (((fkCurso)=" & idCurso & ") AND ((fkCalificacionFinal)=1));"
    'fkCalificacionFinal=1 (apto)
    
    'no importa crear nuevo objeto pk emplemos el mismo recordset
    rsAlumnos.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rsAlumnos.RecordCount > 0 Then
        rsAlumnos.MoveLast
        estAlumnos(2) = rsAlumnos.RecordCount
    Else
        estAlumnos(2) = 0
    End If
    rsAlumnos.Close

'(3)2 ALUMNOS no aptos
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursoalumno" & _
             " WHERE (((fkCurso)=" & idCurso & ") AND ((fkCalificacionFinal)=2));"
    'fkCalificacionFinal=1 (apto)
    
    'no importa crear nuevo objeto pk emplemos el mismo recordset
    rsAlumnos.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rsAlumnos.RecordCount > 0 Then
        rsAlumnos.MoveLast
        estAlumnos(3) = rsAlumnos.RecordCount
    Else
        estAlumnos(3) = 0
    End If
    rsAlumnos.Close
    
'(4)2 ALUMNOS baja
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursoalumno" & _
             " WHERE (((fkCurso)=" & idCurso & ") AND (not isnull(fechaBaja)));"
    'fkCalificacionFinal=1 (apto)
    
    'no importa crear nuevo objeto pk emplemos el mismo recordset
    rsAlumnos.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rsAlumnos.RecordCount > 0 Then
        rsAlumnos.MoveLast
        estAlumnos(4) = rsAlumnos.RecordCount
    Else
        estAlumnos(4) = 0
    End If
    rsAlumnos.Close

'(5)2 ALUMNOS acreditación parcial
    strSql = " SELECT fkPersona, fkCurso" & _
             " FROM r_cursoalumno" & _
             " WHERE (((fkCurso)=" & idCurso & ") AND ((fkCalificacionFinal)=4));"
    'fkCalificacionFinal=4 (acreditación parcial)
    
    'no importa crear nuevo objeto pk emplemos el mismo recordset
    rsAlumnos.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rsAlumnos.RecordCount > 0 Then
        rsAlumnos.MoveLast
        estAlumnos(5) = rsAlumnos.RecordCount
    Else
        estAlumnos(5) = 0
    End If
    rsAlumnos.Close
    
    Set rsAlumnos = Nothing
    
    getEstadisticasAlumnos = estAlumnos
End Function

'--------------------------------------------------------------------------------------------
'               CONSULTAS SOBRE FORMACIÓN
'--------------------------------------------------------------------------------------------
'Nos devuelve la fecha donde se cumplira el tanto% del curso
Public Function CursoTantoXCiento(fkCurso As Integer, tantoXciento As Integer) As String
    Dim strSql As String
    Dim rs As dao.Recordset

    Dim fechaI As String
    Dim fechaTantoXciento As Date
    Dim diaSemana As Integer
    Dim minutosCurso As Long
    Dim tantoXcientoMinutos As Long
    Dim i As Integer
    'Guardamos duracion de las clases en funcion del dia 1=lunes..7=domingo
    Dim duracion(1 To 7) As Integer
    
    'Abrimos recordset (para curso)
    strSql = " SELECT id, fechaInicio, horas" & _
             " FROM t_curso" & _
             " WHERE id =" & fkCurso & ";"
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New Dao.Recordset
    'rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        rs.MoveFirst
        fechaI = Nz(rs!fechaInicio, "")
        minutosCurso = Nz(rs!horas, 0) * 60 'pasamos horas a minutos
        tantoXcientoMinutos = (minutosCurso * tantoXciento) / 100
    End If
    'Cerramos recordset (para curso)
    rs.Close
    
    If (fechaI <> "") Then
        'Abrimos recordset (para horario)
        strSql = " SELECT fkCurso, fkDiaSemana, horaInicio, horaFin" & _
                 " FROM r_horarioaulacurso" & _
                 " WHERE fkCurso=" & fkCurso & _
                 " ORDER BY fkDiaSemana ASC;"
        Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
        'rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
        'inicializamos duracion a 0
        For i = 1 To 7 Step 1
            duracion(i) = 0
        Next i
        'Cargamos array duracion de curso cada día
        If Not (rs.EOF) Then
            rs.MoveFirst
            While Not rs.EOF
                'cargamos duracion con el horario
                duracion(rs!fkDiaSemana) = DateDiff("n", rs!horaInicio, rs!horaFin)
                rs.MoveNext
            Wend
        Else
            CursoTantoXCiento = "No calculable"
            Exit Function
        End If
        
        'Cerramos recordset(para horario)
        rs.Close
        Set rs = Nothing
        
        'Calculamos la fecha del % horas
        fechaTantoXciento = fechaI
        While tantoXcientoMinutos > 0
            diaSemana = Weekday(fechaTantoXciento, vbMonday)
            tantoXcientoMinutos = tantoXcientoMinutos - duracion(diaSemana)
            fechaTantoXciento = DateAdd("d", 1, fechaTantoXciento)
        Wend
        CursoTantoXCiento = fechaTantoXciento
    Else
        CursoTantoXCiento = "No hay fecha inicio"
        Exit Function
    End If
    
End Function

Public Function strHorarioCurso(fkCurso As Integer) As String
    Dim rs As dao.Recordset
    Dim strSql As String
    
    Dim strHorario As String
    Dim dia As String
    Dim HORAI As Date
    Dim horaF As Date
    
    strSql = " SELECT fkCurso, fkDiaSemana, horaInicio, horaFin" & _
             " FROM r_horarioaulacurso" & _
             " WHERE fkCurso=" & fkCurso & _
             " ORDER BY fkDiaSemana ASC;"
    
    'Abrimos recordset
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    strHorario = ""
    If Not (rs.EOF) Then
        rs.MoveFirst
        While Not rs.EOF
            dia = DLookup("[diaSemana]", "A_DiaSemana", "[id]=" & rs!fkDiaSemana)
            HORAI = Nz(rs!horaInicio, "00:00:00")
            horaF = Nz(rs!horaFin, "00:00:00")

            rs.MoveNext
            
            'primer caso
            If (rs.EOF) Then 'ultimo caso
                strHorario = strHorario & " y " & dia & " de " & HORAI & " a " & horaF & "."
            ElseIf (strHorario = "") Then
                If (HORAI = rs!horaInicio) And (horaF = rs!horaFin) Then
                    strHorario = dia
                Else
                    strHorario = dia & " de " & HORAI & " a " & horaF
                End If
            Else 'otros casos
                'caso igual anterior
                If (HORAI = rs!horaInicio) And (horaF = rs!horaFin) Then
                    strHorario = strHorario & ", " & dia
                Else 'caso diferente anterior
                    strHorario = strHorario & ", " & dia & " de " & HORAI & " a " & horaF
                End If
            End If
        Wend
    End If
        
    'Cerramos recordset
    rs.Close
    Set rs = Nothing
    
    'Devolvemos
    If (strHorario = "") Then
        strHorarioCurso = "Horario no introducido."
    Else
        strHorarioCurso = strHorario
    End If
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez   Actualiz.: Jose Manuel Sanchez
'   Fecha:  04/09/2012 Fecha act.: 04/09/2012
'   Name:   cursoCnos
'   Descr:  Obtenemos el listado de cnos separados por OR
'   Param:  idCurso, identificador del curso
'   Retur:  string, listado de cursos unidos por OR
'
'---------------------------------------------------------------------------
Public Function cursoCnos(idCurso As Long) As String
    Dim rs As dao.Recordset
    Dim strSql As String
    Dim cnos As String
        
    strSql = " SELECT fkCurso, fkCno2011" & _
             " FROM r_cursocno" & _
             " WHERE fkCurso=" & idCurso
    
    'Abrimos recordset
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    cnos = "0"
    
    If Not (rs.EOF) Then
        rs.MoveFirst
        While Not (rs.EOF)
            If (cnos = "0") Then
                cnos = rs!fkCno2011
            Else
                cnos = cnos & " OR " & rs!fkCno2011
            End If
            rs.MoveNext
        Wend
    End If
    
    'Cerramos recordset
    rs.Close
    Set rs = Nothing
    
    cursoCnos = cnos
End Function

Public Static Function getInfoCursos(idIfocUsuario As Long, fechaInicio As Date, fechaFin As Date) As Variant
    Dim num As Integer
    Dim strSql, sqlPreseleccionados, sqlAlumnos As String
    Dim rs As dao.Recordset
    Dim fechaI, FECHAF As Date
    Dim estCursos(3) As Integer
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"
    
    sqlPreseleccionados = " SELECT fkCurso as idCurso, Count(fkPersona) as preseleccionados" & _
                          " FROM r_cursopersona" & _
                          " GROUP BY fkCurso"
    
    sqlAlumnos = " SELECT fkCurso as idCurso, Count(fkPersona) as alumnos" & _
                 " FROM r_cursoalumno" & _
                 " GROUP BY fkCurso"
    
    strSql = " SELECT Count(t_curso.id) as cursos, Sum(pre.preseleccionados) as preseleccionados, Sum(alu.alumnos) as alumnos" & _
             " FROM (t_curso" & _
             " LEFT JOIN (" & sqlPreseleccionados & ") as pre ON t_curso.id = pre.idCurso)" & _
             " LEFT JOIN (" & sqlAlumnos & ") as alu ON t_curso.id = alu.idCurso" & _
             " WHERE (t_curso.fechaInicio <= #" & FECHAF & "# AND t_curso.fechaFin >= #" & fechaI & "#)" & _
             " AND ((t_curso.fkIFOCUsuarioAux =" & idIfocUsuario & ") OR (t_curso.fkIFOCUsuarioTec =" & idIfocUsuario & "))"
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        estCursos(1) = Nz(rs!cursos, 0)
        estCursos(2) = Nz(rs!preseleccionados, 0)
        estCursos(2) = Nz(rs!alumnos, 0)
    End If
    
    rs.Close
    Set rs = Nothing
    
    getInfoCursos = estCursos
    
End Function

Public Static Function getFiltroCursoAlumnos(strWhere As String) As String
    getFiltroCursoAlumnos = G_Query.getQueryWhere(52, strWhere)
Debug.Print G_Query.getQueryWhere(52, strWhere)
End Function

Public Static Function getFiltroCursoPreseleccion(strWhere As String) As String
    getFiltroCursoPreseleccion = G_Query.getQueryWhere(53, strWhere)
End Function
