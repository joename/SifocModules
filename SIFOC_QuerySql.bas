Attribute VB_Name = "SIFOC_QuerySql"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/1/2009 - Actualización:  21/1/2009
'   Name:   v_persona_tr
'   Desc:   monta sql de la vista v_persona_tr
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_persona_tr(Optional idIfocUsuario As Long = 0) As String
', Optional idGestionTipo As Integer = 1)
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
    
    strSqlSelect = "r_personaifocusuario.fkPersona AS fkPersona,t_ifocusuario.aka AS aka,t_ifocusuario.fkPersona AS fkIfocUsuario"
    strSqlFrom = "(r_personaifocusuario" & _
                 " LEFT JOIN t_ifocusuario on (r_personaifocusuario.fkIfocUsuario = t_ifocusuario.fkPersona))"
    strSqlWhere = "(r_personaifocusuario.fechaAlta <= now())" & _
                  " and ((r_personaifocusuario.fechaBaja >= now()) or isnull(r_personaifocusuario.fechaBaja))"
    'strSqlWhere = IIf(idIfocUsuario = 0, strSqlWhere, addConditionWhere(strSqlWhere, "r_personaifocusuario.fkIfocAmbito" & idIfocAmbito))
    strSqlWhere = IIf(idIfocUsuario = 0, strSqlWhere, addConditionWhere(strSqlWhere, "fkIfocUsuario =" & idIfocUsuario))
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = " r_personaifocusuario.fkPersona"
    
    v_persona_tr = montarSQL(strSqlSelect, _
                             strSqlFrom, _
                             strSqlWhere, _
                             strSqlGroupBy, _
                             strSqlHaving, _
                             strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/1/2009 - Actualización:  21/1/2009
'   Name:   v_persona_citaultima
'   Desc:   monta sql de la vista v_persona_citaultima
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_persona_citaultima(Optional idIfocUsuario As Long = 0, _
                                     Optional fechaI As Date = "01/01/1900", _
                                     Optional FECHAF As Date = "01/01/1900") As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = "r_citausuario.fkPersona as idPersona, Max(t_cita.fecha) AS FechaUltimaCita"
    
    strSqlFrom = "t_cita INNER JOIN r_citausuario ON t_cita.id = r_citausuario.fkCita"
    
    strSqlWhere = "(r_citausuario.acudeNoacudeAnula = -1)"
    strSqlWhere = IIf(idIfocUsuario = 0, strSqlWhere, addConditionWhere(strSqlWhere, "t_cita.fkIfocUsuarioTec =" & idIfocUsuario))
    strSqlWhere = IIf(fechaI = "01/01/1900", strSqlWhere, addConditionWhere(strSqlWhere, "t_cita.fecha >='" & Format(fechaI, "yyyy-mm-dd") & "'"))
    strSqlWhere = IIf(FECHAF = "01/01/1900", strSqlWhere, addConditionWhere(strSqlWhere, "t_cita.fecha <='" & Format(FECHAF, "yyyy-mm-dd") & "'"))
        
    strSqlGroupBy = "r_citausuario.fkPersona"
    
    strSqlHaving = ""
    
    strSqlOrder = " r_citausuario.fkPersona"
    
    v_persona_citaultima = montarSQL(strSqlSelect, _
                                    strSqlFrom, _
                                    strSqlWhere, _
                                    strSqlGroupBy, _
                                    strSqlHaving, _
                                    strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_persona_citaultima
'   Desc:   monta sql de la vista v_persona_citaultimaTR
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_persona_citaultimaTR(Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
            
    strSqlSelect = "r_citausuario.fkCita AS fkCita,ucita.fkPersona AS idPersona, max(t_cita.fecha) AS fecha,t_cita.fkIfocUsuariotec AS idIfocUsuario"
    
    strSqlFrom = "(t_cita inner join (" & v_persona_citaultima & ") as ucita on (t_cita.id = ucita.fkPersona AND t_cita.fecha = ucita.FechaUltimaCita))"
    
    strSqlWhere = "(r_citausuario.acudeNoacudeAnula = -1)"
    strSqlWhere = IIf(idIfocUsuario = 0, strSqlWhere, addConditionWhere(strSqlWhere, "fkIfocUsuario =" & idIfocUsuario))
    
    strSqlGroupBy = ""
    
    strSqlHaving = ""
    
    strSqlOrder = " r_citausuario.fkPersona"
    
    v_persona_citaultimaTR = montarSQL(strSqlSelect, _
                                        strSqlFrom, _
                                        strSqlWhere, _
                                        strSqlGroupBy, _
                                        strSqlHaving, _
                                        strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_persona_gestionultima
'   Desc:   monta sql de la vista v_persona_gestionultima
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_persona_gestionultima(Optional idIfocUsuario As Long = 0, _
                                        Optional fechaI As Date = "01/01/1900", _
                                        Optional FECHAF As Date = "01/01/1900") As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
        
    strSqlSelect = "r_gestionusuario.fkPersona AS idPersona, Max(t_gestion.fecha) AS FechaUltimaGestion"
    strSqlFrom = " t_gestion" & _
                 " inner join r_gestionusuario on t_gestion.id = r_gestionusuario.fkGestion"
    strSqlWhere = ""
    strSqlWhere = IIf(idIfocUsuario = 0, strSqlWhere, addConditionWhere(strSqlWhere, "t_gestion.fkIfocUsuario =" & idIfocUsuario))
    strSqlWhere = IIf(fechaI = "01/01/1900", strSqlWhere, addConditionWhere(strSqlWhere, "t_gestion.fecha >='" & Format(fechaI, "yyyy-mm-dd") & "'"))
    strSqlWhere = IIf(FECHAF = "01/01/1900", strSqlWhere, addConditionWhere(strSqlWhere, "t_gestion.fecha <='" & Format(FECHAF, "yyyy-mm-dd") & "'"))
    strSqlGroupBy = "r_gestionusuario.fkPersona"
    strSqlHaving = ""
    strSqlOrder = " r_gestionusuario.fkPersona"
    
    v_persona_gestionultima = montarSQL(strSqlSelect, _
                                        strSqlFrom, _
                                        strSqlWhere, _
                                        strSqlGroupBy, _
                                        strSqlHaving, _
                                        strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_persona_serviciosactivos
'   Desc:   monta sql de la vista v_persona_serviciosactivos
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_persona_serviciosactivos() As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
    
    strSqlSelect = "r_serviciousuario.fkPersona as idPersona, group_concat(DISTINCT aka ORDER BY aka ASC SEPARATOR ', ') as servicios"
    strSqlFrom = "r_serviciousuario INNER JOIN a_servicio ON (r_serviciousuario.fkServicio = a_servicio.id)"
    strSqlWhere = "(r_serviciousuario.fechaInicio <= Now()) And ((r_serviciousuario.fechaFin > Now()) Or (r_serviciousuario.fechaFin Is Null))"
    strSqlGroupBy = "r_serviciousuario.fkPersona"
    strSqlHaving = ""
    strSqlOrder = "r_serviciousuario.fkPersona"
    
    v_persona_serviciosactivos = montarSQL(strSqlSelect, _
                                           strSqlFrom, _
                                           strSqlWhere, _
                                           strSqlGroupBy, _
                                           strSqlHaving, _
                                           strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_persona_datos
'   Desc:   monta sql de la vista v_persona_datos
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_persona_datos() As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
    
    strSqlSelect = "t_persona.id AS idPersona,concat_ws(_utf8', ',concat_ws(_utf8' ',t_persona.apellido1,t_persona.apellido2),t_persona.nombre) AS name," & _
            " (((date_format(now(),_utf8'%Y') - date_format(t_persona.fechaNacimiento,_utf8'%Y')) - (date_format(now(),_utf8'00-%m-%d') < date_format(t_persona.fechaNacimiento,_utf8'00-%m-%d')))) AS edad," & _
            " a_nivelempleabilidad.empleabilidad as nivelempleabilidad, t_persona.fechaLastUpdate as fechaLastUpdate"
    strSqlFrom = "(t_persona LEFT JOIN a_nivelempleabilidad ON((t_persona.fkNivelEmpleabilidad = a_nivelempleabilidad.id)))"
    strSqlWhere = ""
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""
    
    v_persona_datos = montarSQL(strSqlSelect, _
                                strSqlFrom, _
                                strSqlWhere, _
                                strSqlGroupBy, _
                                strSqlHaving, _
                                strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_persona_insercionultima
'   Desc:   monta sql de la vista v_persona_datos
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_persona_insercionultima() As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
    
    strSqlSelect = "t_insercion.fkPersona as idPersona, Max(If(IsNull(fechaFin),null,fechafin)) AS FechaFin"
    strSqlFrom = "t_insercion"
    strSqlWhere = ""
    strSqlGroupBy = "t_insercion.fkPersona"
    strSqlHaving = ""
    strSqlOrder = ""
    
    v_persona_insercionultima = montarSQL(strSqlSelect, _
                                        strSqlFrom, _
                                        strSqlWhere, _
                                        strSqlGroupBy, _
                                        strSqlHaving, _
                                        strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_TRusuarios
'   Desc:   monta sql de la vista v_TRusuarios
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_TRusuarios(Optional idIfocUsuario As Long) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
    
    strSqlSelect = "r_personaifocusuario.fkPersona as idPersona, t_ifocusuario.aka, t_ifocusuario.fkPersona AS idIfocUsuario"
    strSqlFrom = "r_personaifocusuario LEFT JOIN t_ifocusuario ON r_personaifocusuario.fkIfocUsuario=t_ifocusuario.fkPersona"
    strSqlWhere = "(((r_personaifocusuario.fechaAlta) <= Now()) And ((r_personaifocusuario.fechaBaja) >= Now())) Or (((r_personaifocusuario.fechaBaja) Is Null))" 'And ((r_personaifocusuario.fkIfocAmbito) = 1)
    strSqlWhere = IIf(idIfocUsuario = 0, strSqlWhere, addConditionWhere(strSqlWhere, "r_personaifocusuario.fkIfocUsuario = " & idIfocUsuario))
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = "r_personaifocusuario.fkPersona"
    
    v_TRusuarios = montarSQL(strSqlSelect, _
                             strSqlFrom, _
                             strSqlWhere, _
                             strSqlGroupBy, _
                             strSqlHaving, _
                             strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_persona_telefonos
'   Desc:   Listado de personas con sus telefonos personales movil y fijo
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_persona_telefonos() As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
    
    strSqlSelect = "t_telefono.fkPersona AS idPersona,group_concat(distinct cast(t_telefono.telefono as char(9) charset utf8) order by t_telefono.fkTipoTelefono2 DESC separator ', ') AS telefonos"
    strSqlFrom = "t_telefono"
    strSqlWhere = "((t_telefono.fkTelefonoTipo = 1) and (t_telefono.fkTipoTelefono1 = 1))"
    strSqlGroupBy = "`t_telefono`.`fkPersona`"
    strSqlHaving = "(t_telefono.fkPersona is not null)"
    strSqlOrder = ""
    
    v_persona_telefonos = montarSQL(strSqlSelect, _
                                    strSqlFrom, _
                                    strSqlWhere, _
                                    strSqlGroupBy, _
                                    strSqlHaving, _
                                    strSqlOrder)
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_usuariosTR_ultimas
'   Desc:   monta sql de la vista v_usuariosTR_ultimas
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_usuariosTR_ultimas(Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String
    
    strSqlSelect = "personas.idPersona, personas.name as Nombre, personas.edad, personas.nivelempleabilidad as nivel, date(personas.fechaLastUpdate) as actualizacion, sactivos.servicios, '' as telefonos, date_format(personas.fechaLastUpdate,'%Y%m%d %H:%i:%s') as fechaactualizacion"
    strSqlFrom = "(((" & v_TRusuarios(idIfocUsuario) & ") as trusuarios" & _
                 " INNER JOIN (" & v_persona_datos() & ") as personas ON trusuarios.idPersona = personas.idPersona )" & _
                 " LEFT JOIN (" & v_persona_serviciosactivos & ") as sactivos ON personas.idPersona = sactivos.idPersona)"
                 '" LEFT JOIN (" & v_persona_telefonos() & ") as telefonos ON trusuarios.idPersona = telefonos.idPersona)" & _

    'Consulta anterior
    'strSqlSelect = "personas.idPersona, personas.name as Nombre, personas.edad, personas.nivelempleabilidad as nivel, date(personas.fechaLastUpdate) as actualizacion, date(ucita.FechaUltimaCita) as cita, date(ugestion.FechaUltimaGestion) as gestion, date(uinsercion.FechaFin) as Insercion, sactivos.servicios, telefonos.telefonos, date_format(personas.fechaLastUpdate,'%Y%m%d %H:%i:%s') as fechaactualizacion, date_format(ucita.FechaUltimaCita,'%Y%m%d %H:%i:%s') as fechacita, date_format(ugestion.FechaUltimaGestion,'%Y%m%d %H:%i:%s') as fechagestion, date_format(uinsercion.FechaFin,'%Y%m%d %H:%i:%s') as fechaInsercion"
    'strSqlFrom = "(((((" & v_TRusuarios(idIfocUsuario) & ") as trusuarios" & _
                 " INNER JOIN (" & v_persona_datos & ") as personas ON trusuarios.idPersona = personas.idPersona )" & _
                 " LEFT JOIN (" & v_persona_telefonos() & ") as telefonos ON trusuarios.idPersona = telefonos.idPersona" & _
                 " LEFT JOIN (" & v_persona_serviciosactivos & ") as sactivos ON personas.idPersona = sactivos.idPersona)" & _
                 " LEFT JOIN (" & v_persona_citaultima & ") as ucita ON personas.idPersona = ucita.idPersona)" & _
                 " LEFT JOIN (" & v_persona_gestionultima & ") as ugestion ON personas.idPersona = ugestion.idPersona)" & _
                 " LEFT JOIN (" & v_persona_insercionultima & ") as uinsercion ON personas.idPersona = uinsercion.idPersona"
    strSqlWhere = ""
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = "personas.idPersona"
    
    v_usuariosTR_ultimas = montarSQL(strSqlSelect, _
                                    strSqlFrom, _
                                    strSqlWhere, _
                                    strSqlGroupBy, _
                                    strSqlHaving, _
                                    strSqlOrder)
Debug.Print "TRultimas > " & v_usuariosTR_ultimas
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_oferta_candidatos
'   Desc:   usuarios de TR que aparecen en ofertas
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_oferta_candidatos(Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = "r_ofertacandidatos.fkOferta as idOferta, r_ofertacandidatos.fkPersona as idPersona, personas.name as nombre, r_ofertacandidatos.fecha, a_ofertasegestado.estado, r_ofertacandidatos.finSeguimiento, r_ofertacandidatos.fkOfertaSegEstado"
    strSqlFrom = " (r_ofertacandidatos" & _
                 " INNER JOIN a_ofertasegestado ON r_ofertacandidatos.fkOfertaSegEstado = a_ofertasegestado.id)" & _
                 " INNER JOIN (" & v_persona_datos & ") as personas ON r_ofertacandidatos.fkPersona = personas.idPersona"
    strSqlWhere = ""
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""
    
    v_oferta_candidatos = montarSQL(strSqlSelect, _
                                    strSqlFrom, _
                                    strSqlWhere, _
                                    strSqlGroupBy, _
                                    strSqlHaving, _
                                    strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_oferta_candidatosTR
'   Desc:   usuarios de TR que aparecen en ofertas
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_oferta_candidatosTR(Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = " t_oferta.id as idOferta, t_oferta.puesto, date(t_oferta.fechaOferta) as oferta, a_ofertaestado.descripcion AS estadoO," & _
                   " ocandidatos.idPersona, ocandidatos.Nombre, '' as telefonos, date(ocandidatos.fecha) as fecha, ocandidatos.estado," & _
                   " date_format(t_oferta.fechaOferta,'%Y%m%d %H:%i:%s') as fechaOferta, date_format(ocandidatos.fecha,'%Y%m%d %H:%i:%s') as fechaEstado," & _
                   " ocandidatos.finSeguimiento, ocandidatos.fkOfertaSegEstado as idEstadoCandidato, fkOfertaEstado as idEstadoOferta"

    strSqlFrom = " ((((" & v_TRusuarios(idIfocUsuario) & ") as trusuarios " & vbNewLine & _
                 " INNER JOIN (" & v_oferta_candidatos & ") as ocandidatos ON trusuarios.idPersona = ocandidatos.idPersona)" & vbNewLine & _
                 " INNER JOIN t_oferta ON ocandidatos.idOferta = t_oferta.id)" & vbNewLine & _
                 " INNER JOIN a_ofertaestado ON t_oferta.fkOfertaEstado = a_ofertaestado.id)"

    'Antigua consulta muy lenta
    'strSqlSelect = " t_oferta.id as idOferta, t_oferta.puesto, date(t_oferta.fechaOferta) as oferta, a_ofertaestado.descripcion AS estadoO," & _
                   " ocandidatos.idPersona, ocandidatos.Nombre, telefonos.telefonos, date(ocandidatos.fecha) as fecha, ocandidatos.estado," & _
                   " date_format(t_oferta.fechaOferta,'%Y%m%d %H:%i:%s') as fechaOferta, date_format(ocandidatos.fecha,'%Y%m%d %H:%i:%s') as fechaEstado," & _
                   " ocandidatos.finSeguimiento, ocandidatos.fkOfertaSegEstado as idEstadoCandidato, fkOfertaEstado as idEstadoOferta"
    
    'strSqlFrom = " ((t_oferta" & _
                 " INNER JOIN a_ofertaestado ON t_oferta.fkOfertaEstado = a_ofertaestado.id)" & _
                 " INNER JOIN (" & v_oferta_candidatos & ") as ocandidatos ON t_oferta.id = ocandidatos.idOferta)" & _
                 " INNER JOIN (" & v_TRusuarios(idIfocUsuario) & ") as trusuarios ON ocandidatos.idPersona = trusuarios.idPersona" & _
                 " LEFT JOIN (" & v_persona_telefonos() & ") as telefonos ON ocandidatos.idPersona = telefonos.idPersona"

    strSqlWhere = ""
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""
    
    v_oferta_candidatosTR = montarSQL(strSqlSelect, _
                                      strSqlFrom, _
                                      strSqlWhere, _
                                      strSqlGroupBy, _
                                      strSqlHaving, _
                                      strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_citas
'   Desc:   citas de usuarios de TR que aparecen en ofertas
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_citas(Optional fechaInicio As Date = "01/01/1900", _
                        Optional fechaFin As Date = "01/01/1900", _
                        Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = "t_cita.id as idCita, a_gestiontipo.tipo, a_citasesion.sesion, t_cita.fecha, a_servicio.aka AS servicio, t_cita.fkIfocUsuarioTec as idIfocUsuario, t_ifocusuario.aka as citador, if(cancelada = 0, 'No', 'Si') as Cancelada, t_cita.fkGestionTipo as idGestionTipo"
    strSqlFrom = " (((t_cita" & _
                 " LEFT JOIN t_ifocusuario ON t_cita.fkIfocUsuarioCit = t_ifocUsuario.fkPersona)" & _
                 " LEFT JOIN a_citasesion ON t_cita.fkCitaSesion = a_citasesion.id)" & _
                 " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
                 " LEFT JOIN a_servicio ON t_cita.fkServicio = a_servicio.id"
                 
    strSqlWhere = IIf(fechaInicio <> "01/01/1900" And fechaFin <> "01/01/1900", _
                  "t_cita.fecha between '" & Format(fechaInicio, "yyyy-mm-dd") & "' AND '" & Format(fechaFin, "yyyy-mm-dd 23:55") & "'", _
                  "")
    strSqlWhere = IIf(idIfocUsuario <> 0, addConditionWhere(strSqlWhere, "t_cita.fkIfocUsuarioTec=" & idIfocUsuario), strSqlWhere)
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""
    
    v_citas = montarSQL(strSqlSelect, _
                        strSqlFrom, _
                        strSqlWhere, _
                        strSqlGroupBy, _
                        strSqlHaving, _
                        strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_citas_personasTR
'   Desc:   citas de usuarios de TR que aparecen en ofertas
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_citas_personasTR(Optional fechaInicio As Date = "01/01/1900", _
                                   Optional fechaFin As Date = "01/01/1900", _
                                   Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = "citas.idCita, citas.tipo, citas.sesion, date(citas.fecha) as fecha, date_format(citas.fecha, '%H:%i') AS Hora," & _
                   " COALESCE(citas.servicio,'') as servicio, citas.citador, t_ifocusuario.aka as Tec, citas.cancelada, " & _
                   " If(r_citausuario.acudeNoacudeAnula=-1,'Si'," & _
                        "If(r_citausuario.acudeNoacudeAnula=0,'NO','Anula')) AS Acude," & _
                   " IF(ISNULL(r_citausuario.fkPersona),0,r_citausuario.fkPersona) as idUsuario," & _
                   " v_apellidosnombre.name as nombre," & _
                   " date_format(citas.fecha,'%Y%m%d %H:%i:%s') as fechaOrder," & _
                   " COALESCE(trusuarios.idIfocUsuario,0) as idIfocUsuarioTR, citas.idIfocUsuario"
    strSqlFrom = " ((((r_citausuario" & _
                 " RIGHT JOIN (" & v_citas(fechaInicio, fechaFin) & ") as citas ON r_citausuario.fkCita = citas.idCita) " & _
                 " LEFT JOIN (" & v_TRusuarios(idIfocUsuario) & ") as trusuarios ON r_citausuario.fkPersona = trusuarios.idPersona)" & _
                 " LEFT JOIN t_ifocusuario ON citas.idIfocUsuario = t_ifocusuario.fkPersona)" & _
                 " LEFT JOIN v_apellidosnombre ON r_citausuario.fkPersona = v_apellidosnombre.id)"
                 
    strSqlWhere = ""
    If (fechaInicio <> "01/01/1900" And fechaFin <> "01/01/1900") Then
        strSqlWhere = addConditionWhere(strSqlWhere, _
                            "fecha between " & Format(fechaInicio, "'yyyy-mm-dd'") & _
                            " AND " & Format(fechaFin, "'yyyy-mm-dd 23:59:59'"))
    End If
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""
    
    v_citas_personasTR = montarSQL(strSqlSelect, _
                                    strSqlFrom, _
                                    strSqlWhere, _
                                    strSqlGroupBy, _
                                    strSqlHaving, _
                                    strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_gestiones
'   Desc:   getiones de usuarios de TR que aparecen en ofertas
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_gestiones(Optional fechaInicio As Date = "01/01/1900", _
                            Optional fechaFin As Date = "01/01/1900", _
                            Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = "t_gestion.id as idGestion, a_gestiontipo.tipo, t_gestion.fecha, date_format(t_gestion.fecha,'%H:%i') as hora, a_formacontacto.descripcion AS Contacto, a_servicio.aka AS servicio, t_gestion.fkIfocUsuario as idIfocUsuario, t_gestion.fkGestionTipo as idGestionTipo"
    strSqlFrom = " (((t_gestion" & _
                 " LEFT JOIN t_ifocusuario ON t_gestion.fkIfocUsuario = t_ifocUsuario.fkPersona)" & _
                 " LEFT JOIN a_gestiontipo ON t_gestion.fkGestionTipo = a_gestiontipo.id)" & _
                 " LEFT JOIN a_formaContacto ON t_gestion.fkFormaContacto = a_formacontacto.id)" & _
                 " LEFT JOIN a_servicio ON t_gestion.fkServicio = a_servicio.id"
    
    strSqlWhere = IIf(fechaInicio <> "01/01/1900" And fechaFin <> "01/01/1900", _
                  "t_gestion.fecha between '" & Format(fechaInicio, "yyyy-mm-dd") & "' AND '" & Format(fechaFin, "yyyy-mm-dd 23:55") & "'", _
                  "")
    strSqlWhere = IIf(idIfocUsuario <> 0, addConditionWhere(strSqlWhere, "t_gestion.fkIfocUsuario=" & idIfocUsuario), strSqlWhere)
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""
    
    v_gestiones = montarSQL(strSqlSelect, _
                            strSqlFrom, _
                            strSqlWhere, _
                            strSqlGroupBy, _
                            strSqlHaving, _
                            strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_gestiones_personasTR
'   Desc:   gestiones de usuarios de TR que aparecen en ofertas
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_gestiones_personasTR(Optional fechaInicio As Date = "01/01/1900", _
                                       Optional fechaFin As Date = "01/01/1900", _
                                       Optional idIfocUsuario As Long = 0) As String
    Dim strSql As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = "gestiones.idGestion, gestiones.tipo, date_format(gestiones.fecha, '%d/%m/%Y') as fecha, date_format(gestiones.fecha, '%H:%i') AS Hora," & _
                   " IF(ISNULL(gestiones.servicio),'',gestiones.servicio) as servicio, t_ifocusuario.aka as IFOCUsuario, " & _
                   " IF(gestiones.idGestionTipo=4,COALESCE(r_gestionusuario.fkOrganizacion,0),COALESCE(r_gestionusuario.fkPersona,0)) as idUsuario," & _
                   " IF(gestiones.idGestionTipo=4,t_organizacion.nombre, v_apellidosnombre.name) as nombre," & _
                   " date_format(gestiones.fecha,'%Y%m%d %H:%i:%s') as fechaOrder," & _
                   " IF(ISNULL(trusuarios.idIfocUsuario),0,trusuarios.idIfocUsuario) as idIfocUsuarioTR, gestiones.idIfocUsuario"
    strSqlFrom = " ((((r_gestionusuario" & _
                 " RIGHT JOIN (" & v_gestiones(fechaInicio, fechaFin) & ") as gestiones ON r_gestionusuario.fkGestion = gestiones.idGestion) " & _
                 " LEFT JOIN (" & v_TRusuarios(idIfocUsuario) & ") as trusuarios ON r_gestionusuario.fkPersona = trusuarios.idPersona)" & _
                 " LEFT JOIN t_ifocusuario ON gestiones.idIfocUsuario = t_ifocusuario.fkPersona)" & _
                 " LEFT JOIN v_apellidosnombre ON r_gestionusuario.fkPersona = v_apellidosnombre.id)" & _
                 " LEFT JOIN t_organizacion ON r_gestionusuario.fkOrganizacion = t_organizacion.id"
                 
    '" IF(gestiones.idIfocAmbito=2,IF(ISNULL(r_gestionusuario.fkOrganizacion),0,r_gestionusuario.fkOrganizacion),IF(ISNULL(r_gestionusuario.fkPersona),0,r_gestionusuario.fkPersona)) as idUsuario," & _

    strSqlWhere = ""
    If (fechaInicio <> "01/01/1900" And fechaFin <> "01/01/1900") Then
        strSqlWhere = addConditionWhere(strSqlWhere, _
                            "fecha between " & Format(fechaInicio, "'yyyy-mm-dd'") & _
                            " AND " & Format(fechaFin, "'yyyy-mm-dd 23:59:59'"))
    End If
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""
    
    strSql = montarSQL(strSqlSelect, _
                       strSqlFrom, _
                       strSqlWhere, _
                       strSqlGroupBy, _
                       strSqlHaving, _
                       strSqlOrder)
    
'Debug.Print strSql
    
    v_gestiones_personasTR = strSql
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_curso_preseleccionTR
'   Desc:   Preselecciones de usuarios de TR que tienen preselecciones
'           en la fecha pasada por parámetro
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_curso_preseleccionTR(Optional fecha As Date = "01/01/1900", _
                                       Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = "r_cursopersona.fkCurso AS idCurso, a_tipocurso.aka AS Tipo, CONCAT(t_curso.nombre,' (',COALESCE(a_cursonivel.nivel,''), ')') AS Curso, date(T_Curso.fechaInicio) AS Inicio, date(T_Curso.fechaFin) AS Fin," & _
                   " date(r_cursopersona.fechaI) AS FechaI, r_cursopersona.fkPersona as idPersona, v_apellidosnombre.name as nombre, A_CursoEstadoPreseleccion.estado AS Estado," & _
                   " A_TipoInscripcion.tipoInscripcion as TI, r_cursopersona.valoracion, date_format(T_Curso.fechaInicio,'%Y%m%d %H:%i:%s') as fechaInicio, date_format(T_Curso.fechaFin,'%Y%m%d %H:%i:%s') as fechaFin"
    strSqlFrom = " (((((((r_cursopersona" & _
                 " INNER JOIN (" & v_TRusuarios(idIfocUsuario) & ") as trusuarios ON r_cursopersona.fkPersona = trusuarios.idPersona)" & _
                 " INNER JOIN T_Curso ON r_cursopersona.fkCurso = T_Curso.id)" & _
                 " INNER JOIN v_apellidosnombre ON r_cursopersona.fkPersona = v_apellidosnombre.id)" & _
                 " LEFT JOIN A_TipoInscripcion ON r_cursopersona.fkTipoInscripcion = A_TipoInscripcion.id)" & _
                 " LEFT JOIN A_CursoEstadoPreseleccion ON r_cursopersona.fkCursoEstadoPreseleccion = A_CursoEstadoPreseleccion.id)" & _
                 " LEFT JOIN A_CursoMotivoNoSeleccion ON R_CursoPersona.fkCursoMotivoNoSeleccion = A_CursoMotivoNoSeleccion.id)" & _
                 " LEFT JOIN a_tipocurso ON T_Curso.fkTipoCurso = a_tipocurso.id)" & _
                 " LEFT JOIN a_cursonivel ON T_Curso.fkCursoNivel = a_cursonivel.id"
                 
    strSqlWhere = ""
    If (fecha <> "01/01/1900") Then
        strSqlWhere = addConditionWhere(strSqlWhere, _
                            "T_Curso.fechaFin >= " & Format(fecha, "'yyyy-mm-dd'"))
    End If
'    If (idIfocUsuario <> 0) Then
'        strSqlWhere = addConditionWhere(strSqlWhere, "trusuarios.idIfocUsuario = " & idIfocUsuario)
'    End If
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""
    
'Debug.Print montarSQL(strSqlSelect, _
                      strSqlFrom, _
                      strSqlWhere, _
                      strSqlGroupBy, _
                      strSqlHaving, _
                      strSqlOrder)
    
    v_curso_preseleccionTR = montarSQL(strSqlSelect, _
                                        strSqlFrom, _
                                        strSqlWhere, _
                                        strSqlGroupBy, _
                                        strSqlHaving, _
                                        strSqlOrder)
End Function


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_curso_alumnoTR
'   Desc:   alumnos de usuarios de TR que tienen preselecciones
'           en la fecha pasada por parámetro
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_curso_alumnoTR(Optional fecha As Date = "01/01/1900", _
                                 Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

'SELECT
'from
'GROUP BY t_cursoalumnoasistencia.fkPersona, t_cursoalumnoasistencia.fkCurso]. AS AsistenciaAlumnoCurso ON (r_cursoalumno.fkPersona = AsistenciaAlumnoCurso.fkPersona) AND (r_cursoalumno.fkCurso = AsistenciaAlumnoCurso.fkCurso)) LEFT JOIN a_calificacionfinal ON r_cursoalumno.fkCalificacionFinal = a_calificacionfinal.id) LEFT JOIN a_motivobajacurso ON r_cursoalumno.fkMotivoBaja = a_motivobajacurso.id) LEFT JOIN a_cursonivel ON T_curso.fkCursoNivel = a_cursonivel.id) LEFT JOIN a_tipocurso ON T_curso.fkTipoCurso = a_tipocurso.id
'WHERE (((r_cursoalumno.fkPersona) = 441))
'ORDER BY T_curso.fechaInicio;

    strSqlSelect = "T_curso.id as idCurso, a_tipocurso.aka AS Tipo," & _
                   " CONCAT(t_curso.nombre,' (',COALESCE(a_cursonivel.nivel,''), ')') AS Curso," & _
                   " T_curso.fechaInicio AS Inicio, T_curso.fechaFin AS Fin, T_curso.horas," & _
                   " r_cursoalumno.fkPersona as idPersona, v_apellidosnombre.name as nombre, CONCAT(Format((alumnodatos.asistencias*100)/t_curso.horas,0), If(asistencias,' %','0 %')) AS Asist," & _
                   " If(a_calificacionfinal.calificacion='Baja',CONCAT(a_calificacionfinal.calificacion, ' + ', a_motivobajacurso.motivoBaja),a_calificacionfinal.calificacion) AS Resultado," & _
                   " r_cursoalumno.entregaCertificado AS Certificado, t_curso.fechaInicio, t_curso.fechaFin"
    strSqlFrom = " ((((((T_curso" & _
                 " INNER JOIN r_cursoalumno ON T_curso.id = r_cursoalumno.fkCurso)" & _
                 " INNER JOIN (" & v_TRusuarios(idIfocUsuario) & ") as trusuarios ON r_cursoalumno.fkPersona = trusuarios.idPersona)" & _
                 " INNER JOIN v_apellidosnombre ON r_cursoalumno.fkPersona = v_apellidosnombre.id)" & _
                 " LEFT JOIN (SELECT t_cursoalumnoasistencia.fkPersona, t_cursoalumnoasistencia.fkCurso, Sum(t_cursoalumnoasistencia.asistencias) AS asistencias FROM t_cursoalumnoasistencia GROUP BY fkPersona, fkCurso) as alumnodatos ON (r_cursoalumno.fkCurso = alumnodatos.fkCurso AND r_cursoalumno.fkPersona = alumnodatos.fkPersona)" & _
                 " LEFT JOIN a_tipocurso ON t_curso.fkTipoCurso=a_tipocurso.id)" & _
                 " LEFT JOIN a_cursonivel ON t_curso.fkCursoNivel=a_cursonivel.id)" & _
                 " LEFT JOIN a_calificacionfinal ON r_cursoalumno.fkCalificacionFinal=a_calificacionfinal.id)" & _
                 " LEFT JOIN a_motivobajacurso ON r_cursoalumno.fkMotivoBaja=a_motivobajacurso.id"
                 
    strSqlWhere = ""
    If (fecha <> "01/01/1900") Then
        strSqlWhere = addConditionWhere(strSqlWhere, _
                            "T_Curso.fechaFin >= " & Format(fecha, "'yyyy-mm-dd'"))
    End If
    If (idIfocUsuario <> 0) Then
        strSqlWhere = addConditionWhere(strSqlWhere, "trusuarios.idIfocUsuario = " & idIfocUsuario)
    End If
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""

    v_curso_alumnoTR = montarSQL(strSqlSelect, _
                                strSqlFrom, _
                                strSqlWhere, _
                                strSqlGroupBy, _
                                strSqlHaving, _
                                strSqlOrder)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/06/2011 - Actualización:  06/06/2011
'   Name:   v_oferta_candidatosTR
'   Desc:   usuarios de TR que aparecen en ofertas
'
'   Param:  -
'   Retur:  sql de vista
'---------------------------------------------------------------------------
Public Function v_oferta_candidatosTRtmppruebas(Optional idIfocUsuario As Long = 0) As String
    Dim strSqlSelect As String
    Dim strSqlFrom As String
    Dim strSqlWhere As String
    Dim strSqlGroupBy As String
    Dim strSqlHaving As String
    Dim strSqlOrder As String

    strSqlSelect = " t_oferta.id as idOferta, t_oferta.puesto, date(t_oferta.fechaOferta) as oferta, a_ofertaestado.descripcion AS estadoO," & _
                   " ocandidatos.idPersona, ocandidatos.Nombre, '' as telefonos, date(ocandidatos.fecha) as fecha, ocandidatos.estado," & _
                   " date_format(t_oferta.fechaOferta,'%Y%m%d %H:%i:%s') as fechaOferta, date_format(ocandidatos.fecha,'%Y%m%d %H:%i:%s') as fechaEstado," & _
                   " ocandidatos.finSeguimiento, ocandidatos.fkOfertaSegEstado as idEstadoCandidato, fkOfertaEstado as idEstadoOferta"

    strSqlFrom = " ((((" & v_TRusuarios(idIfocUsuario) & ") as trusuarios " & vbNewLine & _
                 " INNER JOIN (" & v_oferta_candidatos & ") as ocandidatos ON trusuarios.idPersona = ocandidatos.idPersona)" & vbNewLine & _
                 " INNER JOIN t_oferta ON ocandidatos.idOferta = t_oferta.id)" & vbNewLine & _
                 " INNER JOIN a_ofertaestado ON t_oferta.fkOfertaEstado = a_ofertaestado.id)"

    strSqlWhere = ""
    strSqlGroupBy = ""
    strSqlHaving = ""
    strSqlOrder = ""

    v_oferta_candidatosTRtmppruebas = montarSQL(strSqlSelect, _
                                       strSqlFrom, _
                                       strSqlWhere, _
                                       strSqlGroupBy, _
                                       strSqlHaving, _
                                       strSqlOrder)
End Function

Public Static Sub RewriteQuerySQL(strQueryName As String, strSql As String)
   Dim db As dao.database
   Dim qdf As dao.QueryDef
   Set db = CurrentDb()
   Set qdf = db.QueryDefs(strQueryName)

   qdf.sql = strSql

   '"SELECT [Table].Field1, [Table].Field2 " & vbCrLf & _
    "FROM [Table] " & vbCrLf & _
    "WHERE ([Table].Field1 = " & Chr(34) & strParameter & Chr(34) & ");"
End Sub
