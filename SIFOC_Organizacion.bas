Attribute VB_Name = "SIFOC_Organizacion"
Option Explicit
Option Compare Database

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  08/10/2009
'   Name:   getSqlListadoCitas
'   Desc:   devuelve un string con el sql del listado de organizaciones que cumplen los
'           criterios que se pasan por parámetro
'   Param:
'           idOrganizacion(long), identificador de Organizacion(opcional)
'           nombre(string), nombre de la empresa o parte de el (optional)
'           cif(string), cif o parte de el(opcional)
'           razon(string), razon social o parte de ella(opcional)
'           idServicio(long), identificador de servicio o parte de el(optional)
'           order(integer), order by nombre (1=ASC, 2=DESC, otro=no orden)
'   Retur:  sql(string), listado de gestiones solicitadas
'-------------------------------------------------------------------------------------
Public Function getSqlListadoOrganizaciones(Optional idOrganizacion As Long = 0, _
                                            Optional nombre As String = "", _
                                            Optional cif As String = "", _
                                            Optional razon As String, _
                                            Optional idServicio As Long = 0, _
                                            Optional orderNombre As Integer = 0, _
                                            Optional email As String = "", _
                                            Optional telefono As String, _
                                            Optional cerrada As Boolean = False) As String
    Dim strSql As String
    Dim strselect As String
    Dim strFrom As String
    Dim strWhere As String
    Dim strOrder As String
    
    strselect = "t_organizacion.id , t_Organizacion.cif, t_Organizacion.nombre AS NombreEmpresa, t_organizacion.razonSocial, t_organizacion.email, IIf([fechaCierre] Is Null Or [fechaCierre]>Now(),'Alta','Baja') AS Estado"
    strFrom = "t_organizacion"
    strWhere = ""
    strOrder = ""
    
    'Tratamos condiciones del WHERE
    If (idOrganizacion <> 0) Then 'id
        strWhere = addConditionWhere(strWhere, "(t_organizacion.id=" & idOrganizacion & ")")
    End If
    If (nombre <> "") Then
        strWhere = addConditionWhere(strWhere, "(UCase([nombre]) Like '*" & UCase(nombre) & "*')")
    End If
    If (cif <> "") Then
        strWhere = addConditionWhere(strWhere, "([cif] like '*" & UCase(cif) & "*')")
    End If
    If (razon <> "") Then
        strWhere = addConditionWhere(strWhere, "(UCase([razonSocial]) Like '*" & UCase(razon) & "*')")
    End If
    If (email <> "") Then
        strWhere = addConditionWhere(strWhere, "([email] Like '*" & email & "*')")
    End If
    If (telefono <> "") Then
        strWhere = addConditionWhere(strWhere, "([telefono] Like '*" & telefono & "*')")
    End If
    If (cerrada = True) Then 'Empresa cerrada
        strWhere = addConditionWhere(strWhere, "t_organizacion.fechaCierre> now() OR t_organizacion.fechaCierre Is Null")
    Else
        strWhere = addConditionWhere(strWhere, "t_organizacion.fechaCierre< now()")
    End If
    
    If (idServicio <> 0) Then
        strFrom = strFrom & " LEFT JOIN r_serviciousuario ON t_organizacion.id = r_serviciousuario.fkOrganizacion"
        strWhere = addConditionWhere(strWhere, "r_serviciousuario.fkServicio = " & idServicio)
        strWhere = addConditionWhere(strWhere, "r_serviciousuario.fechaInicio <= now() AND (r_serviciousuario.fechaFin >= now() OR r_serviciousuario.fechaFin is Null)")
    End If
    
    strSql = montarSQL(strselect, _
                       strFrom, _
                       strWhere, _
                       , _
                       , _
                       strOrder)
Debug.Print strSql
    getSqlListadoOrganizaciones = strSql
End Function

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  27/08/2014 - Act:-
'   Name:   altaServicioOrganizacion
'   Desc:   Crea alta servicio empresa
'   Param:  idServicio As Long
'           fechaI
'           idIfocUsuario
'           FECHAF
'           idMotivoBaja
'           mejora
'           observacion
'           idOrganizacion
'   Retur:  boolean
'-------------------------------------------------------------------------------------
Public Function altaServicioOrganizacion(idOrganizacion As Long, _
                                    idServicio As Long, _
                                    fechaI As Date, _
                                    idIfocUsuario As Long, _
                                    Optional FECHAF As Date = "01/01/1900 00:00:00", _
                                    Optional idMotivoBaja As Long = 0, _
                                    Optional mejora As Integer = 0, _
                                    Optional observacion As String = "") As Integer
    altaServicioOrganizacion = altaServicioUsuario(idServicio, fechaI, idIfocUsuario, FECHAF, idMotivoBaja, mejora, observacion, , idOrganizacion)
End Function

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  27/08/2014 - Act:-
'   Name:   checkServiceIntermediacionOrganizacion
'   Desc:   Creamos alta en servicio Intermediación con fecha fin 31/12/año cuando
'           la empresa pone una oferta en caso de que no esté de alta.
'   Param:  idServicio As Long
'           fechaI
'           idIfocUsuario
'           FECHAF
'           idMotivoBaja
'           mejora
'           observacion
'           idOrganizacion
'   Retur:  boolean
'-------------------------------------------------------------------------------------
Public Function checkServiceIntermediacionOrganizacion(idOrganizacion As Long, _
                                                       fecha As Date)
    '29 - Servicio intermediación empresas
    Dim idServicio As Long
    '19 - Caducidad de la demanda
    Dim idMotivoBaja As Long
    Dim FECHAF As Date
    Dim observacion As String
    
    idServicio = 29
    idMotivoBaja = 19
    FECHAF = "31/12/" & Year(fecha)
    observacion = "Alta automática (p.ej. Al crear oferta)"
    
    If Not isUserServiceActive(fecha, idServicio, , idOrganizacion) Then
        altaServicioOrganizacion idOrganizacion, idServicio, fecha, U_idIfocUsuarioActivo, FECHAF, idMotivoBaja, , observacion
        MsgBox "Se da de alta la empresa en el servicio Intermediación hasta 31/12/" & Year(fecha), vbOKOnly, "Alert: SIFOC_Organizacion"
    Else
        MsgBox "Esta organización se encuentra de alta en servicio Intermediación", vbOKOnly, "Alert: SIFOC_Organizacion"
    End If
End Function

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  12/07/2010
'   Name:   getOrganizacionNombre
'   Desc:   devuelve un string del nombre de empresa de la organizacion
'           que se pasa por parámetro
'   Param:  idOrganizacion(long), identificador de Organizacion
'   Retur:  nombre(string), nombre de la empresa
'-------------------------------------------------------------------------------------
Public Function getOrganizacionNombre(idOrganizacion As Long)
    getOrganizacionNombre = DLookup("[nombre]", "t_organizacion", "[id]=" & idOrganizacion)
End Function

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  12/07/2010
'   Name:   getOrganizacionIdTipoEmpresa
'   Desc:   devuelve el identificador de tipo de empresa de la organización
'           que se pasa por parámetro
'   Param:  idOrganizacion(long), identificador de Organizacion
'   Retur:  nombre(string), nombre de la empresa
'-------------------------------------------------------------------------------------
Public Function getOrganizacionIdTipoEmpresa(idOrganizacion As Long)
    getOrganizacionIdTipoEmpresa = DLookup("[fkTipoEmpresa]", "t_organizacion", "[id]=" & idOrganizacion)
End Function

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  12/07/2010 Actualización: 12/07/2013
'   Name:   getOrganizacionIdActividadEmpresarial
'   Desc:   devuelve el id cnae principal de la organización
'           que se pasa por parámetro
'   Param:  idOrganizacion(long), identificador de Organizacion
'   Retur:  nombre(string), nombre de la empresa
'-------------------------------------------------------------------------------------
Public Function getOrganizacionIdActividadEmpresarial(idOrganizacion As Long)
    getOrganizacionIdActividadEmpresarial = Nz(DLookup("[fkCnae2009]", "t_actividadeconomicaorganizacion", "[fkOrganizacion]=" & idOrganizacion & " AND principal = -1"), 0)
End Function

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  12/07/2013 Actualización: 12/07/2013
'   Name:   getOrganizacionActividadEmpresarial
'   Desc:   devuelve el codigo y título cnae principal de la organización
'           que se pasa por parámetro
'   Param:  idOrganizacion(long), identificador de Organizacion
'   Retur:  nombre(string), nombre de la empresa
'-------------------------------------------------------------------------------------
Public Function getOrganizacionActividadEmpresarial(idOrganizacion As Long)
    Dim idCnae As Long
    
    idCnae = getOrganizacionIdActividadEmpresarial(idOrganizacion)
    getOrganizacionActividadEmpresarial = Nz(DLookup("[codigo]", "a_cnae2009", "[id]=" & idCnae) & " - " & DLookup("[titulo]", "a_cnae2009", "[id]=" & idCnae), 0)
End Function

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  12/07/2010
'   Name:   getOrganizacionIdMunicipio
'   Desc:   devuelve el identificador de tipo de empresa de la organización
'           que se pasa por parámetro
'   Param:  idOrganizacion(long), identificador de Organizacion
'   Retur:  nombre(string), nombre de la empresa
'-------------------------------------------------------------------------------------
Public Function getOrganizacionIdMunicipio(idOrganizacion As Long)
    getOrganizacionIdMunicipio = DLookup("[fkMunicipio]", "t_organizacion", "[id]=" & idOrganizacion)
End Function

'--------------------------------------------------------------------------------------
Public Function tmpnewAltasServicioEmpresa()
    Dim idOrg As Long
    Dim str As String
    Dim strSql As String
    Dim counter As Integer
    Dim counter2 As Integer
    Dim idServicio As Long
    Dim ids As String
    Dim ids2 As String
    Dim fi As Date
    Dim ff As Date
    
    fi = DateSerial("2013", "01", "01")
    ff = DateSerial("2013", "12", "31") & " 23:59"
    idServicio = 29 'servicio intermedia organizacion

    'Sólo miramos los carnes profesionales relacionados con Tráfico (idCarneProfesionalArea=2)
    strSql = " SELECT t_oferta.fkOrganizacion AS idOrganizacion, Min(t_oferta.fechaOferta) AS minfecha" & _
             " FROM t_oferta" & _
             " WHERE (((t_oferta.fechaOferta) Between #" & Format(fi, "mm/dd/yyyy 23:59") & "# And #" & Format(ff, "mm/dd/yyyy 23:59") & "#))" & _
             " GROUP BY t_oferta.fkOrganizacion"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    
    If Not rs.EOF Then
        rs.MoveFirst
    End If
    counter = 0
    ids = ""
    While Not rs.EOF
        If Not isUserServiceActiveInInterval(fi, ff, idServicio, , rs!idOrganizacion) Then
            'altaServicioOrganizacion rs!idOrganizacion, idServicio, rs!minfecha, 14, DateSerial(Year(ff), "12", "31"), 19
            counter = counter + 1
            ids = ids & ", " & rs!idOrganizacion
        Else
            ids2 = ids2 & ", " & rs!idOrganizacion
            counter2 = counter2 + 1
        End If
        rs.MoveNext
    Wend
    Debug.Print "Numero altas servicio: " & counter & vbNewLine & "ids: " & ids & vbNewLine & "Numero YA en servicio: " & counter2 & vbNewLine & "ids2: " & ids2
End Function
