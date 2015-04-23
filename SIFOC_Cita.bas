Attribute VB_Name = "SIFOC_Cita"
Option Explicit
Option Compare Database

Private strSql As String
Private strselect As String
Private strFrom As String
Private strWhere As String
Private strGroup As String
Private strHaving As String
Private strOrder As String

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/01/2014 - Actualización:  21/01/2014
'   Name:   SIFOC_Cita
'   Desc:   Cita Module
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------

'------------Getters------------------------
Public Static Function getStrSelect() As String
    getStrSelect = strselect
End Function

Public Static Function getStrFrom() As String
    getStrFrom = strFrom
End Function

Public Static Function getStrWhere() As String
    getStrWhere = strWhere
End Function

Public Static Function getStrGroup() As String
    getStrGroup = strGroup
End Function

Public Static Function getStrOrder() As String
    getStrOrder = strOrder
End Function

'------------Setters--------------------------
Public Static Function setStrSql(sSql As String) As String
    strSql = sSql
    setStrSql = strSql
End Function

Public Static Function setStrSelect(sSelect As String) As String
    strselect = sSelect
    setStrSelect = strselect
End Function

Public Static Function setStrFrom(sFrom As String) As String
    strFrom = sFrom
    setStrFrom = strFrom
End Function

Public Static Function setStrWhere(sWhere As String) As String
    strWhere = sWhere
    setStrWhere = strWhere
End Function

Public Static Function setStrGroup(sGroup As String) As String
    strGroup = sGroup
    setStrGroup = sGroup
End Function

Public Static Function setStrOrder(sOrder As String) As String
    strOrder = sOrder
    setStrOrder = strOrder
End Function

'------------Adders--------------------------
Public Static Function addStrSql(sSql As String) As String
    strSql = addConditionWhere(strSql, sSql)
    addStrSql = strSql
End Function

Public Static Function addStrSelect(sSelect As String) As String
    strselect = addConditionWhere(strselect, sSelect)
    addStrSelect = strselect
End Function

Public Static Function addStrFrom(sFrom As String) As String
    strFrom = addConditionWhere(strFrom, sFrom)
    addStrFrom = strFrom
End Function

Public Static Function addStrWhere(sWhere As String) As String
    strWhere = addConditionWhere(strWhere, sWhere)
    addStrWhere = strWhere
End Function

Public Static Function addStrGroup(sGroup As String) As String
    addStrGroup = addConditionWhere(strGroup, sGroup)
End Function

Public Static Function addStrOrder(sOrder As String) As String
    strOrder = addConditionWhere(strOrder, sOrder)
    addStrOrder = strOrder
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/12/2013 - Actualización:  24/02/2014
'   Name:   initSql
'   Desc:   Init sql Citas base
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSql()
'  id,
'  fechaDemanda,
'  fecha,
'  fkIfocAmbito,
'  fkCitaSesion,
'  fkGestionTipo,
'  fkIfocusuarioCit,
'  duracion,
'  fkIfocUsuarioTec,
'  obs,
'  fkServicio,
'  cancelada,
'  fkOferta,
'  fkCurso,
'  fkServicioSubtipo,
'  timestamp

    Dim sqlEmp As String
    Dim sqlPer As String
    Dim sqlPry As String
  
    sqlEmp = " SELECT r_citausuario.fkCita, Count(r_citausuario.fkPersona) AS numOrganizaciones" & _
              " FROM r_citausuario" & _
              " WHERE ((Not (r_citausuario.fkPersona) Is Null))" & _
              " GROUP BY r_citausuario.fkCita"
              
    sqlPer = " SELECT r_citausuario.fkCita, Count(r_citausuario.fkOrganizacion) AS numPersonas" & _
              " FROM r_citausuario" & _
              " WHERE ((Not (r_citausuario.fkOrganizacion) Is Null))" & _
              " GROUP BY r_citausuario.fkCita"
              
    sqlPry = " SELECT r_citausuario.fkCita, Count(r_citausuario.fkProyectoEmprendedor) AS numProyectos" & _
              " FROM r_citausuario" & _
              " WHERE ((Not (r_citausuario.fkProyectoEmprendedor) Is Null))" & _
              " GROUP BY r_citausuario.fkCita"
    
    strselect = "t_cita.id, fecha, sesion, tipo, a_servicio.aka as servicio, subtipo, uifocc.aka as citador, uifoct.aka as tecnico, duracion, numPersonas, numOrganizaciones, numProyectos, fechaDemanda, observacion, acude, cancelada, fkOferta, idCurso"
    strFrom = "((((((((((t_cita" & _
              " LEFT JOIN (" & sqlEmp & ")as emp ON t_cita.id = emp.fkCita)" & _
              " LEFT JOIN (" & sqlPer & ")as per ON t_cita.id = per.fkCita)" & _
              " LEFT JOIN (" & sqlPry & ")as pry ON t_cita.id = pry.fkCita)" & _
              " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN a_servicio ON t_cita.fkServicio = a_servicio.id)" & _
              " LEFT JOIN a_serviciosubtipo ON t_cita.fkServicioSubtipo = a_serviciosubtipo.id)" & _
              " LEFT JOIN a_ifocambito ON t_cita.fkIfocAmbito = a_ifocambito.id)" & _
              " LEFT JOIN a_citasesion ON t_cita.fkCitaSesion = a_citasesion.id)" & _
              " LEFT JOIN t_ifocusuario as uifocc ON t_cita.fkIfocusuarioCit = uifocc.fkPersona)" & _
              " LEFT JOIN t_ifocusuario as uifoct ON t_cita.fkIfocusuarioTec = uifoct.fkPersona)"
    
    'Estado abierto y pendiente
    strWhere = ""
    strHaving = ""
    strGroup = "t_cita.id, fechaDemanda, fecha, ambito, sesion, tipo, a_servicio.aka, subtipo, uifocc.aka, uifoct.aka, duracion, observacion, acude, cancelada, fkOferta, idCurso, numPersonas, numOrganizaciones, numProyectos"
    strOrder = "fecha DESC"
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/12/2013 - Actualización:  24/02/2014
'   Name:   initSqlClean
'   Desc:   Init sql Citas base
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSqlClean()
    
    strselect = ""
    strFrom = ""
    
    strWhere = ""
    strHaving = ""
    strGroup = ""
    strOrder = ""
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/02/2014 - Actualización:  24/02/2014
'   Name:   initSqlCitaPersonas
'   Desc:   Listado
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSqlCitaPersonas()
'idPersona
't_cita.fecha
'a_ifocambito.ambito
'a_citasesion.sesion
'a_gestiontipo.tipo
'a_servicio.servicio
'a_serviciosubtipo.subtipo
'uifocc.aka AS citador
'uifoct.aka AS tecnico
't_cita.duracion
't_cita.fechaDemanda
't_cita.observacion
't_cita.acude
't_cita.cancelada
't_cita.fkOferta
't_cita.idCurso
  
    strselect = "r_citausuario.fkPersona AS idPersona, v_datospersonales.name, Count(t_cita.id) as numCitas, Sum(duracion) as Minutos, edad, fechaNacimiento, intervalo, dni, sexo, empleabilidad, nacionalidad, poblacion"
    strFrom = "(r_citausuario" & _
              " LEFT JOIN (((((((t_cita" & _
              " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN a_servicio ON t_cita.fkServicio = a_servicio.id)" & _
              " LEFT JOIN a_serviciosubtipo ON t_cita.fkServicioSubtipo = a_serviciosubtipo.id)" & _
              " LEFT JOIN a_ifocambito ON t_cita.fkIfocAmbito = a_ifocambito.id)" & _
              " LEFT JOIN a_citasesion ON t_cita.fkCitaSesion = a_citasesion.id)" & _
              " LEFT JOIN t_ifocusuario as uifocc ON t_cita.fkIfocusuarioCit = uifocc.fkPersona)" & _
              " LEFT JOIN t_ifocusuario as uifoct ON t_cita.fkIfocusuarioTec = uifoct.fkPersona) ON r_citausuario.fkCita = t_cita.id)" & _
              " LEFT JOIN v_datospersonales ON r_citausuario.fkPersona = v_datospersonales.id"
    
    'Estado abierto y pendiente
    strWhere = "(Not (r_citausuario.fkPersona) Is Null)"
    strHaving = ""
    strGroup = "r_citausuario.fkPersona, v_datospersonales.name, edad, fechaNacimiento, intervalo, dni, sexo, empleabilidad, nacionalidad, poblacion"
    strOrder = "r_citausuario.fkPersona ASC"
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/02/2014 - Actualización:  24/02/2014
'   Name:   initSqlCitaOrganizaciones
'   Desc:   Listado
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSqlCitaOrganizaciones()
'idPersona
't_cita.fecha
'a_ifocambito.ambito
'a_citasesion.sesion
'a_gestiontipo.tipo
'a_servicio.servicio
'a_serviciosubtipo.subtipo
'uifocc.aka AS citador
'uifoct.aka AS tecnico
't_cita.duracion
't_cita.fechaDemanda
't_cita.observacion
't_cita.acude
't_cita.cancelada
't_cita.fkOferta
't_cita.idCurso

    strselect = "r_citausuario.fkOrganizacion AS idOrganizacion, t_organizacion.nombre, Count(t_cita.id) as numCitas, Sum(duracion) as Minutos, razonSocial, cif, codigoPostal, web, email, telefono, fax, fechaInscripcion, fechaCreacion, fechaCierre"
    strFrom = "(r_citausuario" & _
              " LEFT JOIN (((((((t_cita" & _
              " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN a_servicio ON t_cita.fkServicio = a_servicio.id)" & _
              " LEFT JOIN a_serviciosubtipo ON t_cita.fkServicioSubtipo = a_serviciosubtipo.id)" & _
              " LEFT JOIN a_ifocambito ON t_cita.fkIfocAmbito = a_ifocambito.id)" & _
              " LEFT JOIN a_citasesion ON t_cita.fkCitaSesion = a_citasesion.id)" & _
              " LEFT JOIN t_ifocusuario as uifocc ON t_cita.fkIfocusuarioCit = uifocc.fkPersona)" & _
              " LEFT JOIN t_ifocusuario as uifoct ON t_cita.fkIfocusuarioTec = uifoct.fkPersona) ON r_citausuario.fkCita = t_cita.id)" & _
              " LEFT JOIN t_organizacion ON r_citausuario.fkOrganizacion = t_organizacion.id"
    
    'Estado abierto y pendiente
    strWhere = "(Not (r_citausuario.fkOrganizacion) Is Null)"
    strHaving = ""
    strGroup = "r_citausuario.fkOrganizacion, t_organizacion.nombre, razonSocial, cif, codigoPostal, web, email, telefono, fax, fechaInscripcion, fechaCreacion, fechaCierre"
    strOrder = "r_citausuario.fkOrganizacion ASC"
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/02/2014 - Actualización:  24/02/2014
'   Name:   initSqlCitaProyectos
'   Desc:   Listado
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSqlCitaProyectos()
'idPersona
't_cita.fecha
'a_ifocambito.ambito
'a_citasesion.sesion
'a_gestiontipo.tipo
'a_servicio.servicio
'a_serviciosubtipo.subtipo
'uifocc.aka AS citador
'uifoct.aka AS tecnico
't_cita.duracion
't_cita.fechaDemanda
't_cita.observacion
't_cita.acude
't_cita.cancelada
't_cita.fkOferta
't_cita.idCurso
    
    strselect = "r_citausuario.fkProyectoEmprendedor AS idProyecto, nombreProy, Count(t_cita.id) as numCitas, Sum(duracion) as Minutos, fechaInicioProy, fechaFinProy, t_proyectoemprendedor.fkOrganizacion"
    strFrom = "(r_citausuario" & _
              " LEFT JOIN (((((((t_cita" & _
              " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN a_servicio ON t_cita.fkServicio = a_servicio.id)" & _
              " LEFT JOIN a_serviciosubtipo ON t_cita.fkServicioSubtipo = a_serviciosubtipo.id)" & _
              " LEFT JOIN a_ifocambito ON t_cita.fkIfocAmbito = a_ifocambito.id)" & _
              " LEFT JOIN a_citasesion ON t_cita.fkCitaSesion = a_citasesion.id)" & _
              " LEFT JOIN t_ifocusuario as uifocc ON t_cita.fkIfocusuarioCit = uifocc.fkPersona)" & _
              " LEFT JOIN t_ifocusuario as uifoct ON t_cita.fkIfocusuarioTec = uifoct.fkPersona) ON r_citausuario.fkCita = t_cita.id)" & _
              " LEFT JOIN t_proyectoemprendedor ON r_citausuario.fkProyectoEmprendedor = t_proyectoemprendedor.id"
    
    'Estado abierto y pendiente
    strWhere = "(Not (r_citausuario.fkProyectoEmprendedor) Is Null)"
    strHaving = ""
    strGroup = "r_citausuario.fkProyectoEmprendedor, nombreProy, fechaInicioProy, fechaFinProy, t_proyectoemprendedor.fkOrganizacion"
    strOrder = "r_citausuario.fkProyectoEmprendedor ASC"
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/02/2014 - Actualización:  24/02/2014
'   Name:   initSqlCurso
'   Desc:   Listado
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSqlCurso()
'idPersona
't_cita.fecha
'a_ifocambito.ambito
'a_citasesion.sesion
'a_gestiontipo.tipo
'a_servicio.servicio
'a_serviciosubtipo.subtipo
'uifocc.aka AS citador
'uifoct.aka AS tecnico
't_cita.duracion
't_cita.fechaDemanda
't_cita.observacion
't_cita.acude
't_cita.cancelada
't_cita.fkOferta
't_cita.idCurso
  
    
    strselect = "t_cita.idCurso AS idCurso, t_curso.nombre, t_cita.fecha, a_ifocambito.ambito, a_citasesion.sesion, a_gestiontipo.tipo, a_servicio.servicio, a_serviciosubtipo.subtipo, uifocc.aka AS citador, uifoct.aka AS tecnico, t_cita.duracion, t_cita.fechaDemanda, t_cita.observacion, t_cita.acude, t_cita.cancelada, t_cita.fkOferta"
    strFrom = " (((((((t_cita" & _
              " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN a_servicio ON t_cita.fkServicio = a_servicio.id)" & _
              " LEFT JOIN a_serviciosubtipo ON t_cita.fkServicioSubtipo = a_serviciosubtipo.id)" & _
              " LEFT JOIN a_ifocambito ON t_cita.fkIfocAmbito = a_ifocambito.id)" & _
              " LEFT JOIN a_citasesion ON t_cita.fkCitaSesion = a_citasesion.id)" & _
              " LEFT JOIN t_ifocusuario as uifocc ON t_cita.fkIfocusuarioCit = uifocc.fkPersona)" & _
              " LEFT JOIN t_ifocusuario as uifoct ON t_cita.fkIfocusuarioTec = uifoct.fkPersona)" & _
              " LEFT JOIN t_curso ON t_cita.idCurso = t_curso.id"
    
    'Estado abierto y pendiente
    strWhere = "(Not (t_cita.idCurso) Is Null)"
    strHaving = ""
    strGroup = "t_cita.idCurso, t_curso.nombre, t_cita.fecha, a_ifocambito.ambito, a_citasesion.sesion, a_gestiontipo.tipo, a_servicio.servicio, a_serviciosubtipo.subtipo, uifocc.aka, uifoct.aka, t_cita.duracion, t_cita.fechaDemanda, t_cita.observacion, t_cita.acude, t_cita.cancelada, t_cita.fkOferta"
    strOrder = "t_cita.fecha DESC"
End Function

Public Static Function getQuery() As String
    strSql = montarSQL(strselect, _
                       strFrom, _
                       strWhere, _
                       strGroup, _
                       strHaving, _
                       strOrder)
    getQuery = strSql
End Function

Public Function tmp()
    initSql
    Debug.Print getQuery
End Function

'createLocalTable
'----------------------------------------------------------------------------------------------
'   Name:   CreaCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  29/09/2009   Actualización: 17/03/2010 Asunción Huertas
'   Desc:   Creamos nueva cita en tabla t_cita (SIN PERSONA al hacer las citas grupales)
'   Param:  ifocAmbito(*integer),
'           CitaSesion(*integer),
'           fecha(*date),
'           idIfocUsuarioCit(*long),
'           idIfocUsuarioTec(*long),
'           obs(string),
'           idServicio(long),
'           idOrganizacion(long),
'           idOferta(long),
'           idAutoempleoProyecto(long),
'           idCurso(long))
'   Return:  0 -> Ok
'           -1 -> Ko
'-----------------------------------------------------------------------------------------------
Private Function creaCita(idGestionTipo As Integer, _
                         idIfocAmbito As Integer, _
                         idCitaSesion As Integer, _
                         fecha As Date, _
                         idIfocUsuarioCit As Long, _
                         idIfocUsuarioTec As Long, _
                         Optional OBS As String = "", _
                         Optional idServicio As Long = 0, _
                         Optional idOferta As Long = 0, _
                         Optional idCurso As Long = 0) As Integer

    Dim strCita As String
    Dim sqlCita As String
    Dim sqlInsert As String
    Dim sqlValues As String
    
On Error GoTo TratarError

    strCita = filterSQL(OBS)
    'isMissing(argumento)
    'Tabla t_gestion (campos)
    'fkPersona*,fkifocambito*,fecha*,gestion*,fkIfocUsuario*,fkFormaContacto,fkMenuGestion,fkOferta
    
    'fecha en insert formato mm/dd/yyyy
    
    sqlInsert = " INSERT INTO t_cita (" & _
                    " fkGestionTipo" & _
                    ", fkIfocAmbito" & _
                    ", fkCitaSesion" & _
                    ", fecha " & _
                    ", fechaDemanda" & _
                    ", fkIfocUsuarioCit" & _
                    ", fkIfocUsuarioTec" & _
                    IIf(OBS = "", "", ", observacion ") & _
                    IIf(idServicio = 0, "", ", fkServicio ") & _
                    IIf(idOferta = 0, "", ", fkOferta ") & _
                    IIf(idCurso = 0, "", ", idCurso ") & _
                    ")"
    sqlValues = " VALUES (" & _
                    idGestionTipo & " " & _
                    ", " & idIfocAmbito & " " & _
                    ", " & idCitaSesion & " " & _
                    ", '" & Format(fecha, "dd/mm/yyyy hh:nn:ss") & "' " & _
                    ", '" & now() & "' " & _
                    ", " & idIfocUsuarioCit & " " & _
                    ", " & idIfocUsuarioTec & " " & _
                    IIf(OBS = "", "", ", '" & strCita & "'") & _
                    IIf(idServicio = 0, "", ", " & idServicio & " ") & _
                    IIf(idOferta = 0, "", ", " & idOferta & " ") & _
                    IIf(idCurso = 0, "", ", " & idCurso & " ") & _
                    ")"

    sqlCita = sqlInsert & sqlValues
Debug.Print sqlCita
    
    CurrentDb.Execute sqlCita
    
    creaCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    creaCita = -1
End Function

'----------------------------------------------------------------------------------------------------------
'   Name:   creaCitaGrupal
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/09/2009    Actualización: 17/03/2010 Asunción Huertas
'   Desc:   Creamos nueva cita en tabla t_Cita y añadimos grupo personas
'   Return:  >0 -> identificador de la cita
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Public Function creaCitaGrupal(idGestionTipo As Integer, _
                               idIfocAmbito As Integer, _
                               idCitaSesion As Integer, _
                               fecha As Date, _
                               idIfocUsuarioCit As Long, _
                               idIfocUsuarioTec As Long, _
                               Optional Personas As String = "", _
                               Optional OBS As String = "", _
                               Optional idServicio As Long = 0, _
                               Optional Organizaciones As String = "", _
                               Optional idOferta As Long = 0, _
                               Optional ProyEmprendedores As String = "", _
                               Optional idCurso As Long = 0) As Long
On Error GoTo TratarError
    
    Dim strCita As String
    'Dim fechaHora As Date
        
    'strCita = filterSQL(OBS)'filtramos en crear cita
    
    If creaCita(idGestionTipo, _
                idIfocAmbito, _
                idCitaSesion, _
                fecha, _
                idIfocUsuarioCit, _
                idIfocUsuarioTec, _
                OBS, _
                idServicio, _
                idOferta, _
                idCurso) = -1 Then
        
        creaCitaGrupal = -1
        Exit Function
    End If
    
    Dim idCita As Long
    
    idCita = getIdCita(idIfocAmbito, idCitaSesion, fecha, idIfocUsuarioCit, idIfocUsuarioTec)

    'Añadimos persona/s
    If Len(Personas) > 0 Then
        If insPersonasCita(idCita, Personas, idIfocUsuarioTec) = -1 Then
            GoTo TratarError
        End If
    End If
    
    'Añadimos organizacion/es
    If Len(Organizaciones) > 0 Then
        If insOrganizacionesCita(idCita, Organizaciones, idIfocUsuarioTec) = -1 Then
            GoTo TratarError
        End If
    End If
    
    'Añadimos proyEmprendedor/es
    If Len(ProyEmprendedores) > 0 Then
        If insProyEmprendedoresCita(idCita, ProyEmprendedores, idIfocUsuarioTec) = -1 Then
            GoTo TratarError
        End If
    End If
    
    creaCitaGrupal = idCita
SalirTratarError:
    Exit Function
TratarError:
    creaCitaGrupal = -1
End Function

'----------------------------------------------------------------------------------------------------------
'   Name:   getIdCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/09/2009   Actualización: 17/03/2010 Asunción Huertas
'   Desc:   Obtenemos el id de cita de la cita con los datos pasados por parámetro
'           PARA OBTENER ID DE CITA GRUPAL
'   Param:  idIfocAmbito(*integer) identificador de tipo de cita,
'           idCitaSesion(*integer) identificador de tipo de sesion,
'           fecha(*),
'           gestion(*),
'           idIfocUsuarioCit(*long),
'           idIfocUsuarioTec(*long))
'   Return: id -> Cita que cocuerda con los parametros pasados
'            0 -> no hay cita que concuerde
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Private Function getIdCita(idIfocAmbito As Integer, _
                          idCitaSesion As Integer, _
                          fecha As Date, _
                          idIfocUsuarioCit As Long, _
                          idIfocUsuarioTec) As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim strGestion As String
    Dim result As Long
On Error GoTo TratarError:
    sql = " SELECT id" & _
          " FROM t_cita" & _
          " WHERE fkIfocAmbito=" & idIfocAmbito & _
          "  AND fkCitaSesion=" & idCitaSesion & _
          "  AND fecha =#" & Format(fecha, "mm/dd/yyyy hh:nn:ss") & "#" & _
          "  AND fkIfocUsuarioCit=" & idIfocUsuarioTec & _
          "  AND fkIfocUsuarioCit=" & idIfocUsuarioCit & ";"
    
    Set rs = New ADODB.Recordset
    result = 0
'debugando sql
    
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF And rs.RecordCount = 1 Then
        rs.MoveFirst
        result = rs!id
    Else
        result = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
SalirTratarError:
    getIdCita = result
    Exit Function
TratarError:
    getIdCita = -1
End Function

'---------------------------------------------------------------------------
'   Name:   esCitaModificable
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009 Actualizacion:
'   Desc:   Nos dice si una cita se pude modificar
'           (sólo modificable hasta dia después inclusive de la fecha de cita
'           para todas las personas, excepto el tecnico que realiza la cita que
'           la puede modificar siempre)
'   Param:  fechaCita
'   Return: true -> cita modificable
'           false ->cita no modificable
'---------------------------------------------------------------------------
Public Function esCitaModificable(fecha As Date, _
                                  idIfocUsuario As Long) As Boolean
    If (fecha > Date - 2) Or (idIfocUsuario = usuarioIFOC()) Then
        esCitaModificable = True
    Else
        esCitaModificable = False
    End If
End Function

'----------------------------------------------------------------------------------------------------------
'   Name:   esCitaEliminable
'   Autor:  Jose Manuel Sanchez
'   Fecha:  25/09/2009
'   Desc:   Devuelve si la gestión la puede modificar el usuario ifoc o no
'   Param:  idGestion(*long)    identificador gestion a eliminar
'           idIfocUsuario(*long) identificador de usuario ifoc
'   Return: TRUE -> Puede ser eliminada por ese usuario
'           FALSE -> No puede ser eliminada por ese usuario
'----------------------------------------------------------------------------------------------------------
Public Function esCitaEliminable(idCita As Long, _
                                 idIfocUsuario As Long) As Boolean
    Dim idIfocUsuarioCita As Long
    Dim fechaEntradaCita As Date
    Dim idIfocUsuarioCitador As Long
On Error GoTo TratarError
    
    idIfocUsuarioCita = DLookup("fkIfocUsuarioTec", "t_cita", "[id]=" & idCita)
    
    If (idIfocUsuario = idIfocUsuarioCita) Then
        esCitaEliminable = True
    Else
        idIfocUsuarioCita = DLookup("fkIfocUsuarioCit", "t_cita", "[id]=" & idCita)
        fechaEntradaCita = DLookup("fechaDemanda", "t_cita", "[id]=" & idCita)
        If (fechaEntradaCita >= Date And idIfocUsuarioCita = idIfocUsuario) Then
            esCitaEliminable = True
        Else
            esCitaEliminable = False
        End If
    End If
    
SalirTratarError:
    Exit Function
TratarError:
    esCitaEliminable = False
End Function

'---------------------------------------------------------------------------
'   Name:   insPersonasCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  29/9/2009
'   Desc:   Inserta la/s personas pasada por parámetro string
'   Param:  idCita(long), identificador de cita
'           Personas(string), identificadores de personas separados por ','
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
'Private
Public Function insPersonasCita(idCita As Long, _
                                Personas As String, _
                                idIfocUsuario As Long) As Integer
    Dim sql As String
    Dim args As Variant
    Dim idPersona As Long
    Dim numPersonas As Integer
    Dim i As Integer
On Error GoTo TratarError
    
    numPersonas = countSubStrings(Personas, ",")
    args = Split(Personas, ",")
        
    For i = 0 To numPersonas - 1
        idPersona = args(i)
        If insPersonaCita(idCita, idPersona, idIfocUsuario) = -1 Then
            GoTo TratarError
        End If
    Next
    
    insPersonasCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    insPersonasCita = -1
    debugando Err.description
End Function

'---------------------------------------------------------------------------
'   Name:   delPersonasDeCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/9/2009 Actualizacion:
'   Desc:   Elimina la/s persona/s asociadas a la cita(todas)
'   Param:  identificador de cita
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delPersonasDeCita(idCita As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkCita" & _
          " FROM r_citausuario" & _
          " WHERE fkCita=" & idCita & " AND fkPersona is NOT NULL;"
    
    CurrentDb.Execute sql
    
    delPersonasDeCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delPersonasDeCita = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Name:   insPersonaCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/09/2009 Actualizacion:
'   Desc:   Añade una persona a la cita
'   Param:  idCita, identificador de la cita donde añadimos la persona
'           idPersona, identificador de la persona donde añadimos la cita
'   Return:
'---------------------------------------------------------------------------
Public Function insPersonaCita(idCita As Long, _
                               idPersona As Long, _
                               idIfocUsuario As Long) As Integer
    On Error GoTo Error
    
    Dim str As String
    str = " INSERT INTO r_citausuario (fkCita, fkPersona, fkIfocUsuario)" & _
          " VALUES (" & idCita & ", " & idPersona & ", " & idIfocUsuario & ");"

'debugando str
    CurrentDb.Execute str
    
    insPersonaCita = 0
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    insPersonaCita = -1
End Function

'---------------------------------------------------------------------------
'   Name:   personasEnCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/09/2009 Actualizacion:
'   Desc:   Devuelve el numero de personas que hay en la cita
'   Param:  idCita, identificador de la cita donde añadimos la persona
'---------------------------------------------------------------------------
Public Function personasEnCita(idCita As Long) As Integer
On Error GoTo Error
    Dim rs As ADODB.Recordset
    Dim str As String
    
    str = " SELECT fkCita, fkPersona " & _
          " FROM r_citausuario " & _
          " WHERE fkCita = " & idCita & ";"

    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rs.EOF Then
        personasEnCita = 0
    Else
        personasEnCita = rs.RecordCount
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    personasEnCita = -1
End Function

'---------------------------------------------------------------------------
'   Name:   delPersonaCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  24/09/2009 Actualizacion:
'   Desc:   Elimina la persona pasada por parámetro asociada a la cita(todas)
'   Param:  idCita(long), identificador de cita
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delPersonaCita(idCita As Long, _
                               idPersona As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkCita" & _
          " FROM r_citausuario" & _
          " WHERE fkCita=" & idCita & " AND fkPersona= " & idPersona & ";"
    
    CurrentDb.Execute sql
    
    delPersonaCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delPersonaCita = -1
    debugando "Error: " & Err.description
End Function

'------------------------------------------------------------------------------------
'   Name:   getSqlListadoCitasUsuario
'   Autor:  Jose Manuel Sanchez
'   Fecha:  22/09/2009  Act: 18/02/2014
'   Desc:   devuelve un string con el sql del listado de citas que cumplen los
'           criterios que se pasan por parámetro
'   Param:  idCita(long), id de cita
'           fechaInicio(date), fecha inicio del listado de citas
'           fechaFin(date), fecha fin del listado de gestiones
'           idPersona(long), identificador de persona (opcional)
'           idOrganizacion(long), identificador de Organizacion(opcional)
'           idOferta(long), identificador de oferta(opcional)
'           idCurso(long), identificador de curso(opcional)
'           idIfocUsuario(long), identificador de usuario ifoc(optional)
'           order(integer), order by fecha(1=ASC, 2=DESC, otro=no orden)
'   Retur:  sql(string), listado de gestiones solicitadas
'-------------------------------------------------------------------------------------
Public Function getSqlListadoCitasUsuario(Optional idCita As Long = 0, _
                                          Optional fechaInicio As Date = "01/01/1900", _
                                          Optional fechaFin As Date, _
                                          Optional idPersona As Long = 0, _
                                          Optional idOrganizacion As Long = 0, _
                                          Optional idOferta As Long = 0, _
                                          Optional idIfocUsuario As Long = 0, _
                                          Optional order As Integer = 0) As String
    Dim strselect As String
    Dim strFrom As String
    Dim strWhere As String
    Dim strOrderBy As String
    
    Dim fechaI As Date
    Dim FECHAF As Date
    
    'Inicializamos variables
    strselect = " t_cita.id," & _
                " IIf([fkCitaSesion]=1,'Cita (I)','Cita (G)') AS IoG," & _
                " a_gestiontipo.tipo AS Tipo," & _
                " t_cita.fechaDemanda AS FechaAlta," & _
                " t_cita.fecha," & _
                " t_ifocusuario.aka AS Tecnico," & _
                " IIf(isnull([v_apellidosnombre].[name]),iif(isnull([t_organizacion].[nombre]),[t_proyectoemprendedor].[nombreProy],[t_organizacion].[nombre]),[v_apellidosnombre].[name]) AS Usuario," & _
                " t_ifocusuario1.aka AS Citador," & _
                " t_cita.observacion," & _
                " IIf([fkIfocAmbito]=2, r_citausuario.acudeNoacudeAnula, r_citausuario.acudeNoacudeAnula) AS acudeNoacudeAnula, " & _
                " t_cita.cancelada," & _
                " r_citausuario.fkPersona," & _
                " r_citausuario.fkOrganizacion," & _
                " t_cita.duracion"
    
    strFrom = "(((((r_citausuario" & _
              " RIGHT JOIN (t_cita LEFT JOIN t_ifocusuario ON t_cita.fkIfocUsuariotec = t_ifocusuario.fkPersona) ON r_citausuario.fkCita = t_cita.id)" & _
              " LEFT JOIN v_apellidosnombre ON r_citausuario.fkPersona = v_apellidosnombre.id)" & _
              " LEFT JOIN t_organizacion ON r_citausuario.fkOrganizacion = t_organizacion.id)" & _
              " LEFT JOIN t_proyectoemprendedor ON r_citausuario.fkProyectoEmprendedor = t_proyectoemprendedor.id)" & _
              " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN t_ifocusuario AS t_ifocusuario1 ON r_citausuario.fkIfocUsuario = t_ifocusuario1.fkPersona"
    
    strWhere = ""
    strOrderBy = ""
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"
    
    
    If (idCita <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_cita.id = " & idCita)
    Else
        If (fechaInicio <> "01/01/1900") Then
            strWhere = addConditionWhere(strWhere, _
                                         "t_cita.fecha Between #" & fechaI & "# AND #" & FECHAF & "#")
        End If
        
        If (idPersona <> 0) Then
            strWhere = addConditionWhere(strWhere, "r_citausuario.fkPersona = " & idPersona)
        End If
        If (idOrganizacion <> 0) Then
            strWhere = addConditionWhere(strWhere, "r_citausuario.fkOrganizacion = " & idOrganizacion)
        End If
        If (idOferta <> 0) Then
            strWhere = addConditionWhere(strWhere, "t_cita.fkOferta = " & idOferta)
        End If
        If (idIfocUsuario <> 0) Then
            strWhere = addConditionWhere(strWhere, "t_cita.fkIfocUsuarioTec = " & idIfocUsuario)
        End If
        
        If (order = 1) Then '1=ASC
            strOrderBy = "t_cita.fecha ASC"
        ElseIf (order = 2) Then '2=DESC
            strOrderBy = "t_cita.fecha DESC"
        End If
    End If
    
'Debug.Print montarSQL(strselect, _
                        strFrom, _
                        strWhere, _
                        , _
                        , _
                        strOrderBy)
    
    getSqlListadoCitasUsuario = montarSQL(strselect, _
                                        strFrom, _
                                        strWhere, _
                                        , _
                                        , _
                                        strOrderBy)
End Function

'------------------------------------------------------------------------------------
'   Name:   getSqlListadoCitas
'   Autor:  Jose Manuel Sanchez
'   Fecha:  22/09/2009
'   Desc:   devuelve un string con el sql del listado de citas que cumplen los
'           criterios que se pasan por parámetro
'   Param:  fechaInicio(date), fecha inicio del listado de citas
'           fechaFin(date), fecha fin del listado de gestiones
'           idPersona(long), identificador de persona (opcional)
'           idOrganizacion(long), identificador de Organizacion(opcional)
'           idOferta(long), identificador de oferta(opcional)
'           idCurso(long), identificador de curso(opcional)
'           idIfocUsuario(long), identificador de usuario ifoc(optional)
'           order(integer), order by fechaHora (1=ASC, 2=DESC, otro=no orden)
'   Retur:  sql(string), listado de gestiones solicitadas
'-------------------------------------------------------------------------------------
Public Function getSqlListadoCitas(Optional fechaInicio As Date = "01/01/1900", _
                                    Optional fechaFin As Date, _
                                    Optional idPersona As Long = 0, _
                                    Optional idOrganizacion As Long = 0, _
                                    Optional idOferta As Long = 0, _
                                    Optional idIfocUsuario As Long = 0, _
                                    Optional order As Integer = 0) As String
    Dim strselect As String
    Dim strFrom As String
    Dim strWhere As String
    Dim strOrderBy As String
    
    Dim fechaI As Date
    Dim FECHAF As Date
    
    'Inicializamos variables
    strselect = " t_cita.id," & _
                " IIf([fkCitaSesion]=1,'Cita (I)','Cita (G)') AS IoG," & _
                " a_gestiontipo.tipo AS Tipo," & _
                " format(t_cita.fecha,'dd/mm/yyyy hh:mm') as Fecha," & _
                " format(t_cita.fecha, 'hh:mm') AS Hora," & _
                " t_ifocusuario.aka AS [Técnic@]," & _
                " IIf(t_cita.acude=-1,'Sí', IIf(t_cita.acude=0,'No','?')) AS Acude," & _
                " IIf(r_citausuario.acudeNoacudeAnula=1,'Sí','No') AS Anula," & _
                " t_ifocusuario1.aka as Citador," & _
                " a_servicio.aka AS Servicio," & _
                " format(t_cita.fecha,'yyyy/mm/dd hh:mm') as fechaOrder"

    strFrom = "((((r_citausuario" & _
              " RIGHT JOIN (t_cita LEFT JOIN t_ifocusuario ON t_cita.fkIfocUsuariotec = t_ifocusuario.fkPersona) ON r_citausuario.fkCita = t_cita.id)" & _
              " LEFT JOIN v_apellidosnombre ON r_citausuario.fkPersona = v_apellidosnombre.id)" & _
              " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN t_ifocusuario AS t_ifocusuario1 ON r_citausuario.fkIfocUsuario = t_ifocusuario1.fkPersona)" & _
              " LEFT JOIN a_servicio ON t_cita.fkServicio = a_servicio.id"

    strWhere = ""
    strOrderBy = ""
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"
    
    If (fechaInicio <> "01/01/1900") Then
        strWhere = addConditionWhere(strWhere, _
                                     "t_cita.fecha Between #" & fechaI & "# AND #" & FECHAF & "#")
    End If
    
    If (idPersona <> 0) Then
        strWhere = addConditionWhere(strWhere, "r_citausuario.fkPersona = " & idPersona)
    End If
    If (idOrganizacion <> 0) Then
        strWhere = addConditionWhere(strWhere, "r_citausuario.fkOrganizacion = " & idOrganizacion)
    End If
    If (idOferta <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_cita.fkOferta = " & idOferta)
    End If
    If (idIfocUsuario <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_cita.fkIfocUsuarioTec = " & idIfocUsuario)
    End If
    
    If (order = 1) Then
        strOrderBy = "t_cita.fecha ASC"
    ElseIf (order = 2) Then
        strOrderBy = "t_cita.fecha DESC"
    End If
    
    getSqlListadoCitas = montarSQL(strselect, _
                                strFrom, _
                                strWhere, _
                                , _
                                , _
                                strOrderBy)
                                
End Function

'-------------------------------------------------------------------------------------------
'   Name:   insProyEmprendedorCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/09/2009
'   Desc:   Añade un proyecto emprendedor a la cita
'   Param:  idCita, identificador de la cita donde añadimos la persona
'           idProyEmprendedor, identificador de proyecto emprendedor donde añadimos la cita
'   Return:
'-------------------------------------------------------------------------------------------
Public Function insProyEmprendedorCita(idCita As Long, _
                                        idProyEmprendedor As Long, _
                                        idIfocUsuario As Long) As Integer
    On Error GoTo Error
    
    Dim str As String
    str = " INSERT INTO r_citausuario (fkCita, fkProyectoEmprendedor, fkIfocUsuario)" & _
          " VALUES (" & idCita & ", " & idProyEmprendedor & ", " & idIfocUsuario & ");"

'debugando str
    CurrentDb.Execute str
    
    insProyEmprendedorCita = 0
    Exit Function
    
Error:
    Debug.Print "Error: " & Err.description
    insProyEmprendedorCita = -1
End Function

'---------------------------------------------------------------------------
'   Name:   insProyEmprendedoresCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  29/9/2009
'   Desc:   Inserta la/s personas pasada por parámetro string
'   Param:  idCita(long), identificador de cita
'           ProyEmprendedores(string), identificadores de personas separados por ','
'   Retur:  -1, error al insertar
'            0, insercion correcta OK
'---------------------------------------------------------------------------
'Private
Public Function insProyEmprendedoresCita(idCita As Long, _
                                         ProyEmprendedores As String, _
                                         idIfocUsuario As Long) As Integer
    Dim sql As String
    Dim args As Variant
    Dim idProyEmprendedor As Long
    Dim numProyEmprendedores As Integer
    Dim i As Integer
On Error GoTo TratarError
    
    args = Split(ProyEmprendedores, ",")
    numProyEmprendedores = UBound(args)
        
    For i = 0 To numProyEmprendedores - 1
        idProyEmprendedor = args(i)
        If insProyEmprendedorCita(idCita, idProyEmprendedor, idIfocUsuario) = -1 Then
            GoTo TratarError
        End If
    Next
    
    insProyEmprendedoresCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    insProyEmprendedoresCita = -1
    debugando Err.description
End Function

'---------------------------------------------------------------------------
'   Name:   delProyEmprendedorCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  24/09/2009
'   Desc:   Elimina el proyecto emprendedor pasado por parámetro asociada a la cita(todas)
'   Param:  idCita(long), identificador de cita
'           idProyEmprendedor(long), identificador de proyecto de cita
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delProyEmprendedorCita(idCita As Long, _
                                        idProyEmprendedor As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkCita" & _
          " FROM r_citausuario" & _
          " WHERE fkCita=" & idCita & " AND fkProyectoEmprendedor= " & idProyEmprendedor & ";"
    
    CurrentDb.Execute sql
    
    delProyEmprendedorCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delProyEmprendedorCita = -1
    debugando "Error: " & Err.description
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/9/2009
'   Name:   delProyEmprendedoresDeCita
'   Desc:   Elimina la/s proyectos emprendedores asociadas a la cita(todas)
'   Param:  identificador de cita
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delProyEmprendedoresDeCita(idCita As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkCita" & _
          " FROM r_citausuario" & _
          " WHERE fkCita=" & idCita & "  AND fkProyectoEmprendedor is not null;"
    
    CurrentDb.Execute sql
    
    delProyEmprendedoresDeCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delProyEmprendedoresDeCita = -1
    Resume SalirTratarError
End Function

'----------------------------------------------------------------------------------------------------------
'   Name:   delCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  25/09/2009
'   Desc:   Eliminamos la cita pasada por parámetro en tabla t_Gestion
'   Param:  idCita(*long)    identificador cita a eliminar
'   Return:  0 -> Ok
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Public Function delCita(idCita As Long) As Integer
    Dim sqlCita As String
    
On Error GoTo TratarError
    
    sqlCita = "DELETE FROM t_cita WHERE id=" & idCita & ";"
    CurrentDb.Execute sqlCita
    delCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delCita = -1
End Function

'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  29/07/2010
'   Name:   delCitaConfiracion
'   Desc:   Eliminamos la cita pasada por parámetro en tabla t_Cita
'           solicitando confirmación de borrado
'   Param:  idCita(*long) identificador gestion a eliminar
'   Return:  0 -> Ok
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Public Function delCitaConfirmacion(idCita As Long) As Integer
    Dim str As String
    Dim strRespuesta As String
    Dim num As Integer
On Error GoTo TratarError
    str = "¿Está seguro de querer borrar la cita (" & idCita & ")?"
    
    If MsgBox(str, vbOKCancel, "Eliminar gestión") = vbOK Then
        If esCaptchaCorrecto() Then
            delCita idCita
        End If
    End If
SalirTratarError:
    delCitaConfirmacion = 0
    Exit Function
TratarError:
    delCitaConfirmacion = -1
End Function

'----------------------------------------------------------------------------
Public Function test1() As String
    'initSqlUnionCitaOrganizacion
    
    'G_strSql = montarSQL(G_strSqlUnionOSelect, _
                         G_strSqlUnionOFrom, _
                         G_strSqlUnionOWhere, _
                         G_strSqlUnionOGroupby, _
                         G_strSqlUnionOHaving, _
                         G_strSqlUnionOOrderBy)
                         
    'test1 = G_strSql
End Function

'---------------------------------------------------------------------------
'   Name:   delOrganizacionCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  10/05/2010
'   Desc:   Elimina la organización pasada por parámetro asociada a la cita(todas)
'   Param:  idCita(long), identificador de cita
'           idOrganizacion(long), identificador de organizacion
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delOrganizacionCita(idCita As Long, _
                                    idOrganizacion As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkCita" & _
          " FROM r_citausuario" & _
          " WHERE fkCita=" & idCita & " AND fkOrganizacion= " & idOrganizacion & ";"
    
    CurrentDb.Execute sql
    
    delOrganizacionCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delOrganizacionCita = -1
    debugando "Error: " & Err.description
End Function


'---------------------------------------------------------------------------
'   Name:   delOrganizacionesDeCita
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  15/9/2009 - Act: 15/9/2009
'   Desc:   Elimina la/s organizaciones asociadas a la cita(todas)
'   Param:  identificador de cita
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delOrganizacionesDeCita(idCita As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkCita" & _
          " FROM r_citausuario" & _
          " WHERE fkCita=" & idCita & "  AND fkOrganizacion is not null;"
    
    CurrentDb.Execute sql
    
    delOrganizacionesDeCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delOrganizacionesDeCita = -1
    Resume SalirTratarError
End Function

'-------------------------------------------------------------------------------------------
'   Name:   insOrganizacionCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/09/2009
'   Desc:   Añade una organizacion a la cita
'   Param:  idCita, identificador de la cita donde añadimos la persona
'           idOrganizacion, identificador de organización donde añadimos la cita
'   Return:
'-------------------------------------------------------------------------------------------
Public Function insOrganizacionCita(idCita As Long, _
                                    idOrganizacion As Long, _
                                    idIfocUsuario As Long) As Integer
    On Error GoTo Error
    
    Dim str As String
    str = " INSERT INTO r_citausuario (fkCita, fkOrganizacion, fkIfocUsuario)" & _
          " VALUES (" & idCita & ", " & idOrganizacion & ", " & idIfocUsuario & ");"

'debugando str
    CurrentDb.Execute str
    
    insOrganizacionCita = 0
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    insOrganizacionCita = -1
End Function

'-------------------------------------------------------------------------------------
'   Name:   insOrganizacionesCita
'   Autor:  Jose Manuel Sanchez
'   Fecha:  29/9/2009
'   Desc:   Inserta la/s personas pasada por parámetro string
'   Param:  idCita(long), identificador de cita
'           Organizaciones(string), identificadores de organizacion separados por ','
'   Retur:  -1, error al insertar
'            0, insercion correcta OK
'-------------------------------------------------------------------------------------
'Private
Public Function insOrganizacionesCita(idCita As Long, _
                                      Organizaciones As String, _
                                      idIfocUsuario As Long) As Integer
    Dim sql As String
    Dim args As Variant
    Dim idOrganizacion As Long
    Dim numOrganizaciones As Integer
    Dim i As Integer
On Error GoTo TratarError
    
    numOrganizaciones = countSubStrings(Organizaciones, ",")
    args = Split(Organizaciones, ",")
    
    'args = Split(Organizaciones, ",")
    'numOrganizaciones = UBound(args)
        
    For i = 0 To numOrganizaciones - 1
        idOrganizacion = args(i)
        If insOrganizacionCita(idCita, idOrganizacion, idIfocUsuario) = -1 Then
            GoTo TratarError
        End If
    Next
    
    insOrganizacionesCita = 0
    
SalirTratarError:
    Exit Function
TratarError:
    insOrganizacionesCita = -1
    debugando Err.description
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  23/07/2010 - Actualización:  23/07/2010
'   Name:   existeCitaOrganizacion
'   Desc:   Comprueba si existe el proyecto en la cita
'   Param:  idCita (long)
'           idProyecto(long)
'   Retur:  numero de registros encontrados
'---------------------------------------------------------------------------
Public Function existeCitaProyecto(idCita As Long, idProyecto As Long) As Integer

    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Registros " & _
             "FROM r_citausuario " & _
             "WHERE fkCita = " & idCita & _
             "  AND fkProyectoEmprendedor =  " & idProyecto & ";"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    existeCitaProyecto = rs!Registros
    
    rs.Close
    Set rs = Nothing
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  23/07/2010 - Actualización:  23/07/2010
'   Name:   existeCitaOrganizacion
'   Desc:   Comprueba si existe la organización en la cita
'   Param:  idCita (long)
'           idOrganizacion(long)
'   Retur:  numero de registros encontrados
'---------------------------------------------------------------------------
Public Function existeCitaOrganizacion(idCita As Long, idOrganizacion As Long) As Integer

    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = " SELECT COUNT(*) AS Registros" & _
             " FROM r_citausuario" & _
             " WHERE fkCita = " & idCita & _
             " AND fkOrganizacion =  " & idOrganizacion & ";"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    existeCitaOrganizacion = rs!Registros
    
    rs.Close
    Set rs = Nothing
End Function

'---------------------------------------------------------------------------------------
'       Consultas de cita para contadores
'---------------------------------------------------------------------------------------
Public Static Function getNumCitasPersona(idPersona As Long, fechaInicio As Date, fechaFin As Date) As Long
    SIFOC_Cita.initSql
    SIFOC_Cita.addStrWhere ("")
    SIFOC_Cita.addStrWhere ("")
    SIFOC_Cita.addStrWhere ("")
End Function

Private Static Function clearSql()
    strselect = ""
    strFrom = ""
    strWhere = ""
    strHaving = ""
    strGroup = ""
    strOrder = ""
End Function


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/02/2014 - Actualización:  24/02/2014
'   Name:   initSqlCitaOferta
'   Desc:   Listado
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSqlCitaOferta()
'idPersona
't_cita.fecha
'a_ifocambito.ambito
'a_citasesion.sesion
'a_gestiontipo.tipo
'a_servicio.servicio
'a_serviciosubtipo.subtipo
'uifocc.aka AS citador
'uifoct.aka AS tecnico
't_cita.duracion
't_cita.fechaDemanda
't_cita.observacion
't_cita.acude
't_cita.cancelada
't_cita.fkOferta
't_cita.idCurso
  
    
    strselect = "t_cita.fkOferta AS idOferta, t_oferta.puesto, t_cita.fecha, a_ifocambito.ambito, a_citasesion.sesion, a_gestiontipo.tipo, a_servicio.servicio, a_serviciosubtipo.subtipo, uifocc.aka AS citador, uifoct.aka AS tecnico, t_cita.duracion, t_cita.fechaDemanda, t_cita.observacion, t_cita.acude, t_cita.cancelada"
    strFrom = " (((((((t_cita" & _
              " LEFT JOIN a_gestiontipo ON t_cita.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN a_servicio ON t_cita.fkServicio = a_servicio.id)" & _
              " LEFT JOIN a_serviciosubtipo ON t_cita.fkServicioSubtipo = a_serviciosubtipo.id)" & _
              " LEFT JOIN a_ifocambito ON t_cita.fkIfocAmbito = a_ifocambito.id)" & _
              " LEFT JOIN a_citasesion ON t_cita.fkCitaSesion = a_citasesion.id)" & _
              " LEFT JOIN t_ifocusuario as uifocc ON t_cita.fkIfocusuarioCit = uifocc.fkPersona)" & _
              " LEFT JOIN t_ifocusuario as uifoct ON t_cita.fkIfocusuarioTec = uifoct.fkPersona)" & _
              " LEFT JOIN t_oferta ON t_cita.fkOferta = t_oferta.id"
    
    'Estado abierto y pendiente
    strWhere = "(Not (t_cita.fkOferta) Is Null)"
    strHaving = ""
    strGroup = "t_cita.fkOferta, t_oferta.puesto, t_cita.fecha, a_ifocambito.ambito, a_citasesion.sesion, a_gestiontipo.tipo, a_servicio.servicio, a_serviciosubtipo.subtipo, uifocc.aka, uifoct.aka, t_cita.duracion, t_cita.fechaDemanda, t_cita.observacion, t_cita.acude, t_cita.cancelada"
    strOrder = "t_cita.fecha DESC"
End Function

'Public Static Function getNumCitasOrganizacion(idOrganizacion As Long, fechaInicio As Date, fechaFin As Date) As Long
'    Dim rs As ADODB.Recordset
'    Dim counter As Long
'
'On Error GoTo TratarError
'
'    SIFOC_Cita.initSql
'    SIFOC_Cita.addStrWhere ("fecha >=#" & Format(fechaInicio, "yyyy-mm-dd hh:nn") & "#")
'    SIFOC_Cita.addStrWhere ("fecha <=#" & Format(fechaFin, "yyyy-mm-dd 23:59") & "#")
'    SIFOC_Cita.addStrWhere ("fkOrganizacion = " & idOrganizacion)
'
'    Set rs = New ADODB.Recordset
'
'Debug.Print SIFOC_Cita.getQuery
'
'    rs.Open SIFOC_Cita.getQuery, CurrentProject.Connection, adOpenStatic, adLockReadOnly
'
'    counter = 0
'    If Not rs.EOF Then
'        counter = rs.RecordCount
'    End If
'
'    rs.Close
'    Set rs = Nothing
'
'    getNumCitasOrganizacion = counter
'
'SalirTratarError:
'    Exit Function
'TratarError:
'    getNumCitasOrganizacion = -1
'    Debug.Print Err.description
'End Function

Public Static Function getNumCitasTecnico(idIfocUsuario As Long, fechaInicio As Date, fechaFin As Date) As Long
    Dim strSql As String
    Dim rs As dao.Recordset
    Dim counter As Long
    
On Error GoTo TratarError

    SIFOC_Cita.clearSql
    SIFOC_Cita.setStrSelect ("Count(id) as gestiones")
    SIFOC_Cita.setStrFrom ("t_cita")
    SIFOC_Cita.addStrWhere ("fecha >=#" & Format(fechaInicio, "yyyy-mm-dd hh:nn") & "#")
    SIFOC_Cita.addStrWhere ("fecha <=#" & Format(fechaFin, "yyyy-mm-dd 23:59") & "#")
    SIFOC_Cita.addStrWhere ("fkIfocUsuarioTec = " & idIfocUsuario)
    SIFOC_Cita.addStrWhere ("acude = -1")
    SIFOC_Cita.setStrGroup ("")
    SIFOC_Cita.setStrOrder ("")
'Debug.Print SIFOC_Cita.getQuery & vbNewLine & getStrOrder

    strSql = SIFOC_Cita.getQuery
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open strsql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    counter = 0
    If Not rs.EOF Then
        counter = rs!Gestiones
    End If
    
    rs.Close
    Set rs = Nothing
    
    getNumCitasTecnico = counter

SalirTratarError:
    Exit Function
TratarError:
    getNumCitasTecnico = -1
    Debug.Print Err.description
End Function

Public Static Function getNumCitasOferta(idOferta As Long, fechaInicio As Date, fechaFin As Date) As Long
    
End Function

Public Static Function getNumCitasPriyectoEmprendedor(idProyectoEmprendedor As Long, fechaInicio As Date, fechaFin As Date) As Long
    
End Function

Public Static Function getNumCitasCurso(idCurso As Long, fechaInicio As Date, fechaFin As Date) As Long
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  11/03/2014 - Actualización: 11/03/2014
'   Name:   getSqlCitaAvisoSms
'   Desc:   Para listado de personas y telefono de aviso cita
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function getSqlCitaAvisoSms(fecha As Date, idTipoCita As Long)
    Dim sqlTel As String
    
    sqlTel = " SELECT fkPersona as idPersona, telefono FROM t_telefono WHERE fkTipoTelefono1 = 1 AND fkTipoTelefono2 = 3"
    
    strselect = "r_citausuario.fkPersona AS idPersona, v_datospersonales.name, tel.telefono, t_cita.fecha, v_ifocusuario.name as tecnico, t_cita.id as idCita "
    strFrom = "(((r_citausuario" & _
              " LEFT JOIN (" & sqlTel & ") as tel ON r_citausuario.fkPersona = tel.idPersona)" & _
              " LEFT JOIN t_cita ON r_citausuario.fkCita = t_cita.id)" & _
              " LEFT JOIN v_datospersonales ON r_citausuario.fkPersona = v_datospersonales.id)" & _
              " LEFT JOIN v_ifocusuario ON r_citausuario.fkPersona = v_ifocusuario.fkPersona"
    
    'Estado abierto y pendiente
    strWhere = "(Not (r_citausuario.fkPersona) Is Null) AND t_cita.fecha between #" & Format(fecha, "mm/dd/yyyy") & "# AND #" & Format(fecha, "mm/dd/yyyy 23:59") & "#"
    If (idTipoCita <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_cita.fkGestionTipo=" & idTipoCita)
    End If
    
    strHaving = ""
    strGroup = ""
    strOrder = "r_citausuario.fkPersona ASC"
    
    getSqlCitaAvisoSms = getQuery
End Function

'-------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Name: isCitaDescripcionNotEmpty
'   Fecha: 13/10/2014 -  Act:13/10/2014
'   Desc: Actualiza tipo de sesion de la cita en caso de haber más
'         de una persona
'   Parm: inicial(boolean) indica si es al abrir formulario
'   Retr: -
'-------------------------------------------------------------------
Public Function isCitaDescripcionNotEmpty(idCita As Long, OBS As String) As Integer
    Dim result As Long
    Dim descr As String
    Dim duracion As Integer
    Dim duracionMax As Integer
    
    descr = Nz(DLookup("observacion", "t_cita", "[id]=" & idCita), "")
    duracion = Nz(DLookup("duracion", "t_cita", "[id]=" & idCita), 0)
    
    result = -1
    If (duracion / 6 < 40) Then
        duracionMax = duracion / 6
    Else
        duracionMax = 40
    End If
    
Debug.Print duracionMax & " - "

    If Len(descr) = 0 Then
        descr = OBS
    End If

    If Len(descr) > 10 _
        And Len(descr) > duracionMax _
        And descr <> "" Then
        
        result = 0
    End If
    
    isCitaDescripcionNotEmpty = result
End Function

