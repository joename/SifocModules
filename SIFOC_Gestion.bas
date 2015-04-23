Attribute VB_Name = "SIFOC_Gestion"
Option Explicit
Option Compare Database

'> Guardamos la forma de contacto de la gestión(sólo se toca por el formulario FormaContacto)
Global U_formaContacto As Integer

Private strSql As String
Private strselect As String
Private strFrom As String
Private strWhere As String
Private strGroup As String
Private strHaving As String
Private strOrder As String

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

Public Static Function initSqlClean()
    strselect = ""
    strFrom = ""
    strWhere = ""
    strHaving = ""
    strGroup = ""
    strOrder = ""
End Function

Public Static Function getQuery() As String
    strSql = montarSQL(strselect, _
                       strFrom, _
                       strWhere, _
                       strGroup, _
                       strHaving, _
                       strOrder)
'Debug.Print strSql
    getQuery = strSql
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
  
    sqlEmp = " SELECT r_gestionusuario.fkGestion, Count(r_gestionusuario.fkPersona) AS numOrganizaciones" & _
              " FROM r_gestionusuario" & _
              " WHERE ((Not (r_gestionusuario.fkPersona) Is Null))" & _
              " GROUP BY r_gestionusuario.fkGestion"
              
    sqlPer = " SELECT r_gestionusuario.fkGestion, Count(r_gestionusuario.fkOrganizacion) AS numPersonas" & _
              " FROM r_gestionusuario" & _
              " WHERE ((Not (r_gestionusuario.fkOrganizacion) Is Null))" & _
              " GROUP BY r_gestionusuario.fkGestion"
              
    'sqlPry = " SELECT r_gestionusuario.fkGestion, Count(r_gestionusuario.fkProyectoEmprendedor) AS numProyectos" & _
              " FROM r_gestionusuario" & _
              " WHERE ((Not (r_gestionusuario.fkProyectoEmprendedor) Is Null))" & _
              " GROUP BY r_gestionusuario.fkGestion"
    
    strselect = "t_gestion.id, fecha, ambito, tipo, a_servicio.aka as servicio, subtipo,  uifoct.aka as tecnico, numPersonas, numOrganizaciones, gestion, fkOferta, fkCurso"
    strFrom = "(((((((t_gestion" & _
              " LEFT JOIN (" & sqlEmp & ")as emp ON t_gestion.id = emp.fkGestion)" & _
              " LEFT JOIN (" & sqlPer & ")as per ON t_gestion.id = per.fkGestion)" & _
              " LEFT JOIN a_gestiontipo ON t_gestion.fkGestionTipo = a_gestiontipo.id)" & _
              " LEFT JOIN a_servicio ON t_gestion.fkServicio = a_servicio.id)" & _
              " LEFT JOIN a_serviciosubtipo ON t_gestion.fkServicioSubtipo = a_serviciosubtipo.id)" & _
              " LEFT JOIN a_ifocambito ON t_gestion.fkIfocAmbito = a_ifocambito.id)" & _
              " LEFT JOIN t_ifocusuario as uifoct ON t_gestion.fkIfocusuario = uifoct.fkPersona)"
    
    'Estado abierto y pendiente
    strWhere = ""
    strHaving = ""
    strGroup = "t_gestion.id, fecha, ambito, tipo, a_servicio.aka, subtipo, uifoct.aka, gestion, fkOferta, fkCurso, numPersonas, numOrganizaciones"
    strOrder = "fecha DESC"
End Function

Public Static Function getNumGestionesTecnico(idIfocUsuario As Long, fechaInicio As Date, fechaFin As Date) As Long
    Dim strSql As String
    Dim rs As dao.Recordset
    Dim counter As Long
    
On Error GoTo TratarError

'    SIFOC_Gestion.setStrSelect ("id")
'    SIFOC_Gestion.setStrFrom ("t_gestion")
    SIFOC_Gestion.initSqlClean
    SIFOC_Gestion.setStrSelect ("Count(id) as gestiones")
    SIFOC_Gestion.setStrFrom ("t_gestion")
    SIFOC_Gestion.addStrWhere ("fecha >=#" & Format(fechaInicio, "yyyy-mm-dd hh:nn") & "#")
    SIFOC_Gestion.addStrWhere ("fecha <=#" & Format(fechaFin, "yyyy-mm-dd 23:59") & "#")
    SIFOC_Gestion.addStrWhere ("fkIfocUsuario = " & idIfocUsuario)
    SIFOC_Gestion.setStrGroup ("")
    SIFOC_Gestion.setStrOrder ("")
    
    strSql = SIFOC_Gestion.getQuery
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open strsql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    counter = 0
    If Not rs.EOF Then
        counter = rs!Gestiones
    End If
    
    rs.Close
    Set rs = Nothing
    
    getNumGestionesTecnico = counter

SalirTratarError:
    Exit Function
TratarError:
    getNumGestionesTecnico = -1
    Debug.Print Err.description
End Function


'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/09/2009
'   Name:   CreaGestion
'   Desc:   Creamos nueva gestión en tabla t_Gestion SIN PERSONA NI ORGANIZACION al ser grupal!!
'   Param:  fkifocambito(*),
'           fecha(*),
'           hora(*),
'           gestion(*),
'           idIfocUsuario(*),
'           idPersona(*),
'           idOrganizacion(*),
'           idOferta(*),
'           idAutoempleoProyecto(*),
'           idCurso(*),
'           idFormaContacto,
'           idMenuGestion)
'   Return:  0 -> Ok
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Private Function creaGestion(idGestionTipo As Integer, _
                            idIfocAmbito As Integer, _
                            fecha As Date, _
                            gestion As String, _
                            idIfocUsuario As Long, _
                            Optional idServicio As Long = 0, _
                            Optional idOferta As Long = 0, _
                            Optional idAutoempleoProyecto As Long = 0, _
                            Optional idCurso As Long = 0, _
                            Optional idFormaContacto As Integer = 0) As Integer
    Dim strGestion As String
    Dim sqlGestion As String
    Dim sqlInsert As String
    Dim sqlValues As String
    
On Error GoTo TratarError

    strGestion = filterSQL(gestion)
    'isMissing(argumento)
    'Tabla t_gestion (campos)
    'fkPersona*,fkifocambito*,fecha*,hora*,gestion*,fkIfocUsuario*,fkFormaContacto,fkMenuGestion,fkOferta
    
    'fecha en insert formato mm/dd/yyyy
    
    sqlInsert = " INSERT INTO t_gestion (" & _
                    " fkGestionTipo" & _
                    ", fkIfocAmbito" & _
                    ", fecha " & _
                    ", gestion" & _
                    ", fkIfocUsuario" & _
                    IIf(idServicio = 0, "", ", fkServicio ") & _
                    IIf(idOferta = 0, "", ", fkOferta ") & _
                    IIf(idAutoempleoProyecto = 0, "", ", fkProyectoEmprendedor ") & _
                    IIf(idCurso = 0, "", ", fkCurso ") & _
                    IIf(idFormaContacto = 0, "", ", fkFormaContacto ") & ")"
    sqlValues = " VALUES (" & _
                    idGestionTipo & " " & _
                    ", " & idIfocAmbito & " " & _
                    ", '" & Format(fecha, "dd/mm/yyyy hh:mm:ss") & "' " & _
                    ", '" & strGestion & "' " & _
                    ", " & idIfocUsuario & " " & _
                    IIf(idServicio = 0, "", ", " & idServicio & " ") & _
                    IIf(idOferta = 0, "", ", " & idOferta & " ") & _
                    IIf(idAutoempleoProyecto = 0, "", ", " & idAutoempleoProyecto & " ") & _
                    IIf(idCurso = 0, "", ", " & idCurso & " ") & _
                    IIf(idFormaContacto = 0, "", ", " & idFormaContacto & " ") & _
                    ");"

    sqlGestion = sqlInsert & sqlValues
'debug.print sqlGestion
    CurrentDb.Execute sqlGestion
    
    creaGestion = 0
    
SalirTratarError:
    Exit Function
TratarError:
    creaGestion = -1
    Resume SalirTratarError
End Function

'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/09/2009
'   Name:   CreaGestionGrupal
'   Desc:   Creamos nueva gestión en tabla t_Gestion y añadimos grupo personas y/o organizaciones!!
'   Param:  fkifocambito(*),
'           fecha(*),
'           hora(*),
'           gestion(*),
'           idIfocUsuario(*),
'           Personas(*),
'           Organizaciones(*),
'           idOferta(*),
'           idAutoempleoProyecto(*),
'           idCurso(*),
'           idFormaContacto,
'           idMenuGestion)
'   Return:  >0 -> identificador de la gestión
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Public Function creaGestionGrupal(idGestionTipo As Integer, _
                                  idIfocAmbito As Integer, _
                                  fecha As Date, _
                                  gestion As String, _
                                  idIfocUsuario As Long, _
                                  Optional idServicio As Long = 0, _
                                  Optional Personas As String = "", _
                                  Optional Organizaciones As String = "", _
                                  Optional idOferta As Long = 0, _
                                  Optional idAutoempleoProyecto As Long = 0, _
                                  Optional idCurso As Long = 0, _
                                  Optional idFormaContacto As Integer = 0) As Long
On Error GoTo TratarError
    
    Dim strGestion As String
    
    'strGestion = filterSQL(gestion) 'filtro en creaGestion
    
    If creaGestion(idGestionTipo, _
                    idIfocAmbito, _
                    fecha, _
                    gestion, _
                    idIfocUsuario, _
                    idServicio, _
                    idOferta, _
                    idAutoempleoProyecto, _
                    idCurso, _
                    idFormaContacto) = -1 Then
        
        creaGestionGrupal = -1
        Exit Function
    End If
    
    Dim idGestion As Long
    
    idGestion = getIdGestion(idIfocAmbito, fecha, , idIfocUsuario)

    'Añadimos persona/s
    If Len(Personas) > 0 Then
        If insPersonasGestion(idGestion, Personas, idIfocUsuario) = -1 Then
            GoTo TratarError
        End If
    End If
    
    'Añadimos organización/es
    If Len(Organizaciones) > 0 Then
        If insOrganizacionesGestion(idGestion, Organizaciones, idIfocUsuario) = -1 Then
            GoTo TratarError
        End If
    End If
    
    creaGestionGrupal = idGestion
SalirTratarError:
    Exit Function
TratarError:
    creaGestionGrupal = -1
    Debug.Print Err.description
    Resume SalirTratarError
End Function

'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/09/2009
'   Name:   getIdGestion
'   Desc:   Obtenemos el id de gestión de la gestión con los datos pasados por parámetro
'           PARA OBTENER ID DE GESTION GRUPAL
'   Param:  fkifocambito(*),
'           fecha(*),
'           hora(*),
'           gestion(*),
'           idIfocUsuario(*))
'   Return:  0 -> Ok
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Private Function getIdGestion(fkIfocAmbito As Integer, _
                             fecha As Date, _
                             Optional gestion As String = "", _
                             Optional idIfocUsuario = 0) As Long
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim strGestion As String
    
    strGestion = filterSQL(gestion)
    
    sql = " SELECT id, timestamp" & _
          " FROM t_gestion" & _
          " WHERE fkifocambito=" & fkIfocAmbito & " AND fecha =#" & Format(fecha, "mm/dd/yyyy hh:nn:ss") & "#" & _
          IIf(gestion = "", "", " AND gestion = '" & strGestion & "'") & _
          IIf(idIfocUsuario = 0, "", " AND fkIfocUsuario=" & idIfocUsuario) & _
          " ORDER BY timestamp DESC"
    
'    sql = " SELECT id" & _
'          " FROM t_gestion" & _
'          " WHERE fkifocambito=" & fkIfocAmbito & " AND fecha =#" & Format(fecha, "mm/dd/yyyy") & "#" & _
'          " AND hora =#" & Format(fecha, "mm/dd/yyyy") & " " & Format(hora, "hh:nn:ss") & "#" & _
'          IIf(gestion = "", "", " AND gestion = '" & strGestion & "'") & _
'          IIf(idIfocUsuario = 0, "", "AND fkIfocUsuario=" & idIfocUsuario) & _
'          ";"
    
    Set rs = New ADODB.Recordset
    
'debugando sql

    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then 'And rs.RecordCount = 1 Then
        rs.MoveFirst
        getIdGestion = rs!id
    Else
        getIdGestion = 0
    End If
    
    rs.Close
    Set rs = Nothing
    
End Function

'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  06/09/2011 - Act:  06/09/2011
'   Name:   getIdGestionTipo
'   Desc:   Obtenemos el id de tipo de gestión con el id ifocambito
'   Param:  idIfocambito(*)
'   Return:  idGestionTipo
'----------------------------------------------------------------------------------------------------------
Public Function getIdGestionTipo(idIfocAmbito As Integer) As Long
    Dim sql As String
    
    Select Case idIfocAmbito
        Case 1: getIdGestionTipo = 1 'ambito persona(1) a empleo(1) o formacion(2)
        Case 2: getIdGestionTipo = 4 'ambito empresa(4) a empresa(2)
        Case 3: getIdGestionTipo = 5 'ambito autoempleo(3) a emprendedor(5)
        Case 4: getIdGestionTipo = 3 'ambito ofertas(4) a oferta(3)
    End Select
                             
End Function

'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  06/09/2011 - Act:  06/09/2011
'   Name:   getIdIfocAmbito
'   Desc:   Obtenemos el id de tipo de gestión con el id ifocambito
'   Param:  idGestionTipo(*)
'   Return:  idIfocAmbito
'----------------------------------------------------------------------------------------------------------
Public Function getIdIfocAmbito(idGestionTipo As Integer) As Long
    Dim sql As String
    
    Select Case idGestionTipo
        Case 1: getIdIfocAmbito = 1 'tipo g empleo(1) a persona(1)
        Case 2: getIdIfocAmbito = 1 'tipo g formacion(2)a persona(1)
        Case 3: getIdIfocAmbito = 4 'tipo g oferta(3) a ofertas(4)
        Case 4: getIdIfocAmbito = 2 'tipo g empresa(4) a empresa(2)
        Case 5: getIdIfocAmbito = 3 'tipo g emprendedor(5) a autoempleo(3)
    End Select
                             
End Function

'--------------------------------------------------------------------------------------------
'               Creamos gestion a persona de oferta
'--------------------------------------------------------------------------------------------
Public Function creaGestionOferta(idOferta As Long, _
                                  idPersona As Long, _
                                  observacion As String, _
                                  Optional idFormaContacto As Integer = 0)
    Dim sql As String

    If (idFormaContacto = 0) Then
        'GestionTipo = 3 -> Oferta
        'ID_MENUGES numero 26 de MENU_SUB_GESTIONES ->Agencia - Ofertas(Informacion y Seguimiento)
        creaGestionGrupal 3, _
                          1, _
                          now(), _
                          observacion, _
                          usuarioIFOC(), _
                          8, _
                          CStr(idPersona), _
                          , _
                          idOferta
                    
    ElseIf (idFormaContacto > 0) Then
        creaGestionGrupal 3, _
                          1, _
                          now(), _
                          observacion, _
                          usuarioIFOC(), _
                          8, _
                          CStr(idPersona), _
                          , _
                          idOferta, _
                          , _
                          , _
                          idFormaContacto
    End If

'debugando sql

    'CurrentDb.Execute sql
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Name:   esGestionModificable
'   Desc:   Nos dice si una gestion se pude modificar
'           (sólo modificable hasta dia después inclusive de la fecha de cita
'           para todas las perosonas o excepto el técnico que realiza la cita
'           que la puede modificar siempre)
'   Param:  fechaCita
'   Return: true -> gestion modificable
'           false ->gestion no modificable
'---------------------------------------------------------------------------
Public Function esGestionModificable(fecha As Date, _
                                     idIfocUsuario As Long) As Boolean
    If (fecha > Date - 2) Or (idIfocUsuario = usuarioIFOC()) Then
        esGestionModificable = True
    Else
        esGestionModificable = False
    End If
End Function

'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  25/09/2009
'   Name:   esGestionEliminable
'   Desc:   Devuelve si la gestión la puede modificar el usuario ifoc o no
'   Param:  idGestion(*long)    identificador gestion a eliminar
'           idIfocUsuario(*long) identificador de usuario ifoc
'   Return: TRUE -> Puede ser eliminada por ese usuario
'           FALSE -> No puede ser eliminada por ese usuario
'----------------------------------------------------------------------------------------------------------
Public Function esGestionEliminable(idGestion As Long, _
                                    idIfocUsuario As Long) As Boolean
    Dim idIfocUsuarioGestion As Long
    Dim fechaEntradaGestion As Date
    
On Error GoTo TratarError
    
    idIfocUsuarioGestion = DLookup("fkIfocUsuario", "t_gestion", "[id]=" & idGestion)
    
    If (idIfocUsuario = idIfocUsuarioGestion) Then
        esGestionEliminable = True
    Else
        'idIfocUsuarioGestion = DLookup("fkIfocUsuario", "t_gestion", "[id]=" & idGestion)
        'fechaEntradaGestion = DLookup("fecha", "t_gestion", "[id]=" & idGestion)
        'If (fechaEntradaGestion >= Date And idIfocUsuarioGestion = idIfocUsuario) Then
        '    esGestionEliminable = True
        'Else
        '    esGestionEliminable = False
        'End If
        
        esGestionEliminable = False
    End If
    
SalirTratarError:
    Exit Function
TratarError:
    esGestionEliminable = False
End Function

'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  25/09/2009
'   Name:   delGestion
'   Desc:   Eliminamos la gestión pasada por parámetro en tabla t_Gestion
'   Param:  idGestion(*long)    identificador gestion a eliminar
'   Return:  0 -> Ok
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Public Function delGestion(idGestion As Long) As Integer
    Dim sqlGestion As String
    
On Error GoTo TratarError
    
    sqlGestion = "DELETE FROM t_gestion WHERE id=" & idGestion & ";"
    CurrentDb.Execute sqlGestion
    delGestion = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delGestion = -1
End Function

'----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  25/09/2009
'   Name:   delGestionConfiracion
'   Desc:   Eliminamos la gestión pasada por parámetro en tabla t_Gestion
'           solicitando confirmación de borrado
'   Param:  idGestion(*long) identificador gestion a eliminar
'   Return:  0 -> Ok
'           -1 -> Ko
'----------------------------------------------------------------------------------------------------------
Public Function delGestionConfirmacion(idGestion As Long) As Integer
    Dim str As String
    Dim strRespuesta As String
    Dim num As Integer
On Error GoTo TratarError
    str = "¿Está seguro de querer borrar la gestión (" & idGestion & ")?"
    
    If MsgBox(str, vbOKCancel, "Eliminar gestión") = vbOK Then
        If esCaptchaCorrecto() Then
            delGestion idGestion
        End If
    End If
SalirTratarError:
    delGestionConfirmacion = 0
    Exit Function
TratarError:
    delGestionConfirmacion = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/9/2009
'   Name:   delPersonasDeGestion
'   Desc:   Elimina la/s persona/s asociadas a la gestion(todas)
'   Param:  identificador de gestion
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delPersonasDeGestion(idGestion As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkGestion" & _
          " FROM r_gestionusuario" & _
          " WHERE fkGestion=" & idGestion & ";"
    
    CurrentDb.Execute sql
    
    delPersonasDeGestion = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delPersonasDeGestion = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/9/2009
'   Name:   delOrganizacionesDeGestion
'   Desc:   Elimina la/s organizaciones/s asociadas a la gestion(todas)
'   Param:  identificador de gestion
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delOrganizacionesDeGestion(idGestion As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkGestion" & _
          " FROM r_gestionusuario" & _
          " WHERE (fkGestion=" & idGestion & ") AND (fkOrganizacion is not null);"
    
    CurrentDb.Execute sql
    
    delOrganizacionesDeGestion = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delOrganizacionesDeGestion = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/9/2009
'   Name:   delPersonaGestion
'   Desc:   Elimina la persona pasada por parámetro asociada a la gestion(todas)
'   Param:  identificador de gestion
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delPersonaGestion(idGestion As Long, _
                                  idPersona As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkGestion" & _
          " FROM r_gestionusuario" & _
          " WHERE fkGestion=" & idGestion & " AND fkPersona= " & idPersona & ";"
    
    CurrentDb.Execute sql
    
    delPersonaGestion = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delPersonaGestion = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/9/2009
'   Name:   delOrganizacionGestion
'   Desc:   Elimina la organizacion pasada por parámetro asociada a la gestion(todas)
'   Param:  identificador de gestion
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function delOrganizacionGestion(idGestion As Long, _
                                       idOrganizacion As Long) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
On Error GoTo TratarError
        
    sql = " DELETE fkGestion" & _
          " FROM r_gestionusuario" & _
          " WHERE fkGestion=" & idGestion & " AND fkOrganizacion= " & idOrganizacion & ";"
    
    CurrentDb.Execute sql
    
    delOrganizacionGestion = 0
    
SalirTratarError:
    Exit Function
TratarError:
    delOrganizacionGestion = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/9/2009
'   Name:   insPersonaGestion
'   Desc:   Inserta la persona pasada por parámetro en la gestion
'   Param:  idGestion, identificador de gestion
'           idPersona, identificador de persona
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function insPersonaGestion(idGestion As Long, _
                                  idPersona As Long, _
                                  idIfocUsuario As Long) As Integer
    Dim sql As String
On Error GoTo TratarError

    insPersonaGestion = 0
    
    sql = " INSERT INTO r_gestionusuario ( fkGestion, fkPersona, fkIfocUsuario )" & _
          " VALUES (" & idGestion & ", " & idPersona & ", " & idIfocUsuario & ");"
          
    CurrentDb.Execute sql
    
SalirTratarError:
    Exit Function
TratarError:
    insPersonaGestion = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/9/2009
'   Name:   insPersonasGestion
'   Desc:   Inserta la/s personas pasada por parámetro string
'   Param:  idGestion(long), identificador de gestion
'           Personas(string), identificadores de personas separados por ','
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
'Private
Public Function insPersonasGestion(idGestion As Long, _
                                   Personas As String, _
                                   idIfocUsuario As Long, _
                                   Optional separador As String = ",") As Integer
    Dim args As Variant
    Dim idPersona As Long
    Dim numPersonas As Integer
    Dim i As Integer
On Error GoTo TratarError
    
    insPersonasGestion = 0
    
    numPersonas = countSubStrings(Personas, ",")
    args = Split(Personas, separador)
    
    For i = 0 To numPersonas - 1
        idPersona = args(i)
        If insPersonaGestion(idGestion, idPersona, idIfocUsuario) = -1 Then
            GoTo TratarError
        End If
    Next
    
SalirTratarError:
    Exit Function
TratarError:
    insPersonasGestion = -1
    debugando Err.description
    Resume 'SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/9/2009
'   Name:   insOrganizacionGestion
'   Desc:   Inserta la organizacion pasada por parámetro en la gestion
'   Param:  idGestion, identificador de gestion
'           idPersona, identificador de persona
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
Public Function insOrganizacionGestion(idGestion As Long, _
                                       idOrganizacion As Long, _
                                       idIfocUsuario As Long) As Integer
    Dim sql As String
On Error GoTo TratarError
    
    insOrganizacionGestion = 0

    sql = " INSERT INTO r_gestionusuario ( fkGestion, fkOrganizacion, fkIfocUsuario )" & _
          " VALUES (" & idGestion & ", " & idOrganizacion & ", " & idIfocUsuario & ");"
          
    CurrentDb.Execute sql
SalirTratarError:
    Exit Function
TratarError:
    insOrganizacionGestion = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/9/2009
'   Name:   insOrganizacionesGestion
'   Desc:   Inserta la/s organizacion/es pasada/s por parámetro string
'   Param:  idGestion(long), identificador de gestion
'           Organizaciones(string), identificadores de organizaciones separados por ','
'   Retur:  -1, error en eliminar
'            0, elimación correcta OK
'---------------------------------------------------------------------------
'Private
Public Function insOrganizacionesGestion(idGestion As Long, _
                                         Organizaciones As String, _
                                         idIfocUsuario As Long) As Integer
    Dim sql As String
    Dim args As Variant
    Dim numOrganizaciones As Integer
    Dim idOrganizacion As Long
    Dim i As Integer
On Error GoTo TratarError
    
    insOrganizacionesGestion = 0
    
    numOrganizaciones = countSubStrings(Organizaciones, ",")
    args = Split(Organizaciones, ",")
        
    For i = 0 To numOrganizaciones - 1
        idOrganizacion = args(i)
        If insOrganizacionGestion(idGestion, idOrganizacion, idIfocUsuario) = -1 Then
            GoTo TratarError
        End If
    Next
    
SalirTratarError:
    Exit Function
TratarError:
    insOrganizacionesGestion = -1
    debugando Err.description
    Resume 'SalirTratarError
End Function

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  22/09/2009
'   Name:   getSqlListadoGestiones
'   Desc:   devuelve un string con el sql del listado de gestiones que cumplen los
'           criterios que se pasan por parámetro
'   Param:  fechaInicio(date), fecha inicio del listado de gestiones
'           fechaFin(date), fecha fin del listado de gestiones
'           idPersona(long), identificador de persona (opcional)
'           idOrganizacion(long), identificador de Organizacion(opcional)
'           idOferta(long), identificador de oferta(opcional)
'           idProyectoEmprendedor(long), identificador de autoempleo(opcional)
'           idCurso(long), identificador de curso(opcional)
'           idIfocUsuario(long), identificador de usuario ifoc(optional)
'           order(boolean), introduce order by fecha y hora si es true. Default=false
'   Retur:  sql(string), listado de gestiones solicitadas
'-------------------------------------------------------------------------------------
Public Function getSqlListadoGestiones(Optional fechaInicio As Date = "01/01/1900", _
                                    Optional fechaFin As Date, _
                                    Optional idPersona As Long = 0, _
                                    Optional idOrganizacion As Long = 0, _
                                    Optional idOferta As Long = 0, _
                                    Optional idProyectoEmprendedor As Long = 0, _
                                    Optional idCurso As Long = 0, _
                                    Optional idIfocUsuario As Long = 0, _
                                    Optional order As Integer = 0) As String
    Dim strselect As String
    Dim strFrom As String
    Dim strWhere As String
    Dim strOrderBy As String
    Dim strSql As String
    
    Dim fechaI As Date
    Dim FECHAF As Date
    
    'Inicializamos variables
    strselect = " t_gestion.id," & _
                " 'Gestión' AS IoG," & _
                " a_gestiontipo.tipo AS Tipo," & _
                " Format(t_gestion.fecha,'dd/mm/yyyy') as Fecha," & _
                " Format(t_gestion.fecha, 'hh:nn') as Hora," & _
                " v_ifocusuario.aka AS [Técnic@]," & _
                " '' AS Acude," & _
                " '' AS Anula," & _
                " '' AS Citador," & _
                " a_servicio.aka AS Servicio," & _
                " format(t_gestion.fecha,'yyyy/mm/dd hh:mm') as fechaOrder"
                
    strFrom = "(((t_gestion LEFT JOIN r_gestionusuario ON t_gestion.id = r_gestionusuario.fkGestion)" & _
                " LEFT JOIN a_gestiontipo ON t_gestion.fkGestionTipo = a_gestiontipo.id)" & _
                " LEFT JOIN v_ifocusuario ON t_gestion.fkIfocUsuario = v_ifocusuario.fkPersona)" & _
                " LEFT JOIN a_servicio ON t_gestion.fkServicio = a_servicio.id"
    strWhere = ""
    strOrderBy = ""
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"
    
    If (fechaInicio <> "01/01/1900") Then
        strWhere = addConditionWhere(strWhere, "t_gestion.fecha Between #" & fechaI & "# AND #" & FECHAF & "#")
    End If
    
    If (idPersona <> 0) Then
        strWhere = addConditionWhere(strWhere, "r_gestionusuario.fkPersona = " & idPersona)
    End If
    If (idOrganizacion <> 0) Then
        strWhere = addConditionWhere(strWhere, "r_gestionusuario.fkOrganizacion = " & idOrganizacion)
    End If
    If (idOferta <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_gestion.fkOferta = " & idOferta)
    End If
    If (idProyectoEmprendedor <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_gestion.fkProyectoEmprendedor = " & idProyectoEmprendedor)
    End If
    If (idCurso <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_gestion.fkCurso = " & idCurso)
    End If
    If (idIfocUsuario <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_gestion.fkIfocUsuario = " & idIfocUsuario)
    End If
    
    If (order = 1) Then
        strOrderBy = "fecha ASC"
    ElseIf (order = 2) Then
        strOrderBy = "fecha DESC"
    End If
    
    strSql = montarSQL(strselect, _
                      strFrom, _
                      strWhere, _
                      , _
                      , _
                      strOrderBy)
Debug.Print strSql
    getSqlListadoGestiones = strSql
End Function

Public Function testGestion() As Long
    'creaGestionGrupal 1, 1, Now(), Time(), _
        "GESTION AUTOMATICA: Baja en servicios por caducidad. Datos de inserción no actualizados, de 2009 o antes. Si devuelve llamada citar a entrevista expres con documetación actualizada", _
        4480, , _
        "14505,4742,14455,18441,18596,1047,18629,649,17628,8132,9119,18402,18708,18707,13700,18730,18652,15416,3780,17643,18599,18772,13542,17898,18513,4293,18202,15407,18849,18831,17728,18808,18791,17892,18906,18952,18919,18701,14174,12296,17026,18737,1395,18401,18832,16315,17340,5640,19076,18769,18822,19046,3797,18716,15914,17165,17174,18760,1514,143,15098,12897,18051,12540,15530,12300,1807,14091,15159,18682,18838,16590,17930,18209,18413,11791,15450,18210,18642,7615,3601,18723,18905,17549,12149,18650,15415,18140,19383,4022,13032,16475,9697,18130,18114,18220,18467,18680,18783,18795,19229", _
        , , , , 1

    'bajaServiciosUsuarios Now(), 19, 4480, _
            "14505,4742,14455,18441,18596,1047,18629,649,17628,8132,9119,18402,18708,18707,13700,18730,18652,15416,3780,17643,18599,18772,13542,17898,18513,4293,18202,15407,18849,18831,17728,18808,18791,17892,18906,18952,18919,18701,14174,12296,17026,18737,1395,18401,18832,16315,17340,5640,19076,18769,18822,19046,3797,18716,15914,17165,17174,18760,1514,143,15098,12897,18051,12540,15530,12300,1807,14091,15159,18682,18838,16590,17930,18209,18413,11791,15450,18210,18642,7615,3601,18723,18905,17549,12149,18650,15415,18140,19383,4022,13032,16475,9697,18130,18114,18220,18467,18680,18783,18795,19229", _
            , , "Baja automática en servicios por caducidad."
    'Dim ids As String
    'ids = "3785,3818,3935,4017,4064,4065,4094,4101,4118,4143,4156,4167,4244,4256,4279,4313,4327,4347,4370,4372,4382,4423,4436,4445,4565,4636,4657,4724,4753,4774,4784,4792,4806,4822,4836,4850,4877,5075,5096,5102,5124,5146,5193,5213,5271,5273,5315,5322,5383,5403,5429,5596,5640,5649,5657,5689,16907,16946,16958,16964,16995,17030,17047,17076,17126,17151,17170,17184,17199,17248,17260,17277,17306,17325,17436,17450,17456,17457,17458,17465,17484,17500,17540,17550,17564,17587,17601,17615,17633,17643,17647,17648,17658,17663,17665,17666,17699,17701,17705,17711,17724,17737,30,43,89,101,154,192,260,378,393,437,491,563,577,587,614,664,708,788,804,985,1004,1016,1019,1036,1051,1137,1148,1159,1229,1247,1253,1404,1455,1479,1517,1533,1575,1585,1595,1628,1738,1806,1810,1835,1843,1872,1878,1880,1882,1905,1921,1958,1969,2015,2017,2052,2105,2116,2117,2127,2144,2163,2196,2259," & _
          "2321,2333,2371,2406,2422,2528,2553,2555,2572,2618,2749,2827,2840,2885,2897,3007,3024,3106,3111,3247,3251,3254,3255,3298,3322,3365,3376,3420,3431,3471,3482,3516,3542,3566,3596,3618,3624,3681,3718,5694,5713,5722,5726,5761,5784,5838,5862,5864,5865,5890,6065,6250,6259,6423,6452,6616,6684,6695,6702,6747,6773,6833,6853,6898,6916,6927,6964,6976,7033,7133,7138,7147,7165,7279,7370,7371,7393,7530,7588,7601,7620,7621,7670,7786,7913,8013,8042,8047,8055,8286,8302,8316,8335,8378,8428,8430,8581,8629,8776,8895,8925,8963,8981,9011,9117,9206,9440,9527,9547,9753,9812,10167,10409,10516,10818,10956,10979,11071,11126,11127,11165,11171,11172,11173,11179,11197,11208,11216,11217,11241,11257,11311,11357,11359,11379,11391,11400,11417,11437,11441,11512,11517,11532,11564,11599,11654,11659,11662,11675,11678,11679,11689,11690,11696,11701,11709,11711,11725,11729,11756,11766," & _
          "11767,11769,11771,11776,11777,11794,11797,11813,11819,11825,11826,11832,11838,11846,11872,11885,11888,12193,12194,12196,12198,12211,12224,12248,12299,12321,12325,12343,12347,12365,12374,12376,12382,12386,12393,12417,12483,12551,12553,12585,12587,12604,12606,12660,12676,12682,12711,12734,12741,12760,12779,12786,12882,12910,12932,12933,12983,12986,12987,12990,13006,13009,13046,13107,13132,13191,13241,13253,13261,13320,13440,13457,13461,13474,13527,13594,13679,13713,13718,13743,13753,13925,13937,13972,13988,14020,14026,14050,14056,14061,14062,14083,14091,14094,14098,14100,14108,14140,14147,14151,14156,14173,14174,14184,14230,14254,14271,14273,14280,14294,14301,14321,14328,14334,14342,14383,14394,14426,14433,14434,14435,14449,14454,14463,14474,14484,14497,14501,14507,14577,14585,14605,14620,14631,14632,14659,14660,14683,14684,14701,14728,14742,14755," & _
          "14771,14836,14906,14928,14938,14943,14954,14955,14991,15013,15122,15135,15320,15325,15327,15329,15359,15394,15425,15436,15506,15510,15589,15610,15625,15667,15703,15725,15806,15829,15831,15839,15861,15871,15890,15911,15915,15966,15977,16014,16031,16048,16075,16091,16112,16116,16130,16133,16138,16156,16161,16210,16229,16275,16296,16371,16442,16448,16449,16490,16523,16566,16583,16616,16619,16659,16680,16682,16705,16730,16740,16782,16824,16825,16887,16890,16891,16896,17761,17767,17799,17809,17815,17817,17821,17831,17834,17835,17846,17858,17859,17866,17888,17889,17907,17915,17919,17926,17952,17954,17960,17967,17979,17983,17992,18040,18051,18062,18076,18088,18107,18116,18122,18132,18183,18204,18251,18275,18278,18297,18331,18337,18369,18378,18384,18392,18403,18439,18466,18485,18488,18508,18523,18549,18559,18573,18586,18591,18592,18593,18595,18602,18607," & _
          "18623,18637,18695,18704,18739,18795,18807,18821,18854,18856,18857,18867,18879,18914,18919,18926,18975,18985,19044,19045,19059,19060,19069,19078,19090,19099,19114,19138,19142,19152,19153,19159,19161,19162,19164,19165,19186,19191,19216,19218,19251,19259,19260,19281,19307,19334,19363,19412,19450,19453,19459,19462,19473,19493,19516,19519,19537,19538,19559,19581,19587,19616,19621,19628,19641,19643,19661,19666,19720,19722,19733,19734,19739,19741,19742,19752,19767,19795,19804,19829,19864,19900,19904,19913,19926,19948,19965,19972,19984,20016,20025,20031,20070,20072,20101,20102,20119,20121,20123,20141,20144,20153,20171,20206,20222,20264,20272,20277,20309,20324,20356,20383"
    'Debug.Print creaGestionGrupal(5, 3, Now(), Now(), "Envío información de la jornada 'Fem Empresa, encuentro de emprendedores' organizada por Vicepresidencia económica del Gobern Balear.", 14327, , ids, , , , , 2)
    'Debug.Print insPersonasGestion(106853, "18854,18856,18857,18867,18879,18914,18919,18926,18975,18985,19044,19045,19059,19060,19069,19078,19090,19099,19114,19138,19142,19152,19153,19159,19161,19162,19164,19165,19186,19191,19216,19218,19251,19259,19260,19281,19307,19334,19363,19412,19450,19453,19459,19462,19473,19493,19516,19519,19537,19538,19559,19581,19587,19616,19621,19628,19641,19643,19661,19666,19720,19722,19733,19734,19739,19741,19742,19752,19767,19795,19804,19829,19864,19900,19904,19913,19926,19948,19965,19972,19984,20016,20025,20031,20070,20072,20101,20102,20119,20121,20123,20141,20144,20153,20171,20206,20222,20264,20272,20277,20309,20324,20356,20383", 14)
    Debug.Print creaGestionGrupal(1, 1, now(), now(), "prueba gestion", 14, , "14,1")
End Function
