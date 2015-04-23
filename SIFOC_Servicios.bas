Attribute VB_Name = "SIFOC_Servicios"
'--------------------------------------------------------------------------------------------
'   Name:   altaServicioUsuario
'   Autor:  Asunción Huertas - Upd: José Manuel Sanchez
'   Fecha:  22/03/2010   Actualización: 21/09/2011
'   Desc:   Da de alta una persona/empresa en un servicio con fecha inicio (obligatorio)
'           y resto de datos opcionales
'   Param:  idServicio
'           fechaI
'           idIfocUsuario
'           fechaF
'           idMotivoBaja
'           mejora
'           observacion
'           idPersona
'           idOrganizacion
'   Retur:   0, ok
'           -1, ko
'--------------------------------------------------------------------------------------------
Public Function altaServicioUsuario(idServicio As Long, _
                                    fechaI As Date, _
                                    idIfocUsuario As Long, _
                                    Optional FECHAF As Date = "01/01/1900 00:00:00", _
                                    Optional idMotivoBaja As Long = 0, _
                                    Optional mejora As Integer = 0, _
                                    Optional observacion As String = "", _
                                    Optional idPersona As Long = 0, _
                                    Optional idOrganizacion As Long = 0) As Integer
On Error GoTo Error
    
    Dim strSql As String
    Dim strFields As String
    Dim strValues As String
    Dim fechaInicio As Date
    Dim fechaFin As Date
    Dim OBS As String
    
    OBS = filterSQL(observacion)
    
    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        
        fechaInicio = Format(fechaI, "mm/dd/yyyy hh:nn:ss")
        fechaFin = Format(FECHAF, "mm/dd/yyyy hh:nn:ss")
        
        strFields = IIf((idPersona <> 0), "fkPersona", "fkOrganizacion") & _
                    ", fkServicio" & _
                    ", fechaInicio" & _
                    ", fkIfocUsuarioAlta" & _
                    IIf(idMotivoBaja = 0, "", ", fkMotivoBaja") & _
                    IIf(FECHAF = "01/01/1900 00:00:00", "", ", fechaFin") & _
                    IIf(idMotivoBaja = 0, "", ", fkIfocUsuarioBaja") & _
                    IIf(mejora = 0, "", ", mejora") & _
                    IIf(observacion = "", "", ", observacion ")
        strValues = IIf((idPersona <> 0), idPersona, idOrganizacion) & _
                    ", " & idServicio & _
                    ", #" & fechaInicio & "#" & _
                    ", " & idIfocUsuario & _
                    IIf(idMotivoBaja = 0, "", ", " & idMotivoBaja) & _
                    IIf(FECHAF = "01/01/1900 00:00:00", "", ", #" & fechaFin & "#") & _
                    IIf(idMotivoBaja = 0, "", ", " & idIfocUsuario) & _
                    IIf(mejora = 0, "", ", " & mejora) & _
                    IIf(observacion = "", "", ", """ & OBS & """")
        
        strSql = " INSERT INTO r_serviciousuario (" & strFields & ")" & _
                 " VALUES (" & strValues & ");"
    
'Debug.Print strSql
        'Debug.Print strSQL
        CurrentDb.Execute strSql
        altaServicioUsuario = 0
    Else
        MsgBox "", vbOKOnly, "Alert: SIFOC_Servicios"
    End If
    
SalirTratarError:
    Exit Function
Error:
    debugando "Error: " & Err.description
    altaServicioUsuario = -1
End Function

'---------------------------------------------------------------------------
'   Name:   bajaServicioUsuario
'   Autor:  Asunción Huertas  - Upd: José Manuel Sanchez
'   Fecha:  05/05/2010   Actualización: 11/06/2010
'   Desc:   Da de baja una persona/empresa en un servicio
'   Retur:   0, ok
'           -1, ko
'---------------------------------------------------------------------------
Public Function bajaServicioUsuario(idServicioUsuario As Long, _
                                    FECHAF As Date, _
                                    idMotivoBaja As Long, _
                                    mejora As Integer, _
                                    observacion As String) As Integer
On Error GoTo Error

    Dim strSql As String
    Dim fechaFin As Date

    If (idServicioUsuario <> 0) Then
        
        'Al dar de baja, se debe actualizar la fecha de baja, el motivo de baja y el usuario que realiza la baja
        fechaFin = Format(FECHAF, "mm/dd/yyyy hh:nn:ss")
        strSql = " UPDATE r_serviciousuario" & _
                 " SET fechaFin = #" & fechaFin & "#" & _
                 ", fkMotivoBaja = " & idMotivoBaja & _
                 ", mejora = " & mejora & _
                 ", observacion = '" & observacion & "'" & _
                 ", fkIfocUsuarioBaja = " & usuarioIFOC() & _
                 " WHERE id = " & idServicioUsuario & ";"
        
        'Debug.Print strSQL
        CurrentDb.Execute strSql
        bajaServicioUsuario = 0
    Else
        bajaServicioUsuario = -1
    End If

SalirTratarError:
    Exit Function
Error:
    debugando "Error: " & Err.description
    bajaServicioUsuario = -1
End Function

'---------------------------------------------------------------------------
'   Name:   bajaidServicioidUsuario
'   Autor:  José Manuel Sanchez  - Upd: José Manuel Sanchez
'   Fecha:  16/12/2011   Actualización: 16/12/2011
'   Desc:   Da de baja una persona/empresa en un servicio
'   Retur:   0, ok
'           -1, ko
'---------------------------------------------------------------------------
Public Function bajaidServicioidUsuario(idPersona As Long, _
                                        idServicio As Long, _
                                        FECHAF As Date, _
                                        idMotivoBaja As Long, _
                                        Optional observacion As String = "") As Integer
On Error GoTo Error

    Dim strSql As String
    Dim fechaFin As Date

    'Al dar de baja, se debe actualizar la fecha de baja, el motivo de baja y el usuario que realiza la baja
    fechaFin = Format(FECHAF, "mm/dd/yyyy hh:nn:ss")
    strSql = " UPDATE r_serviciousuario" & _
             " SET fechaFin = #" & fechaFin & "#" & _
             ", fkMotivoBaja = " & idMotivoBaja & _
             ", observacion = """ & observacion & """" & _
             ", fkIfocUsuarioBaja = " & U_idIfocUsuarioActivo & _
             " WHERE fkPersona = " & idPersona & " AND fkServicio = " & idServicio & _
                    " AND fechafin is null AND fechaInicio < #" & fechaFin & "#;"
    
'Debug.Print strSql
    CurrentDb.Execute strSql
    bajaidServicioidUsuario = 0

SalirTratarError:
    Exit Function
Error:
    Debug.Print "Error SIFOC_Servicios(bajaidServicioidUsuario): " & Err.description
    bajaidServicioidUsuario = -1
End Function

'---------------------------------------------------------------------------
'   Name:   bajaServiciosUsuario
'   Autor:  José Manuel Sanchez  - Upd: José Manuel Sanchez
'   Fecha:  03/08/2011   Actualización: 03/08/2011
'   Desc:   Da de baja una persona/empresa de todos sus servicios
'   Retur:   0, ok
'           -1, ko
'---------------------------------------------------------------------------
Public Function bajaServiciosUsuario(fechaFin As Date, _
                                     idMotivoBaja As Long, _
                                     idIfocUsuario As Long, _
                                     Optional idPersona As Long = 0, _
                                     Optional idOrganizacion As Long = 0, _
                                     Optional mejora As Integer = 0, _
                                     Optional observacion As String = "") As Integer
On Error GoTo Error

    Dim strSql As String
    Dim FECHAF As Date

    result = 0
    
    If (idPersona <> 0) Then
        
        'Al dar de baja, se debe actualizar la fecha de baja, el motivo de baja y el usuario que realiza la baja
        FECHAF = Format(fechaFin, "mm/dd/yyyy hh:nn:ss")
        strSql = " UPDATE r_serviciousuario" & _
                 " SET fechaFin = #" & FECHAF & "#" & _
                 ", fkMotivoBaja = " & idMotivoBaja & _
                 ", fkIfocUsuarioBaja = " & idIfocUsuario & _
                 IIf(mejora <> 0, ", mejora = " & mejora, "") & _
                 IIf(observacion <> "", ", observacion = '" & observacion & "'", "") & _
                 " WHERE fechaFin is null AND fkPersona =" & idPersona & ";"
        
        CurrentDb.Execute strSql
    ElseIf (idOrganizacion <> 0) Then
        
        'Al dar de baja, se debe actualizar la fecha de baja, el motivo de baja y el usuario que realiza la baja
        FECHAF = Format(fechaFin, "mm/dd/yyyy hh:nn:ss")
        strSql = " UPDATE r_serviciousuario" & _
                 " SET fechaFin = #" & FECHAF & "#" & _
                 ", fkMotivoBaja = " & idMotivoBaja & _
                 ", fkIfocUsuarioBaja = " & idIfocUsuario & _
                 IIf(mejora <> 0, ", mejora = " & mejora, "") & _
                 IIf(observacion <> "", ", observacion = '" & observacion & "'", "") & _
                 " WHERE fechaFin is null AND fkOrganizacion =" & idOrganizacion & ";"
        
        CurrentDb.Execute strSql
    Else
        result = -1
    End If

SalirTratarError:
    bajaServiciosUsuario = result
    Exit Function
Error:
    debugando "Error: " & Err.description
    bajaServiciosUsuario = -1
End Function

'---------------------------------------------------------------------------
'   Name:   bajaServiciosUsuarios
'   Autor:  José Manuel Sanchez  - Upd: José Manuel Sanchez
'   Fecha:  03/08/2011   Actualización: 03/08/2011
'   Desc:   Da de baja una persona/empresa de todos sus servicios
'   Retur:   0, ok
'           -1, ko
'---------------------------------------------------------------------------
Public Function bajaServiciosUsuarios(fechaFin As Date, _
                                      idMotivoBaja As Long, _
                                      idIfocUsuario As Long, _
                                      Optional Personas As String = "", _
                                      Optional Organizaciones As String = "", _
                                      Optional mejora As Integer = 0, _
                                      Optional observacion As String = "") As Integer
On Error GoTo Error

    Dim result As Integer

    result = 0
    
    Dim usrs
    Dim numUsr As Integer
    Dim idUsr
    
    If (Personas <> "") Then
        usrs = Split(Personas, ",")
        numUsr = UBound(usrs)
        For Each idUsr In usrs
            result = bajaServiciosUsuario(fechaFin, idMotivoBaja, idIfocUsuario, CLng(idUsr), , mejora, observacion)
        Next
    ElseIf (Organizaciones <> "") Then
        usrs = Split(Organizaciones, ",")
        numUsr = UBound(usrs)
        For Each idUsr In usrs
            result = bajaServiciosUsuario(fechaFin, idMotivoBaja, idIfocUsuario, , CLng(idUsr), mejora, observacion)
        Next
    Else
        result = -1
    End If

SalirTratarError:
    bajaServiciosUsuarios = result
    Exit Function
Error:
    debugando "Error: " & Err.description
    bajaServiciosUsuarios = -1
End Function

'---------------------------------------------------------------------------
'   Name:   actualizacionServicioUsuario
'   Autor:  Asunción Huertas
'   Fecha:  22/03/2010   Actualización: 05/05/2010
'   Desc:   Actualiza un servicio de alta en una persona/empresa
'   Retur:   0, ok
'           -1, ko
'---------------------------------------------------------------------------
Public Function actualizacionServicioUsuario(idServicioUsuario As Long, _
                                             fechaI As Date, _
                                             Optional mejora As Integer, _
                                             Optional observacion As String) As Integer
On Error GoTo Error

    Dim strSql As String
    Dim fechaInicio As Date

    If (idServicioUsuario <> 0) Then
        
        If (fechaI = "00:00:00") Then 'Si no se indica fecha inicio, se actualiza la mejora y la observacion
            strSql = " UPDATE r_serviciousuario" & _
                     " SET mejora = " & mejora & _
                     ", observacion = '" & observacion & "'" & _
                     " WHERE id = " & idServicioUsuario & ";"
        Else    'Si se indica fecha inicio, se actualiza solo la fecha inicio
            fechaInicio = Format(fechaI, "mm/dd/yyyy hh:nn:ss")
            strSql = " UPDATE r_serviciousuario" & _
                     " SET fechaInicio = #" & fechaInicio & "#" & _
                     " WHERE id = " & idServicioUsuario & ";"
        End If
        
        'Debug.Print strSQL
        CurrentDb.Execute strSql
        actualizacionServicioUsuario = 0
    Else
        actualizacionServicioUsuario = -1
    End If

    Exit Function

Error:
    debugando "Error: " & Err.description
    actualizacionServicioUsuario = -1
End Function

'--------------------------------------------------------------------------------------------
'   Name:   eliminaServicioUsuario
'   Autor:  Asunción Huertas
'   Fecha:  22/03/2010   Actualización:
'   Desc:   Elimina el servicio de una persona/empresa
'   Retur:   0, ok
'           -1, ko
'--------------------------------------------------------------------------------------------
Public Function eliminaServicioUsuario(idServicioUsuario As Long) As Integer
On Error GoTo TratarError
    
    Dim strSql As String
    
    strSql = " DELETE * FROM r_serviciousuario" & _
             " WHERE (id = " & idServicioUsuario & ");"
    
    'Debug.Print strSQL
    CurrentDb.Execute strSql
    eliminaServicioUsuario = 0

SalirTratarError:
    Exit Function
TratarError:
    debugando "Error: " & Err.description
    eliminaServicioUsuario = -1
End Function

'---------------------------------------------------------------------------
'   Name:   numServiciosUsuarioActivosEnAmbitoEnFecha
'   Autor:  Asunción Huertas
'   Fecha:  11/03/2010
'   Desc:   Indica el número de servicios activos en la fecha indicada
'           de la persona o empresa en el ámbito indicado o en todos,
'           si no se pasa ningún ámbito
'   Param:  fecha de actividad(date),
'           idAmbito(integer), identificador del ámbito(OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Número de servicios activos del usuario en el ámbito y en la fecha indicada
'           -1, ko
'---------------------------------------------------------------------------
Public Function numServiciosUsuarioActivosEnAmbitoEnFecha(fecha As Date, _
                                                Optional idAmbito As Integer = 0, _
                                                Optional idPersona As Long = 0, _
                                                Optional idOrganizacion As Long = 0) As Integer
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    Dim numServiciosActivos As Integer
    
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        
        strSql = " SELECT DISTINCT " & _
                 IIf((idPersona = 0), "", " r_serviciousuario.fkPersona, r_serviciousuario.fkServicio") & _
                 IIf((idOrganizacion = 0), "", " r_serviciousuario.fkOrganizacion, r_serviciousuario.fkServicio") & _
                 " FROM r_serviciousuario LEFT JOIN a_servicio ON r_serviciousuario.fkServicio=a_servicio.id" & _
                 " WHERE " & _
                    IIf((idPersona = 0), "", " ((r_serviciousuario.fkPersona = " & idPersona & ")") & _
                    IIf((idOrganizacion = 0), "", "((r_serviciousuario.fkOrganizacion = " & idOrganizacion & ")") & _
                 " AND ((r_serviciousuario.fechaInicio) <= #" & fechaActividad & "#)" & _
                 " AND (((r_serviciousuario.fechaFin) > #" & fechaActividad & "#) OR ((r_serviciousuario.fechaFin) Is Null))" & _
                 IIf((idAmbito = 0), "", " AND (a_servicio.fkIfocAmbito = " & idAmbito & ")") & _
                 ");"
               
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        numServiciosActivos = rs.RecordCount
            
        rs.Close
        Set rs = Nothing
        
        numServiciosUsuarioActivosEnAmbitoEnFecha = numServiciosActivos
    Else
        numServiciosUsuarioActivosEnAmbitoEnFecha = -1
    End If
    
    Exit Function
       
Error:
    Debug.Print "Error: " & Err.description
    numServiciosUsuarioActivosEnAmbitoEnFecha = -1
End Function

'-------------------------------------------------------------------------------------------------
'   Name:   numServiciosUsuarioFuturosEnAmbitoEnFecha
'   Autor:  Asunción Huertas
'   Fecha:  04/05/2010
'   Desc:   Indica el número de servicios que estarán activos posteriormente a la fecha indicada
'           de la persona o empresa en el ámbito indicado o en todos, si no se pasa ningún ámbito
'   Param:  fecha de actividad(date),
'           idAmbito(integer), identificador del ámbito(OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Número de servicios activos del usuario en el ámbito con fecha posterior a la fecha indicada
'           (servicios futuros)
'           -1, ko
'--------------------------------------------------------------------------------------------------
Public Function numServiciosUsuarioFuturosEnAmbitoEnFecha(fecha As Date, _
                                                Optional idAmbito As Integer = 0, _
                                                Optional idPersona As Long = 0, _
                                                Optional idOrganizacion As Long = 0) As Integer
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    Dim numServiciosFuturos As Integer
    
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        
        strSql = " SELECT DISTINCT " & _
                 IIf((idPersona = 0), "", " r_serviciousuario.fkPersona, r_serviciousuario.fkServicio") & _
                 IIf((idOrganizacion = 0), "", " r_serviciousuario.fkOrganizacion, r_serviciousuario.fkServicio") & _
                 " FROM r_serviciousuario LEFT JOIN a_servicio ON r_serviciousuario.fkServicio=a_servicio.id" & _
                 " WHERE " & _
                    IIf((idPersona = 0), "", " ((r_serviciousuario.fkPersona = " & idPersona & ")") & _
                    IIf((idOrganizacion = 0), "", "((r_serviciousuario.fkOrganizacion = " & idOrganizacion & ")") & _
                 " AND ((r_serviciousuario.fechaInicio) > #" & fechaActividad & "#)" & _
                 " AND (((r_serviciousuario.fechaFin) > #" & fechaActividad & "#) OR ((r_serviciousuario.fechaFin) Is Null))" & _
                 IIf((idAmbito = 0), "", " AND (a_servicio.fkIfocAmbito = " & idAmbito & ")") & _
                 ");"
               
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        numServiciosFuturos = rs.RecordCount
            
        rs.Close
        Set rs = Nothing
        
        numServiciosUsuarioFuturosEnAmbitoEnFecha = numServiciosFuturos
    Else
        numServiciosUsuarioFuturosEnAmbitoEnFecha = -1
    End If
    
    Exit Function
       
Error:
    debugando "Error: " & Err.description
    numServiciosUsuarioFuturosEnAmbitoEnFecha = -1
End Function

'---------------------------------------------------------------------------
'   Name:   numServiciosUsuarioActivosEnFecha
'   Autor:  Asunción Huertas
'   Fecha:  24/03/2010
'   Desc:   Si se le pasa un servicio, indica si el usuario está de alta
'           en el servicio indicado en la fecha indicada
'           Si no se le pasa un servicio, indica si el usuario está de alta
'           en algún servicio en la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio (OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Número de servicios activos del usuario en la fecha indicada
'           -1, ko
'---------------------------------------------------------------------------
Public Function numServiciosUsuarioActivosEnFecha(fecha As Date, _
                                        Optional idServicio As Long = 0, _
                                        Optional idPersona As Long = 0, _
                                        Optional idOrganizacion As Long = 0) As Integer
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    Dim numServiciosActivos As Integer
    
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT " & _
                 IIf((idPersona = 0), "", " fkPersona") & _
                 IIf((idOrganizacion = 0), "", " fkOrganizacion") & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 " AND ((fechaInicio) <= #" & fechaActividad & "#)" & _
                 " AND (((fechaFin) > #" & fechaActividad & "#) OR ((fechaFin) Is Null))" & _
                 ");"
    
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        numServiciosActivos = rs.RecordCount
            
        rs.Close
        Set rs = Nothing
        
        numServiciosUsuarioActivosEnFecha = numServiciosActivos
    Else
        numServiciosUsuarioActivosEnFecha = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    numServiciosUsuarioActivosEnFecha = -1
End Function

'----------------------------------------------------------------------------------------------
'   Name:   numServiciosUsuarioFuturosEnFecha
'   Autor:  Asunción Huertas
'   Fecha:  04/05/2010
'   Desc:   Si se le pasa un servicio, indica si el usuario estará de alta en el servicio
'           posteriormente a la fecha indicada
'           Si no se le pasa un servicio, indica si el usuario estará de alta en algún servicio
'           posteriormente a la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio (OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Número de servicios activos del usuario con fecha posterior a la fecha indicada
'           (servicios futuros)
'           -1, ko
'----------------------------------------------------------------------------------------------
Public Function numServiciosUsuarioFuturosEnFecha(fecha As Date, _
                                        Optional idServicio As Long = 0, _
                                        Optional idPersona As Long = 0, _
                                        Optional idOrganizacion As Long = 0) As Integer
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    Dim numServiciosFuturos As Integer
    
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT " & _
                 IIf((idPersona = 0), "", " fkPersona") & _
                 IIf((idOrganizacion = 0), "", " fkOrganizacion") & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 " AND ((fechaInicio) > #" & fechaActividad & "#)" & _
                 " AND (((fechaFin) > #" & fechaActividad & "#) OR ((fechaFin) Is Null))" & _
                 ");"
    
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        numServiciosFuturos = rs.RecordCount
            
        rs.Close
        Set rs = Nothing
        
        numServiciosUsuarioFuturosEnFecha = numServiciosFuturos
    Else
        numServiciosUsuarioFuturosEnFecha = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    numServiciosUsuarioFuturosEnFecha = -1
End Function

'------------------------------------------------------------------------------------------------
'   Name:   servicioUsuarioActivoEnFecha
'   Autor:  Asunción Huertas
'   Fecha:  24/03/2010
'   Desc:   Si se le pasa un servicio, indica el id de la relación del usuario
'           que está de alta en el servicio en la fecha indicada
'           Si no se le pasa un servicio, indica el id de la primera relación del usuario que
'           está de alta en algún servicio en la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio (OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  id de la relación del usuario que está activo en el servicio en la fecha indicada
'           -1, ko
'------------------------------------------------------------------------------------------------
Public Function servicioUsuarioActivoEnFecha(fecha As Date, _
                                    Optional idServicio As Long = 0, _
                                    Optional idPersona As Long = 0, _
                                    Optional idOrganizacion As Long = 0) As Long
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")
    
    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT id" & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 " AND ((fechaInicio) <= #" & fechaActividad & "#)" & _
                 " AND (((fechaFin) > #" & fechaActividad & "#) OR ((fechaFin) Is Null))" & _
                 ");"
        
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        If Not rs.EOF Then
            rs.MoveFirst
            servicioUsuarioActivoEnFecha = rs!id
        Else
            servicioUsuarioActivoEnFecha = -1
        End If
            
        rs.Close
        Set rs = Nothing
    Else
        servicioUsuarioActivoEnFecha = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    servicioUsuarioActivoEnFecha = -1
End Function

'-------------------------------------------------------------------------------------------
'   Name:   servicioUsuarioFuturoEnFecha
'   Autor:  Asunción Huertas
'   Fecha:  04/05/2010
'   Desc:   Si se le pasa un servicio, indica el id de la relación del usuario
'           que estará de alta en el servicio posteriormente a la fecha indicada
'           Si no se le pasa un servicio, indica el id de la primera relación del usuario que
'           estará de alta en algún servicio posteriormente a la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio (OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  id de la relación del usuario que estará activo en el servicio posteriormente a la fecha indicada
'           -1, ko
'-------------------------------------------------------------------------------------------
Public Function servicioUsuarioFuturoEnFecha(fecha As Date, _
                                    Optional idServicio As Long = 0, _
                                    Optional idPersona As Long = 0, _
                                    Optional idOrganizacion As Long = 0) As Long
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")
    
    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT id" & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 " AND ((fechaInicio) > #" & fechaActividad & "#)" & _
                 " AND (((fechaFin) > #" & fechaActividad & "#) OR ((fechaFin) Is Null))" & _
                 ");"
        
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        If Not rs.EOF Then
            rs.MoveFirst
            servicioUsuarioFuturoEnFecha = rs!id
        Else
            servicioUsuarioFuturoEnFecha = -1
        End If
            
        rs.Close
        Set rs = Nothing
    Else
        servicioUsuarioFuturoEnFecha = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    servicioUsuarioFuturoEnFecha = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  13/08/2009   Actualización: Asunción Huertas 04/05/2010
'   Name:   ListaServiciosActivos
'   Desc:   Devuelve un string con los servicios activos del usuario
'   Param:  idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Retur:  Devuelve un string con los servicios activos del usuario
'---------------------------------------------------------------------------
Public Function listaServiciosActivos(Optional idPersona As Long = 0, _
                                      Optional idOrganizacion As Long = 0) As String
On Error GoTo Error
    
    Dim rs As ADODB.Recordset
    Dim strSql As String
    Dim Servicios As String
    Dim ahora As String
    
    ahora = now 'Format(Now, "mm/dd/yyyy hh:nn:ss")
    
    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT a_servicio.aka" & _
                 " FROM r_serviciousuario LEFT JOIN a_servicio ON r_serviciousuario.fkServicio = a_servicio.id" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (r_serviciousuario.fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (r_serviciousuario.fkOrganizacion=" & idOrganizacion & ")") & _
                 " AND ((r_serviciousuario.fechaFin > #" & ahora & "#) Or ((r_serviciousuario.fechaFin) Is Null)))" & _
                 " ORDER BY a_servicio.aka;"
    
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
        Servicios = ""
        If Not rs.EOF Then
            rs.MoveFirst
            While Not rs.EOF
                If Len(Servicios) = 0 Then
                    Servicios = rs!aka
                Else
                    Servicios = Servicios & ", " & rs!aka
                End If
                rs.MoveNext
            Wend
        Else
            Servicios = "Ninguno"
        End If
    
        listaServiciosActivos = Servicios
    
        rs.Close
        Set rs = Nothing
    Else
        listaServiciosActivos = ""
    End If
    
    Exit Function
       
Error:
    debugando "Error: " & Err.description
    listaServiciosActivos = ""
End Function

'--------------------------------------------------------------------------------------------
'   Author: José Manuel Sánchez Báez
'   Fecha:  27/02/2008   Actualización: 26/03/2010 Asunción Huertas
'   Name:   ActivaServicioFormacion
'   Desc:   Da de alta en servicio de Formación
'   Param:  persona -> busca altas de ese id de persona.
'           ifocUsuario -> usuario de ifoc que da el alta(téc o aux)
'           [fechaInicio] -> fecha de la baja de la acción
'           [fechaFin] -> fecha de la baja de la acción
'   Return: 0 -> ok
'           -1 -> ko
'--------------------------------------------------------------------------------------------
Public Function ActivaServicioFormacion(idPersona As Long, _
                                        idIfocUsuario As Long, _
                                        fechaInicio As Date, _
                                        Optional fechaFin As Date, _
                                        Optional motivobaja As Long = 9, _
                                        Optional observacion As String) As Integer
On Error GoTo Error
    
    Dim altasServicio As Integer
    Dim idServicioUsuarioActivo As Long
    Dim resultado As Integer
    
    If IsMissing(fechaFin) Then
        fechaFin = fechaInicio + 365
    End If
    
    'Comprobamos si la persona tiene activo el servicio FORMACION
    altasServicio = numServiciosUsuarioActivosEnFecha(fechaInicio, 23, idPersona)

    Select Case altasServicio
        Case -1 'Error
            resultado = -1
        Case 0  'Alta y baja en servicio Formación (Fechas del curso)
            If altaServicioUsuario(23, fechaInicio, idIfocUsuario, fechaFin, motivobaja, 0, "", idPersona) = 0 Then
                MsgBox "El alta en el servicio de formación se ha realizado correctamente", vbOKOnly, "Alert: SIFOC_Servicios"
            Else
                MsgBox "El alta en el servicio de formación no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
            End If
        Case Else 'Baja del servicio Formación activo (Fecha Fin del curso)
            idServicioUsuarioActivo = servicioUsuarioActivoEnFecha(fechaInicio, 23, idPersona)
            If servicioActivo <> -1 Then
                If bajaServicioUsuario(idServicioUsuarioActivo, fechaFin, motivobaja, 0, "") = 0 Then
                    MsgBox "La actualización del servicio de formación se ha realizado correctamente", vbOKOnly, "Alert: SIFOC_Servicios"
                Else
                    MsgBox "La actualización del servicio de formación no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
                End If
            Else
                MsgBox "La actualización del servicio de formación no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
            End If
    End Select
    
    ActivaServicioFormacion = 0
    Exit Function
       
Error:
    Debug.Print "Error: " & Err.description
    ActivaServicioFormacion = -1
End Function

'---------------------------------------------------------------------------
'   Name:   updServicioUsuario
'   Autor:  José Manuel Sánchez - Actualiza:  José Manuel Sánchez
'   Fecha:  28/05/2010 - Actualización: 28/05/2010
'   Desc:   Actualiza un servicio de alta en una persona/empresa
'   Param:  *idServicioUsuario(long), id registro de alta en tabla
'           *fechaI(date)
'           fechaF(date)
'           idServicio(long), servicio en el que se realiza el alta,baja
'           observacion(string)
'           idMotivoBaja(long), identificador de motivo baja servicio
'           mejora(integer), 0 - en paro, -1 - en mejora
'   Retur:   0, ok
'           -1, ko
'---------------------------------------------------------------------------
Public Function updServicioUsuario(idServicioUsuario As Long, _
                                   fechaI As Date, _
                                   Optional FECHAF As Date = 0, _
                                   Optional idMotivoBaja As Long = 0, _
                                   Optional mejora As Integer = 0, _
                                   Optional observacion As String = "") As Integer
On Error GoTo TratarError
    Dim resultado As Integer
    Dim strSql As String
    Dim fechaInicio As Date
    Dim fechaFin As Date
    
    fechaInicio = Format(fechaI, "mm/dd/yyyy hh:nn:ss")
    fechaFin = Format(FECHAF, "mm/dd/yyyy hh:nn:ss")
    
    resultado = -1
    
    If (idServicioUsuario <> 0) Then 'Es una modificación o baja del servicio ya creado
        strSql = " UPDATE r_serviciousuario SET " & _
                 " fkIfocUsuarioAlta = " & usuarioIFOC() & _
                 ", fechaInicio = #" & fechaInicio & "#" & _
                 IIf(FECHAF = "01/01/1900 00:00:00", ", fechaFin = null", ", fechaFin = #" & fechaFin & "#") & _
                 IIf(idMotivoBaja = 0, ", fkMotivoBaja = null", ", fkMotivoBaja = " & idMotivoBaja) & _
                 IIf(idMotivoBaja = 0, "", ", fkIfocUsuarioBaja = " & usuarioIFOC()) & _
                 IIf(mejora = 0, "", ", mejora = " & mejora) & _
                 ", observacion = '" & observacion & "'" & _
                 " WHERE id = " & idServicioUsuario & ";"
        'Debug.Print strSql
        CurrentDb.Execute strSql
        resultado = 0
    End If

Salir:
    updServicioUsuario = resultado
    Exit Function
TratarError:
    debugando "Error: " & Err.description
    updServicioUsuario = -1
End Function


'-------------------------------------------------------------------------
'   Name: controlTR_altaServicio
'   Autor: Asunción Huertas - upd: José Manuel Sanchez
'   Fecha: 25/03/2010 - Update: 10/06/2010
'   Desc: Revisa que tenga asignado TR en el ámbito del nuevo servicio
'         y si no tiene propone la asignación del usuario actual
'   Param:  *idServicio(long),
'           *fecha(date),
'           idPersona
'-------------------------------------------------------------------------
Public Sub controlTR_altaServicio(idServicio As Long, _
                                  fecha As Date, _
                                  Optional idPersona As Long = 0, _
                                  Optional idOrganizacion As Long = 0)
    Dim idIfocAmbito As Integer
    Dim nombreIfocAmbito As String
    Dim numTRPersona As Long
    Dim respuesta

    idIfocAmbito = Nz(DLookup("[fkIfocAmbito]", "[a_servicio]", "[id] = " & idServicio), -1)
    nombreIfocAmbito = Nz(DLookup("[ambito]", "[a_ifocambito]", "[id] = " & idIfocAmbito), "")
                          
    If idIfocAmbito = -1 Then
        MsgBox "El control de TR asignado no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
    Else
        numTRPersona = numTecnicosAsignadosAUsuario(fecha, idPersona, idOrganizacion)
        If numTRPersona = 0 Then
            If idPersona > 0 Then
                respuesta = MsgBox("Esta persona no tiene asignado TR en el ámbito " & nombreIfocAmbito & vbNewLine & _
                                    "" & vbNewLine & _
                                    "¿Deseas asignarte como TR?", vbYesNo, "Alert: SIFOC_Servicios")
                If respuesta = vbYes Then
                    If asignaTecnicoAPersona(idPersona, U_idIfocUsuarioActivo, fecha) = -1 Then
                        MsgBox "La asignación de TR no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
                    Else
                        MsgBox "La asignación de TR se ha realizado correctamente" & vbNewLine & _
                        "con fecha inicio " & Format(fecha, "dd/mm/yyyy"), vbOKOnly, "Alert: SIFOC_Servicios"
                    End If
                End If
            Else
                respuesta = MsgBox("Esta empresa no tiene asignado TR en el ámbito " & nombreIfocAmbito & vbNewLine & _
                                    "" & vbNewLine & _
                                    "¿Deseas asignarte como TR?", vbYesNo, "Alert: SIFOC_Servicios")
                If respuesta = vbYes Then
                    If asignaTecnicoAEmpresa(idOrganizacion, U_idIfocUsuarioActivo, fecha) = -1 Then
                        MsgBox "La asignación de TR no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
                    Else
                        MsgBox "La asignación de TR se ha realizado correctamente" & vbNewLine & _
                        "con fecha inicio " & Format(fecha, "dd/mm/yyyy"), vbOKOnly, "Alert: SIFOC_Servicios"
                    End If
                End If
            End If
        ElseIf numTRPersona = -1 Then
           MsgBox "El control de TR asignado no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
        End If
    End If
End Sub

'-------------------------------------------------------------------------------------
'   Name: controlTR_bajaServicio
'   Autor: Asunción Huertas - upd: José Manuel Sanchez
'   Fecha: 25/03/2010 - Update: 10/06/2010
'   Desc: Revisa que tenga otros servicios activos o futuros en el ámbito del
'         servicio eliminado y si no tiene, propone la desasignación del TR del ámbito

'-------------------------------------------------------------------------------------
Public Sub controlTR_bajaServicio(idServicio As Long, _
                                  fecha As Date, _
                                  Optional idPersona As Long = 0, _
                                  Optional idOrganizacion As Long = 0)
                            
    Dim idIfocAmbito As Integer
    Dim nombreIfocAmbito As String
    Dim idIfocUsuario As Long
    Dim respuesta

    idIfocAmbito = Nz(DLookup("[fkIfocAmbito]", "[a_servicio]", "[id] = " & idServicio), -1)
    nombreIfocAmbito = Nz(DLookup("[ambito]", "[a_ifocambito]", "[id] = " & idIfocAmbito), "")
                          
    If idIfocAmbito = -1 Then
        MsgBox "El control de TR asignado no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
    Else
        'Si la Persona/Empresa no tiene otros servicios activos o futuros en el ámbito del servicio eliminado
        If numServiciosUsuarioActivosEnAmbitoEnFecha(fecha, idIfocAmbito, idPersona, idOrganizacion) = 0 And _
           numServiciosUsuarioFuturosEnAmbitoEnFecha(fecha, idIfocAmbito, idPersona, idOrganizacion) = 0 Then
            Select Case numTecnicosAsignadosAUsuario(fecha, idPersona, idOrganizacion)
                Case -1
                    MsgBox "El control de TR asignado no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Servicios"
                Case 0    'Persona/Empresa sin TR asignado en el ámbito del servicio eliminado, no hacer nada
                Case Else 'Persona/Empresa con TR asignado en el ámbito del servicio eliminado, se propone desasignar el TR activo
                    If idPersona > 0 Then
                        respuesta = MsgBox("La persona no está de alta en ningún servicio del ámbito " & nombreIfocAmbito & " en fecha " & Format(fecha, "dd/mm/yyyy") & vbNewLine & _
                                           "¿Desea dar de baja al TR actual? ", vbYesNo, "Alert: SIFOC_Servicios")
                        If respuesta = vbYes Then
                            idIfocUsuario = tecnicoAsignadoAUsuario(fecha, idPersona, idOrganizacion)
                            If bajaPersonaDeTecnico(idPersona, idIfocUsuario, fecha) = 0 Then
                                MsgBox "La baja del TR se ha realizado correctamente" & vbNewLine & _
                                "con fecha " & Format(fecha, "dd/mm/yyyy"), vbOKOnly, "Alert: SIFOC_Servicios"
                            Else
                                MsgBox "La baja del TR no se ha podido realizar" & vbNewLine & _
                                       "" & vbNewLine & _
                                       "La persona tiene TR pero no está de alta en ningún servicio del ámbito " & nombreIfocAmbito, vbOKOnly, "Alert: SIFOC_Servicios"
                            End If
                        Else
                            MsgBox "La persona tiene TR pero no está de alta en ningún servicio del ámbito " & nombreIfocAmbito, vbOKOnly, "Alert: SIFOC_Servicios"
                        End If
                    Else
                        respuesta = MsgBox("La empresa no está de alta en ningún servicio del ámbito " & nombreIfocAmbito & " en fecha " & Format(fecha, "dd/mm/yyyy") & vbNewLine & _
                                           "¿Desea dar de baja al TR actual? ", vbYesNo, "Alert: SIFOC_Servicios")
                        If respuesta = vbYes Then
                            idIfocUsuario = tecnicoAsignadoAUsuario(fecha, idPersona, idOrganizacion)
                            If bajaEmpresaDeTecnico(idOrganizacion, idIfocUsuario, fecha) = 0 Then
                                MsgBox "La baja del TR se ha realizado correctamente" & vbNewLine & _
                                "con fecha " & Format(fecha, "dd/mm/yyyy"), vbOKOnly, "Alert: SIFOC_Servicios"
                            Else
                                MsgBox "La baja del TR no se ha podido realizar" & vbNewLine & _
                                       "" & vbNewLine & _
                                       "La empresa tiene TR pero no está de alta en ningún servicio del ámbito " & nombreIfocAmbito, vbOKOnly, "Alert: SIFOC_Servicios"
                            End If
                        Else
                            MsgBox "La empresa tiene TR pero no está de alta en ningún servicio del ámbito " & nombreIfocAmbito, vbOKOnly, "Alert: SIFOC_Servicios"
                        End If
                    End If
            End Select
        End If
    End If
End Sub

'--------------------------------------------------------------------------------------------------
'   Name: controlServicioSolicitudes
'   Autor: Asunción Huertas
'   Fecha: 20/04/2010
'   Desc: Controla que si la empresa tiene solicitudes en curso tenga el servicio indicado activo
'         y si todas están finalizadas o no tiene, tampoco esté activo el servicio
'--------------------------------------------------------------------------------------------------
Public Sub controlServicioSolicitudes(idEmpresa As Long, _
                                      idServicio As Long, _
                                      numSolicitudesEnCurso As Integer)

    Dim idServicioUsuarioActivo As Long
    Dim nombreServicio As String
    Dim mejoraServicio As Integer
    Dim observacionServicio As String
    Dim idIfocUsuario As Long
        
    nombreServicio = Nz(DLookup("[servicio]", "a_servicio", "[id]=" & idServicio), 0)

    If numSolicitudesEnCurso = 0 Then
        'Si la empresa tiene activo el servicio indicado, se da de baja del servicio
        If numServiciosUsuarioActivosEnFecha(now(), idServicio, , idEmpresa) > 0 Then
            MsgBox "La empresa no tiene solicitudes en curso y" & vbNewLine & _
                   "está de alta en el servicio " & nombreServicio & "." & vbNewLine & _
                   "Se va a dar de baja en dicho servicio.", vbOKOnly, "Alert: SIFOC_Servicios"
            
            idServicioUsuarioActivo = servicioUsuarioActivoEnFecha(now(), idServicio, , idEmpresa)
            If idServicioUsuarioActivo <> -1 Then
                mejoraServicio = Nz(DLookup("[mejora]", "r_serviciousuario", "[id]=" & idServicioUsuarioActivo), 0)
                observacionServicio = (Nz(DLookup("[observacion]", "r_serviciousuario", "[id]=" & idServicioUsuarioActivo), 0)) & "/Baja automática por finalización de solicitudes"
                If bajaServicioUsuario(idServicioUsuarioActivo, now(), 17, mejoraServicio, observacionServicio) = 0 Then
                    MsgBox "La baja en el servicio  " & nombreServicio & " se ha realizado correctamente", vbOKOnly, "Alert: SIFOC_Servicios"
                    creaGestionGrupal 4, _
                                      2, _
                                      now(), _
                                      "Baja el " & Format(now(), "dd/mm/yyyy") & " en servicio de " & nombreServicio & " por finalización de solicitudes", _
                                      U_idIfocUsuarioActivo, _
                                      idServicio, _
                                      , _
                                      CStr(idEmpresa)
                    'Control de TR en ámbito del servicio eliminado
                    controlTR_bajaServicio idServicio, now(), , idEmpresa
                Else
                    MsgBox "La baja en el servicio " & nombreServicio & " no se ha podido realizar" & vbNewLine & _
                           "La empresa no tiene solicitudes en curso pero sigue de alta en el servicio", vbOKOnly, "Alert: SIFOC_Servicios"
                End If
            Else
                MsgBox "La baja en el servicio " & nombreServicio & " no se ha podido realizar" & vbNewLine & _
                       "La empresa no tiene solicitudes en curso pero sigue de alta en el servicio", vbOKOnly, "Alert: SIFOC_Servicios"
            End If
        End If
    Else
        'Si la empresa no tiene activo el servicio indicado, se da de alta en el servicio
        If numServiciosUsuarioActivosEnFecha(now(), idServicio, , idEmpresa) = 0 Then
            MsgBox "La empresa tiene solicitudes en curso y" & vbNewLine & _
                   "no está de alta en el servicio " & nombreServicio & "." & vbNewLine & _
                   "Se va a dar de alta en dicho servicio.", vbOKOnly, "Alert: SIFOC_Servicios"
                   
            If altaServicioUsuario(idServicio, now(), idIfocUsuario, "01/01/1900 00:00:00", 0, 0, "", , idEmpresa) = 0 Then
                MsgBox "El alta en el servicio " & nombreServicio & " se ha realizado correctamente", vbOKOnly, "Alert: SIFOC_Servicios"
                'Control de TR en ámbito del nuevo servicio
                controlTR_altaServicio idServicio, now(), , idEmpresa
            Else
                MsgBox "El alta en el servicio " & nombreServicio & " no se ha podido realizar" & vbNewLine & _
                       "La empresa tiene solicitudes en curso pero no está de alta en el servicio", vbOKOnly, "Alert: SIFOC_Servicios"
            End If
        End If
    End If

End Sub

'-------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez - Update: Jose Manuel Sanchez
'   Fecha:  27/05/2010 - Actualización:  27/05/2010
'   Name:   esServicioModificable
'   Desc:   Nos indica si el servicio de la persona se puede
'           modificar/eliminar por el técnico TR
'   Param:  idPersona(long), identificador de idPersona
'           idIfocUsuario(long), identificador de idIfocUsuario
'   Retur:  True,   servicio modificable o eliminable
'           False,  servicio NO modificable NI eliminable
'-------------------------------------------------------------------
Public Function esServicioModificable(idPersona As Long, _
                                      idIfocUsuario As Long) As Boolean
On Error GoTo TratarError
    Dim resultado As Boolean
    
    If isTRDePersona(idIfocUsuario, idPersona) Then
        resultado = True
    Else
        resultado = False
    End If
    
    esServicioModificable = resultado
SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error: " & Err.description
    esServicioModificable = False
End Function

'---------------------------------------------------------------------------
'   Name:   isUserServiceActive()
'   Autor:  Jose Manuel Sanchez
'   Fecha:  12/02/2014 - Upd: 28/08/2014
'   Desc:   Devuleve cierto si ha estado activo en el servicio pasado por
'           parámetros en la fecha indicada
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Boolean, si está de alta en el intervalo
'---------------------------------------------------------------------------
Public Function isUserServiceActive(fecha As Date, _
                                    idServicio As Long, _
                                    Optional idPersona As Long = 0, _
                                    Optional idOrganizacion As Long = 0) As Boolean
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    isUserServiceActive = 0
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT " & _
                 IIf((idPersona = 0), "", " fkPersona") & _
                 IIf((idOrganizacion = 0), "", " fkOrganizacion") & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 IIf((idServicioUsuario = 0), "", " AND (id<>" & idServicioUsuario & ")") & _
                 " AND ((fechaInicio) <= #" & fechaActividad & "#)" & _
                 " AND (((fechaFin) > #" & fechaActividad & "#) OR ((fechaFin) Is Null))" & _
                 ")"
    
Debug.Print strSql
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        If (rs.RecordCount > 0) Then
            isUserServiceActive = True
        End If
        
        rs.Close
        Set rs = Nothing
        
    End If
    
    Exit Function
    
Error:
    Debug.Print "Error: " & Err.description
    isUserServiceActive = False
End Function

'---------------------------------------------------------------------------
'   Name:   isUserServiceActiveInInterval()
'   Autor:  Jose Manuel Sanchez
'   Fecha:  12/02/2014 - Upd: 28/08/2014
'   Desc:   Devuleve cierto si ha estado activo en el servicio pasado
'           por parámetros en algún momento del intervalo
'   Param:  fechaInicio(date)
'           fechaFin(date)
'           idServicio(long), identificador de servicio
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Boolean, true si está de alta en el intervalo
'---------------------------------------------------------------------------
Public Function isUserServiceActiveInInterval(fechaInicio As Date, _
                                              fechaFin As Date, _
                                              idServicio As Long, _
                                              Optional idPersona As Long = 0, _
                                              Optional idOrganizacion As Long = 0) As Boolean
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    isUserServiceActiveInInterval = False
    'fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT " & _
                 IIf((idPersona = 0), "", " fkPersona") & _
                 IIf((idOrganizacion = 0), "", " fkOrganizacion") & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 IIf((idServicioUsuario = 0), "", " AND (id<>" & idServicioUsuario & ")") & _
                 " AND ((fechaInicio) <= #" & fechaFin & "#)" & _
                 " AND (((fechaFin) > #" & fechaInicio & "#) OR ((fechaFin) Is Null))" & _
                 ")"
    
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        If (rs.RecordCount > 0) Then
            isUserServiceActiveInInterval = True
        End If
        
        rs.Close
        Set rs = Nothing
    End If
    
    Exit Function
    
Error:
    Debug.Print "Error: " & Err.description
    isUserServiceActiveInInterval = False
End Function

'---------------------------------------------------------------------------
'   Name:   ServicioActivoUsuario
'   Autor:  Antonio Nadal
'   Fecha:  08/06/2010
'   Desc:   Si se le pasa un servicio, indica si el usuario está de alta
'           en el servicio indicado en la fecha indicada
'           Si no se le pasa un servicio, indica si el usuario está de alta
'           en algún servicio en la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio (OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Número de servicios activos del usuario en la fecha indicada
'           -1, ko
'---------------------------------------------------------------------------
Public Function ServicioActivoUsuario(fecha As Date, _
                                        Optional idServicio As Long = 0, _
                                        Optional idPersona As Long = 0, _
                                        Optional idOrganizacion As Long = 0, _
                                        Optional idServicioUsuario As Long = 0) As Integer
On Error GoTo Error
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT " & _
                 IIf((idPersona = 0), "", " fkPersona") & _
                 IIf((idOrganizacion = 0), "", " fkOrganizacion") & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 IIf((idServicioUsuario = 0), "", " AND (id<>" & idServicioUsuario & ")") & _
                 " AND ((fechaInicio) <= #" & fechaActividad & "#)" & _
                 " AND (((fechaFin) > #" & fechaActividad & "#) OR ((fechaFin) Is Null))" & _
                 ");"
    
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        If (rs.RecordCount > 0) Then
            ServicioActivoUsuario = 1
        Else
            ServicioActivoUsuario = 0
        End If
            
        rs.Close
        Set rs = Nothing
        

    Else
        ServicioActivoUsuario = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    ServicioActivoUsuario = -1
End Function

'----------------------------------------------------------------------------------------------
'   Name:   ServicioPosteriorUsuario
'   Autor:  Antonio Nadal
'   Fecha:  08/06/2010
'   Desc:   Si se le pasa un servicio, indica si el usuario estará de alta en el servicio
'           posteriormente a la fecha indicada
'           Si no se le pasa un servicio, indica si el usuario estará de alta en algún servicio
'           posteriormente a la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio (OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Número de servicios activos del usuario con fecha posterior a la fecha indicada
'           (servicios futuros)
'           -1, ko
'----------------------------------------------------------------------------------------------
Public Function ServicioPosteriorUsuario(fecha As Date, _
                                        Optional idServicio As Long = 0, _
                                        Optional idPersona As Long = 0, _
                                        Optional idOrganizacion As Long = 0) As Integer
On Error GoTo TratarError
    
    Dim fechaActividad As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    fechaActividad = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT " & _
                 IIf((idPersona = 0), "", " fkPersona") & _
                 IIf((idOrganizacion = 0), "", " fkOrganizacion") & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 " AND ((fechaInicio) > #" & fechaActividad & "#)" & _
                 " AND (((fechaFin) > #" & fechaActividad & "#) OR ((fechaFin) Is Null))" & _
                 ");"
    
        'Debug.Print strSQL
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        If (rs.RecordCount > 0) Then
            ServicioPosteriorUsuario = 1
        Else
            ServicioPosteriorUsuario = 0
        End If
            
        rs.Close
        Set rs = Nothing
        
    Else
        ServicioPosteriorUsuario = -1
    End If
    
    Exit Function
    
TratarError:
    debugando "Error: " & Err.description
    ServicioPosteriorUsuario = -1
End Function

'----------------------------------------------------------------------------------------------
'   Name:   numServiciosIntervalo
'   Autor:  Antonio Nadal
'   Fecha:  08/06/2010
'   Desc:   Si se le pasa un servicio, indica si el usuario estará de alta en el servicio
'           posteriormente a la fecha indicada
'           Si no se le pasa un servicio, indica si el usuario estará de alta en algún servicio
'           posteriormente a la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio (OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Número de servicios activos del usuario con fecha posterior a la fecha indicada
'           (servicios futuros)
'           -1, ko
'----------------------------------------------------------------------------------------------
Public Function numServiciosIntervalo(fechaini As Date, _
                                        fechaFin As Date, _
                                        Optional idServicio As Long = 0, _
                                        Optional idPersona As Long = 0, _
                                        Optional idOrganizacion As Long = 0) As Integer
On Error GoTo Error
    
    Dim fechaActividadIni, fechaActividadFin As Date
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    fechaActividadIni = Format(fechaini, "mm/dd/yyyy hh:nn:ss")
    fechaActividadFin = Format(fechaFin, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT " & _
                 IIf((idPersona = 0), "", " fkPersona") & _
                 IIf((idOrganizacion = 0), "", " fkOrganizacion") & _
                 " FROM r_serviciousuario" & _
                 " WHERE (" & _
                 IIf((idPersona = 0), "", " (fkPersona=" & idPersona & ")") & _
                 IIf((idOrganizacion = 0), "", " (fkOrganizacion=" & idOrganizacion & ")") & _
                 IIf((idServicio = 0), "", " AND (fkServicio=" & idServicio & ")") & _
                 " AND ((fechaInicio) >= #" & fechaActividadIni & "#)" & _
                 " AND (((fechaFin) <= #" & fechaActividadFin & "#) OR ((fechaFin) Is Null))" & _
                 ");"
                 
        'Debug.Print strsql
        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        numServiciosIntervalo = rs.RecordCount
            
        rs.Close
        Set rs = Nothing
        
    Else
        numServiciosIntervalo = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    numServiciosIntervalo = -1
End Function

