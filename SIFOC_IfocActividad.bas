Attribute VB_Name = "SIFOC_IfocActividad"
Option Explicit
Option Compare Database


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  07/04/2014 - Actualización:  07/04/2014
'   Name:   getSqlMyLastWeekActivity
'   Desc:   Obtenemos sql con actividad de última semana
'   Param:  idPerson as Long
'   Retur:  sql as String
'---------------------------------------------------------------------------
Public Static Function getSqlActivity(idPerson As Long) As String
    Dim sql As String
    
    Dim from As Date
    from = DateAdd("d", -5, now)
    
    'Campos para mi actividad
    'Tipo(Tarea, TareaD, ...), fecha, descripcion
    
    sql = getSqlMyTasks(idPerson) & _
          " UNION " & _
          getSqlMyDemands(idPerson) & _
          " UNION " & _
          getSqlGestiones(idPerson, from) & _
          " UNION " & _
          getSqlCitas(idPerson, from) & _
          " ORDER BY Fecha DESC"

    SIFOC_QuerySql.RewriteQuerySQL "myActivity", sql

    getSqlActivity = "myActivity"
End Function

Private Function getSqlMyTasks(idPerson As Long) As String
    Dim strselect As String
    
    SIFOC_IfocTarea.initSqlClean
    strselect = "id, 'Tarea' as Actividad, t_ifoctarea.timestamp as Fecha, t_ifoctarea.tarea as Descripcion"
    SIFOC_IfocTarea.setSelect strselect
    SIFOC_IfocTarea.setFrom "t_ifoctarea"
    SIFOC_IfocTarea.setWhere "fkTareaEstado < 3 AND t_ifoctarea.fkIFOCUsuariores=" & U_idIfocUsuarioActivo
    SIFOC_IfocTarea.setOrder ""
    
    getSqlMyTasks = SIFOC_IfocTarea.getQuery
End Function

Private Function getSqlMyDemands(idPerson As Long) As String
    Dim strselect As String
    
    SIFOC_IfocTarea.initSqlClean
    
    strselect = "id, 'TareaDem' as Actividad, t_ifoctarea.timestamp as Fecha, t_ifoctarea.tarea as Descripcion"
    SIFOC_IfocTarea.setSelect strselect
    SIFOC_IfocTarea.setFrom "t_ifoctarea"
    SIFOC_IfocTarea.setWhere "fkTareaEstado < 3 AND t_ifoctarea.fkIFOCUsuario=" & U_idIfocUsuarioActivo
    SIFOC_IfocTarea.setOrder ""
    
    getSqlMyDemands = SIFOC_IfocTarea.getQuery
End Function

Private Function getSqlGestiones(idIfocUsuario As Long, fecha As Date) As String
    Dim strselect As String
    
    Dim sql As String
    
    SIFOC_Gestion.initSqlClean
    SIFOC_Gestion.setStrSelect ("id, 'Gestion' as Actividad, fecha, gestion as Descripcion")
    SIFOC_Gestion.setStrFrom ("t_gestion")
    SIFOC_Gestion.addStrWhere ("fecha >='" & Format(fecha, "yyyy-mm-dd hh:nn") & "'")
    SIFOC_Gestion.addStrWhere ("fkIfocUsuario = " & idIfocUsuario)
    SIFOC_Gestion.setStrGroup ("")
    SIFOC_Gestion.setStrOrder ("")
    
    getSqlGestiones = SIFOC_Gestion.getQuery
End Function

Public Function getSqlCitas(idIfocUsuario As Long, fecha As Date) As String
    Dim strselect As String
    
    Dim sql As String
    
    SIFOC_Cita.initSqlClean
    SIFOC_Cita.setStrSelect ("id, 'Cita' as Actividad, fecha, observacion as Descripcion")
    SIFOC_Cita.setStrFrom ("t_cita")
    SIFOC_Cita.addStrWhere ("fecha >='" & Format(fecha, "yyyy-mm-dd hh:nn") & "'")
    SIFOC_Cita.addStrWhere ("fecha <'" & Format(now + 1, "yyyy-mm-dd hh:nn") & "'")
    SIFOC_Cita.addStrWhere ("fkIfocUsuariotec = " & idIfocUsuario & " OR fkIfocUsuarioCit = " & idIfocUsuario)
    SIFOC_Cita.setStrGroup ("")
    SIFOC_Cita.setStrOrder ("")
    
    getSqlCitas = SIFOC_Cita.getQuery
End Function
