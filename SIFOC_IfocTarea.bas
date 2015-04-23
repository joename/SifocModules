Attribute VB_Name = "SIFOC_IfocTarea"
Option Explicit
Option Compare Database

Dim strSql As String
Dim strselect As String
Dim strFrom As String
Dim strWhere As String
Dim strGroup As String
Dim strHaving As String
Dim strOrder As String

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/12/2013 - Actualización:  21/12/2013
'   Name:   SIFOC_IfocTarea
'   Desc:   IfocTarea module
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------

'------------Getters------------------------
Public Static Function getSelect() As String
    getSelect = strselect
End Function

Public Static Function getFrom() As String
    getFrom = strFrom
End Function

Public Static Function getWhere() As String
    getWhere = strWhere
End Function

Public Static Function getGroup() As String
    getGroup = strGroup
End Function

Public Static Function getOrder() As String
    getOrder = strOrder
End Function

'------------Setters--------------------------
Public Static Function setSql(sSql As String) As String
    strSql = sSql
    setSql = strSql
End Function

Public Static Function setSelect(sSelect As String) As String
    strselect = sSelect
    setSelect = strselect
End Function

Public Static Function setFrom(sFrom As String) As String
    strFrom = sFrom
    setFrom = strFrom
End Function

Public Static Function setWhere(sWhere As String) As String
    strWhere = sWhere
    setWhere = strWhere
End Function

Public Static Function setGroup(sGroup As String) As String
    setGroup = sGroup
End Function

Public Static Function setOrder(sOrder As String) As String
    strOrder = sOrder
    setOrder = strOrder
End Function

'------------Adders--------------------------
Public Static Function addSql(sSql As String) As String
    strSql = addConditionWhere(strSql, sSql)
    addSql = strSql
End Function

Public Static Function addSelect(sSelect As String) As String
    strselect = addConditionWhere(strselect, sSelect)
    addSelect = strselect
End Function

Public Static Function addFrom(sFrom As String) As String
    strFrom = addConditionWhere(strFrom, sFrom)
    addFrom = strFrom
End Function

Public Static Function addWhere(sWhere As String) As String
    strWhere = addConditionWhere(strWhere, sWhere)
    addWhere = strWhere
End Function

Public Static Function addGroup(sGroup As String) As String
    addGroup = addConditionWhere(strGroup, sGroup)
End Function

Public Static Function addOrder(sOrder As String) As String
    strOrder = addConditionWhere(strOrder, sOrder)
    addOrder = strOrder
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/12/2013 - Actualización:  21/12/2013
'   Name:   dni
'   Desc:   initSql module
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSql() As String
    strselect = "t_ifoctarea.id, t_ifoctarea.timestamp, t_ifoctarea.tarea, t_ifoctarea.descripcion, t_ifoctarea.fecha, t_ifoctarea.fechaAproxRes, t_ifoctarea.fkIfocTareaCategoria, t_ifoctarea.fkTareaPrioridad, t_ifoctarea.fkTareaEstado, t_ifoctarea.fkIFOCUsuario, t_ifoctarea.fkIFOCUsuarioRes, t_ifoctarea.valoracion, Count(t_ifoctareaaccion.fkIfocTarea) AS acciones, Sum(t_ifoctareaaccion.duracion) AS duracion,datediff('d', t_ifoctarea.fecha, t_ifoctarea.timestamp) AS tresp, pdte"
    strFrom = "t_ifoctarea LEFT JOIN t_ifoctareaaccion ON t_ifoctarea.id = t_ifoctareaaccion.fkIfocTarea"
    'Estado abierto y pendiente
    strWhere = ""
    strHaving = ""
    strGroup = "t_ifoctarea.id, t_ifoctarea.tarea, t_ifoctarea.descripcion, t_ifoctarea.fecha, t_ifoctarea.fechaAproxRes, t_ifoctarea.fkIfocTareaCategoria, t_ifoctarea.fkTareaPrioridad, t_ifoctarea.fkTareaEstado, t_ifoctarea.fkIFOCUsuario, t_ifoctarea.fkIFOCUsuarioRes, t_ifoctarea.valoracion, t_ifoctarea.timestamp, pdte"
    strOrder = "t_ifoctarea.timestamp DESC"
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/12/2013 - Actualización:  21/12/2013
'   Name:   dni
'   Desc:   initSqlClean module
'   Param:  ?
'   Retur:  ?
'---------------------------------------------------------------------------
Public Static Function initSqlClean() As String
    strselect = ""
    strFrom = ""
    'Estado abierto y pendiente
    strWhere = ""
    strHaving = ""
    strGroup = ""
    strOrder = ""
End Function

Public Static Function updIfocTareaTimestamp(idIfocTarea As Long) As Integer
    Dim val As String
    Dim idTareaPrioridad As Long
    'val = Nz(Forms!GestionIfocTarea!txt_Valoracion, "") & " " 'ojito con esto, debe estar abierto con la tarea correcta formulario
On Error GoTo TratarError
    idTareaPrioridad = DLookup("[fkTareaPrioridad]", "t_ifoctarea", "[id]=" & idIfocTarea)

    strSql = " UPDATE t_ifoctarea t SET t.fkTareaPrioridad =" & idTareaPrioridad & _
             " WHERE id =" & idIfocTarea
    
    'strSql = " UPDATE t_ifoctarea t SET t.timestamp = now()" & _
             " WHERE id =" & idIfocTarea
    
    'strSql = " UPDATE t_ifoctarea t SET t.valoracion = '" & val & " '" & _
             " WHERE id =" & idIfocTarea
    
    CurrentDb.Execute strSql
    
SalirTratarError:
    Exit Function
TratarError:
    Debug.Print "Error (updIfocTareaTimestamp): " & vbNewLine & _
            Err.description & "Alert: SIFOC_IfocUsuario"
    updIfocTareaTimestamp = -1
End Function

Public Static Function infoTareaResActivas(idResponsable As Long) As String
    Dim str As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = " SELECT Count(t_ifoctarea.id) AS tareas" & _
          " FROM t_ifoctarea" & _
          " WHERE (t_ifoctarea.fkTareaEstado < 3) AND ((t_ifoctarea.fkIFOCUsuarioRes)=" & idResponsable & ")"
    
    Set rs = New ADODB.Recordset
    
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        str = ""
        str = str & "Tareas pendientes:     " & rs!TAREAS & vbNewLine
    End If
    
    rs.Close
    
    sql = " SELECT Count(t_ifoctareaaccion.fkIfocTarea) AS acciones, Sum(t_ifoctareaaccion.duracion) AS tiempo" & _
          " FROM t_ifoctarea LEFT JOIN t_ifoctareaaccion ON t_ifoctarea.id = t_ifoctareaaccion.fkIfocTarea" & _
          " WHERE (t_ifoctarea.fkTareaEstado < 3) AND ((t_ifoctarea.fkIFOCUsuarioRes)=" & idResponsable & ")"
    
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        str = str & "Acciones realizadas:   " & rs!acciones & vbNewLine
        str = str & "Tiempo(min.) empleado: " & rs!tiempo & vbNewLine
    End If
    
    rs.Close
    Set rs = Nothing
    
    infoTareaResActivas = str
    
End Function

Public Static Function infoTareaDemActivas(idDemandante As Long) As String
    Dim str As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = " SELECT Count(t_ifoctarea.id) AS tareas" & _
          " FROM t_ifoctarea" & _
          " WHERE (t_ifoctarea.fkTareaEstado < 3) AND ((t_ifoctarea.fkIFOCUsuario)=" & idDemandante & ")"
    
    Set rs = New ADODB.Recordset
    
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        str = ""
        str = str & "Tareas pendientes:     " & rs!TAREAS & vbNewLine
    End If
    
    rs.Close
    
    sql = " SELECT Count(t_ifoctareaaccion.fkIfocTarea) AS acciones, Sum(t_ifoctareaaccion.duracion) AS tiempo" & _
          " FROM t_ifoctarea LEFT JOIN t_ifoctareaaccion ON t_ifoctarea.id = t_ifoctareaaccion.fkIfocTarea" & _
          " WHERE (t_ifoctarea.fkTareaEstado < 3) AND ((t_ifoctarea.fkIFOCUsuario)=" & idDemandante & ")"
    
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        str = str & "Acciones realizadas:   " & rs!acciones & vbNewLine
        str = str & "Tiempo(min.) empleado: " & rs!tiempo & vbNewLine
    End If
    
    rs.Close
    Set rs = Nothing
    
    infoTareaDemActivas = str
    
End Function

Public Static Function getInfoTareasRealizadas(idResponsable As Long, fechaInicio As Date, fechaFin As Date) As Variant
    Dim strSql, sqlAcciones As String
    Dim rs As dao.Recordset
    Dim fechaI, FECHAF As Date
    Dim estTareasRea(3) As Integer
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"
    
    sqlAcciones = " SELECT fkIfocTarea as idIfocTarea, Count(t_ifoctareaaccion.fkIfocTarea) as acciones, Sum(t_ifoctareaaccion.duracion) as tiempo" & _
                  " FROM t_ifoctareaaccion" & _
                  " GROUP BY fkIfocTarea"

    strSql = " SELECT Count(t_ifoctarea.id) as tareas, Sum(acc.acciones) as acciones, Sum(acc.tiempo) as tiempo" & _
          " FROM t_ifoctarea LEFT JOIN (" & sqlAcciones & ") as acc ON t_ifoctarea.id = acc.idIfocTarea" & _
          " WHERE (t_ifoctarea.fecha Between #" & fechaI & "# AND #" & FECHAF & "#)" & _
            " AND (t_ifoctarea.fkIFOCUsuarioRes =" & idResponsable & ")"
    
    G_Connection.setDatabase
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        estTareasRea(1) = Nz(rs!TAREAS, 0)
        estTareasRea(2) = Nz(rs!acciones, 0)
        estTareasRea(3) = Nz(rs!tiempo, 0)
    End If
    
    rs.Close
    Set rs = Nothing
Debug.Print strSql & vbNewLine & estTareasRea(1) & "-" & estTareasRea(2) & "-" & estTareasRea(3)
    getInfoTareasRealizadas = estTareasRea
    
End Function

Public Static Function getInfoTareasDemandadas(idDemandante As Long, fechaInicio As Date, fechaFin As Date) As Variant
    Dim num As Integer
    Dim strSql, sqlAcciones As String
    Dim rs As dao.Recordset
    Dim fechaI, FECHAF As Date
    Dim estTareasDem(3) As Integer
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"
    
    sqlAcciones = " SELECT fkIfocTarea as idIfocTarea, Count(t_ifoctareaaccion.fkIfocTarea) as acciones, Sum(t_ifoctareaaccion.duracion) as tiempo" & _
                  " FROM t_ifoctareaaccion" & _
                  " GROUP BY fkIfocTarea"

    strSql = " SELECT Count(t_ifoctarea.id) as tareas, Sum(acc.acciones) as acciones, Sum(acc.tiempo) as tiempo" & _
             " FROM t_ifoctarea LEFT JOIN (" & sqlAcciones & ") as acc ON t_ifoctarea.id = acc.idIfocTarea" & _
             " WHERE (t_ifoctarea.fecha Between #" & fechaI & "# AND #" & FECHAF & "#)" & _
             " AND (t_ifoctarea.fkIFOCUsuario =" & idDemandante & ")"
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        estTareasDem(1) = Nz(rs!TAREAS, 0)
        estTareasDem(2) = Nz(rs!acciones, 0)
        estTareasDem(3) = Nz(rs!tiempo, 0)
    End If
    
    rs.Close
    Set rs = Nothing
    
    getInfoTareasDemandadas = estTareasDem
    
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

Public Function hayTareasPendientesPorDemandante(idDemandante) As Boolean
    Dim respuesta As Boolean
    Dim rs As dao.Recordset
    Dim strSql As String
    
    SIFOC_IfocTarea.initSqlClean
    SIFOC_IfocTarea.setSelect ("id")
    SIFOC_IfocTarea.setFrom ("t_ifoctarea")
    SIFOC_IfocTarea.setWhere ("pdte = 0")
    SIFOC_IfocTarea.addWhere ("fkIfocUsuario = " & idDemandante)
    SIFOC_IfocTarea.addWhere ("fkTareaEstado < 3")
    
    G_Connection.setDatabase
    
    strSql = SIFOC_IfocTarea.getQuery
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
    If Not rs.EOF Then
        If (rs.RecordCount > 0) Then
            respuesta = True
        Else
            respuesta = False
        End If
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    hayTareasPendientesPorDemandante = respuesta
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  21/07/2014 - Actualización:  21/07/2014
'   Name:   hayTareasPendientesPorDemandante
'   Desc:   indica si la persona demandante tiene tareas pendiente más
'           de un mes
'   Param:  idDemandante, id persona que solicita tarea
'   Retur:  boolean, true si tiene tareas pendientes más de un mes
'                    false si no tiene tareas pendientes más de un mes
'---------------------------------------------------------------------------
Public Function hayTareasPendientesPorDemandanteConRetraso(idDemandante As Long, dias As Integer) As Boolean
    Dim respuesta As Boolean
    Dim rs As dao.Recordset
    Dim strSql As String
    
    SIFOC_IfocTarea.initSqlClean
    SIFOC_IfocTarea.setSelect ("id, fkIfocUsuario, fkTareaEstado, timestamp")
    SIFOC_IfocTarea.setFrom ("t_ifoctarea")
    SIFOC_IfocTarea.setWhere ("pdte = 0")
    SIFOC_IfocTarea.addWhere ("fkIfocUsuario = " & idDemandante)
    SIFOC_IfocTarea.addWhere ("fkTareaEstado < 3")
    SIFOC_IfocTarea.addWhere ("timestamp < #" & Format(now - dias, "mm/dd/yyyy") & "#")
    
    strSql = SIFOC_IfocTarea.getQuery
    
    G_Connection.setDatabase
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
    If Not rs.EOF Then
        If (rs.RecordCount > 0) Then
Debug.Print "Dias(" & dias & ") idtarea " & "-id:" & rs!id & "-idifoc:" & rs!fkIFOCUsuario & "-est:" & rs!fkTareaEstado & "-t:" & rs!timestamp & ">" & strSql
            respuesta = True
        Else
            respuesta = False
        End If
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    hayTareasPendientesPorDemandanteConRetraso = respuesta
End Function

Public Static Function insertIfocTarea(title As String, desc As String, ifocUser As Long)
    Dim str As String
    
    str = " INSERT INTO t_ifoctarea (tarea, descripcion, fecha, fkIfocTareaCategoria, fkTareaPrioridad, fkTareaEstado, fkIfocUsuario, fkIfocUsuarioRes)" & _
          " VALUES ('" & title & "','" & desc & "',now(),1,1,1," & ifocUser & ",14)"

Debug.Print str
    
    u_db.Execute str
End Function

Public Static Function checkPendingTasks(formName As String)
    
    '---Init - Comprobacion tareas pendiente demandate (con más de un mes)
    If InStr(1, formName, "MenuGeneral", vbTextCompare) > 0 Or InStr(1, formName, "IfocTarea", vbTextCompare) > 0 Then
        Exit Function
    End If
    
    If hayTareasPendientesPorDemandanteConRetraso(U_idIfocUsuarioActivo, 30) Then
Debug.Print "hello 30"
        Dim msg As String
        msg = "Tienes tareas demandadas con acciones pendientes de hace más de un mes." & vbNewLine & _
              "La aplicación SIFOC se ha bloqueado." & vbNewLine & _
              "Revísalas para poder continuar!"
        G_Timer.wait (1)
Debug.Print "check tasks 30 " & U_idIfocUsuarioActivo & " - true "
        MsgBox msg, vbOKOnly, "Alert: SIFOC_Control + control tarea pendientes +30"
    ElseIf hayTareasPendientesPorDemandanteConRetraso(U_idIfocUsuarioActivo, 15) Then
Debug.Print "hello 15"
            Dim num As Integer
            Randomize
            num = Int(Rnd() * 5)
            If (num = 1) Then
                msg = "Tienes tareas demandadas con acciones pendientes de hace más de 15 días." & vbNewLine & _
                      "La aplicación SIFOC se bloqueará si mantienes tareas pendientes." & vbNewLine & _
                      "De forma aleatoria se irá informando de esta/s tareas!"
                G_Timer.wait (1)
Debug.Print "check tasks 15 " & formName & " - ifocuser: " & U_idIfocUsuarioActivo & " - true"
                MsgBox msg, vbOKOnly, "Alert: SIFOC_Control + control tarea pendientes +15"
            End If
    End If
    
End Function

Public Function help()
    updIfocTareaTimestamp (1)
End Function


