Attribute VB_Name = "SIFOC_ProyectoEmprendedor"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------
'   Name: checkServiceWhenComeToAnAppointment
'   Desc: New service Aut when person come to the first appointment
'---------------------------------------------------------------------
Public Static Function checkServiceWhenComeToAnAppointment(idPersona As Long, idServicio As Long)
    Dim endService As Date
    
    If Not isUserServiceActive(now, idServicio, idPersona) And (idServicio = 9) Then 'Servicio activo
        '9 - Servicio emprendedores - AUT
        '19 - Baja por caducidad de la demanda
        endService = DateSerial(Year(now), 12, 31)
        If altaServicioUsuario(9, now, U_idIfocUsuarioActivo, endService, 19, , "Alta automática cuando acude a cita emprendedor", idPersona) = 0 Then
            MsgBox "Se realizó alta en servicio emprendedores hasta 31/12/" & Year(now)
        Else
            MsgBox "Error al realizar el alta en servicio emprendedores.", vbOKOnly, "Alert: SIFOC_Emprendedor"
        End If
    ElseIf (idServicio <> 9) Then
        MsgBox "No se creará alta automática porque el servicio de la cita no es autoempleo.", vbOKOnly, "Alert: SIFOC_Emprendedor"
    End If
End Function

'---------------------------------------------------------------------------
'   Autor:  Asunción Huertas
'   Fecha:  07/2009
'   Desc:   Módulo con funciones para
'           * obtener información de los proyectos de emprendedores
'           * relacionar los proyectos de emprendedores con las citas
'---------------------------------------------------------------------------

'Devuelve el número de socios del proyecto
Public Function CalcularNumSocios(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT COUNT(*) AS Socios " & _
             "FROM r_ProyectoEmprendedorPersona " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & " and socio=-1;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumSocios = rs!Socios
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el total de la inversión del proyecto
Public Function CalcularTotalInversion(idProyecto As Integer) As Long
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "SELECT COUNT(*) AS Registros, SUM(Importe) AS Inversion " & _
             "FROM r_ProyectoEmprendedorInversion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & ";"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    If (rs!Registros = 0) Then
        CalcularTotalInversion = 0
    Else
        CalcularTotalInversion = rs!inversion
    End If
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el total de la financiación del proyecto
Public Function CalcularTotalFinanciacion(idProyecto As Integer) As Long
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Registros, SUM(Importe) AS Financiacion " & _
             "FROM r_ProyectoEmprendedorFinanciacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & ";"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    If (rs!Registros = 0) Then
        CalcularTotalFinanciacion = 0
    Else
        CalcularTotalFinanciacion = rs!Financiacion
    End If
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el número de ayudas del proyecto
Public Function CalcularNumAyudas(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Ayudas " & _
             "FROM r_citausuario, r_citaproyectoemprendedorAyuda " & _
             "WHERE r_citausuario.fkProyectoEmprendedor = " & idProyecto & _
             " and r_citaproyectoemprendedorAyuda.fkCita = r_citausuario.fkCita;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumAyudas = rs!Ayudas
    rs.Close
    Set rs = Nothing
End Function
   
'Devuelve el número de trámites del proyecto
Public Function CalcularNumTramites(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Tramites " & _
             "FROM r_citausuario, r_CitaTramite " & _
             "WHERE r_citausuario.fkProyectoEmprendedor = " & idProyecto & _
             " and r_CitaTramite.fkCita = r_citausuario.fkCita;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumTramites = rs!Tramites
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el número de actuaciones en curso del proyecto
Public Function CalcularNumActuacionesEnCurso(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS ActuacionesEnCurso " & _
             "FROM t_ProyectoEmprendedorPlanActuacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & " and fkProyectoEmprendedorEstado=1;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumActuacionesEnCurso = rs!ActuacionesEnCurso
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el número de actuaciones finalizadas del proyecto
Public Function CalcularNumActuacionesFinalizadas(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS ActuacionesFinalizadas " & _
             "FROM t_ProyectoEmprendedorPlanActuacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & " and fkProyectoEmprendedorEstado=2;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumActuacionesFinalizadas = rs!ActuacionesFinalizadas
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el número de actuaciones no realizadas del proyecto
Public Function CalcularNumActuacionesNoRealizadas(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS ActuacionesNoRealizadas " & _
             "FROM t_ProyectoEmprendedorPlanActuacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & " and fkProyectoEmprendedorEstado=3;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumActuacionesNoRealizadas = rs!ActuacionesNoRealizadas
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el número de actuaciones pasadas
Public Function CalcularNumActuacionesPasadas(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS ActuacionesPasadas " & _
             "FROM t_ProyectoEmprendedorPlanActuacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & _
             "  AND (fkProyectoEmprendedorEstado=1) " & _
             "  AND (NOT IsNull(fechaFin)) AND (fechaFin < Date());"
  
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumActuacionesPasadas = rs!ActuacionesPasadas
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el número de actuaciones presentes
Public Function CalcularNumActuacionesPresentes(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS ActuacionesPresentes " & _
             "FROM t_ProyectoEmprendedorPlanActuacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & _
             "  AND (fkProyectoEmprendedorEstado=1) " & _
             "  AND (fechaInicio <= Date()) AND ((IsNull(fechaFin) OR fechaFin >= Date()));"
   
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumActuacionesPresentes = rs!ActuacionesPresentes
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el número de actuaciones futuras
Public Function CalcularNumActuacionesFuturas(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS ActuacionesFuturas " & _
             "FROM t_ProyectoEmprendedorPlanActuacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & _
             "  AND (fkProyectoEmprendedorEstado=1) " & _
             "  AND (fechaInicio > Date());"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumActuacionesFuturas = rs!ActuacionesFuturas
    rs.Close
    Set rs = Nothing
End Function

'Devuelve la fecha de inicio de más antigua de todas las actuaciones
Public Function CalcularFechaInicioActuaciones(idProyecto As Integer) As Date
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Registros, MIN(fechaInicio) AS FechaInicio " & _
             "FROM t_ProyectoEmprendedorPlanActuacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & " and fkProyectoEmprendedorEstado<>3;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    If (rs!Registros = 0) Or (IsNull(rs!fechaInicio)) Then
        CalcularFechaInicioActuaciones = DateSerial(1800, 1, 1)
    Else
        CalcularFechaInicioActuaciones = rs!fechaInicio
    End If
    rs.Close
    Set rs = Nothing
End Function

'Devuelve la fecha de fin última de todas las actuaciones
Public Function CalcularFechaFinActuaciones(idProyecto As Integer) As Date
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Registros, MAX(fechaFin) AS FechaFin " & _
             "FROM t_ProyectoEmprendedorPlanActuacion " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & " and fkProyectoEmprendedorEstado<>3;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    If (rs!Registros = 0) Or (IsNull(rs!fechaFin)) Then
        CalcularFechaFinActuaciones = DateSerial(9999, 12, 31)
    Else
        CalcularFechaFinActuaciones = rs!fechaFin
    End If
    rs.Close
    Set rs = Nothing
End Function

'Devuelve el número de citas del proyecto
Public Function CalcularNumCitas(idProyecto As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Citas " & _
             "FROM r_citausuario " & _
             "WHERE fkProyectoEmprendedor = " & idProyecto & ";"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    CalcularNumCitas = rs!Citas
    rs.Close
    Set rs = Nothing
End Function

'Devuelve la fecha de la primera cita
Public Function CalcularFechaPrimeraCita(idProyecto As Integer) As Date
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Registros, MIN(t_Cita.fecha) AS Fecha " & _
             "FROM r_citausuario, t_Cita " & _
             "WHERE r_citausuario.fkProyectoEmprendedor = " & idProyecto & _
             " and t_Cita.id = r_citausuario.fkCita;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    If (rs!Registros = 0) Or (IsNull(rs!fecha)) Then
        CalcularFechaPrimeraCita = DateSerial(1800, 1, 1)
    Else
        CalcularFechaPrimeraCita = rs!fecha
    End If
    rs.Close
    Set rs = Nothing
End Function

'Devuelve la fecha de la última cita
Public Function CalcularFechaUltimaCita(idProyecto As Integer) As Date
    Dim rs As New ADODB.Recordset
    Dim strSql As String
      
    strSql = "SELECT COUNT(*) AS Registros, MAX(t_Cita.fecha) AS Fecha " & _
             "FROM r_citausuario, t_Cita " & _
             "WHERE r_citausuario.fkProyectoEmprendedor = " & idProyecto & _
             " and t_Cita.id = r_citausuario.fkCita;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    If (rs!Registros = 0) Or (IsNull(rs!fecha)) Then
        CalcularFechaUltimaCita = DateSerial(9999, 12, 31)
    Else
        CalcularFechaUltimaCita = rs!fecha
    End If
    rs.Close
    Set rs = Nothing
End Function

'Comprueba si la persona está de alta en el servicio de autoempleo
'Public Function noAltaServicio(persona As Long, fechaEntrada, fechaSalida As Date) As Boolean
'
'    Dim rs As New ADODB.Recordset
'    Dim strSql As String
'
'    strSql = " SELECT COUNT(*) AS Registros " & _
'             " FROM r_serviciousuario " & _
'             " WHERE fkPersona = " & persona & _
'             "  AND fkServicio = 9 " & _
'             "  AND fechaInicio <= #" & Format(fechaEntrada, "mm/dd/yyyy") & "#" & _
'             "  AND (IsNull(fechaFin) OR fechaFin >= #" & Format(fechaSalida, "mm/dd/yyyy") & " 23:59:59#);"
'
'    'debugando strSql
'
'    Set rs = New ADODB.Recordset
'    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
'
'    If rs!Registros = 0 Then noAltaServicio = True
'
'    rs.Close
'    Set rs = Nothing
'End Function

'Devuelve el número de socios del proyecto
Public Function lastIdProyectPerson(idPerson As Long) As Long
    Dim rs As New ADODB.Recordset
    Dim strSql As String
    Dim idProyecto As Long
      
    strSql = " SELECT last(id) as idP" & _
             " FROM t_proyectoemprendedor LEFT JOIN r_ProyectoEmprendedorPersona ON t_proyectoemprendedor.id = r_proyectoemprendedorpersona.fkProyectoemprendedor" & _
             " WHERE fkPersona = " & idPerson
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    If Not (rs.EOF) Then
        rs.MoveFirst
        idProyecto = Nz(rs!idP, 0)
    Else
        idProyecto = 0
    End If
    rs.Close
    Set rs = Nothing
    lastIdProyectPerson = idProyecto
End Function

