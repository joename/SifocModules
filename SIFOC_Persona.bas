Attribute VB_Name = "SIFOC_Persona"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/03/2010 - Actualización:  15/03/2010
'   Name:   nombreApellidos
'   Desc:   Devuelve el apellido y nombre de la persona que se pasa por
'           parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
Public Function nombreApellidos(idPersona As Long) As String
    Dim str As String
    
    str = Nz(DLookup("name", "v_nombreapellidos", "[id]=" & idPersona), "Sin especificar")
    
    nombreApellidos = str
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/03/2010 - Actualización:  15/03/2010
'   Name:   apellidosNombre
'   Desc:   Devuelve el apellido y nombre de la persona que se pasa por
'           parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
Public Function apellidosNombre(idPersona As Long) As String
    Dim str As String
    
    str = Nz(DLookup("name", "v_apellidosnombre", "[id]=" & idPersona), "Sin especificar")
    apellidosNombre = str
End Function

Public Function altaServicioPersona(idPersona As Long, _
                                    idServicio As Long, _
                                    fechaI As Date, _
                                    idIfocUsuario As Long, _
                                    Optional FECHAF As Date = "01/01/1900 00:00:00", _
                                    Optional idMotivoBaja As Long = 0, _
                                    Optional mejora As Integer = 0, _
                                    Optional observacion As String = "") As Integer
    altaServicioPersona = altaServicioOrganizacion(idPersona, idServicio, fechaI, idIfocUsuario, FECHAF, idMotivoBaja, mejora, observacion)
End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/03/2010 - Actualización:  17/03/2010
'   Name:   prestación
'   Desc:   Devuelve los si está cobrando prestación
'   Param:  idPersona(long), identificador de persona
'   Retur:  String que indica si está cobrando prestación.
'           Sí, fecha o ¿fecha? si no hay
'           No, no cobra prestación
'---------------------------------------------------------------------------------
Public Function PRESTACION(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    'Sólo miramos los carnes profesionales relacionados con Tráfico (idCarneProfesionalArea=2)
    strSql = " SELECT fechainicio, fechafin, cantidad" & _
             " FROM t_prestaciones" & _
             " WHERE (t_prestaciones.fkPersona=" & idPersona & ")" & _
             " AND (((t_prestaciones.fechainicio)<Now() Or (t_prestaciones.fechainicio) Is Null) AND ((t_prestaciones.fechafin)>=Now() Or (t_prestaciones.fechafin) Is Null));"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
        
    If Not rs.EOF Then
        rs.MoveFirst
        str = "de"
        If IsDate(rs!fechaInicio) Then
            str = str & " " & Format(rs!fechaInicio, "dd/mm/yy")
        Else
            str = str & " ¿fecha?"
        End If
        str = str & " a"
        If IsDate(rs!fechaFin) Then
            str = str & " " & Format(rs!fechaFin, "dd/mm/yy")
        Else
            str = str & " ¿fecha?"
        End If
    Else
        str = "No cobra prestación"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    PRESTACION = str
    
End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/03/2010 - Actualización:  17/03/2010
'   Name:   permisoTrabajo
'   Desc:   Devuelve los si tiene permiso de trabajo
'   Param:  idPersona(long), identificador de persona
'   Retur:  String que indica si tiene permiso de trabajo.
'           Sí, hasta fecha o ¿fecha?
'           No, sin permiso de trabajo
'---------------------------------------------------------------------------------
Public Function permisoTrabajo(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    'Sólo miramos los carnes profesionales relacionados con Tráfico (idCarneProfesionalArea=2)
    strSql = " SELECT fechaCaducidad" & _
             " FROM t_permisotrabajo" & _
             " WHERE (t_permisotrabajo.fkPersona=" & idPersona & ")" & _
             " AND (((t_permisotrabajo.fechaCaducidad)>=Now()) Or (t_permisotrabajo.fechaCaducidad) Is Null);"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
        
    If Not rs.EOF Then
        rs.MoveFirst
        str = "hasta"
        If IsDate(rs!fechaCaducidad) Then
            str = str & " " & Format(rs!fechaCaducidad, "dd/mm/yy")
        Else
            str = str & " ¿fecha?"
        End If
    Else
        str = "Sin permiso."
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    permisoTrabajo = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/08/2010 - Actualización:  16/08/2010
'   Name:   numSs
'   Desc:   Devuelve el número de la seguridad social de la persona pasada por parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con número seguridad social
'---------------------------------------------------------------------------------
Public Function numss(idPersona As Long) As String
    Dim str As String
    
    str = Nz(DLookup("seguridadsocial", "t_datospersona", "[fkPersona]=" & idPersona), 0)
    
    If (str = "0") Then
        str = "Sin especificar"
    End If
    
    numss = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/08/2010 - Actualización:  16/08/2010
'   Name:   soib
'   Desc:   Devuelve última fecha soib y tipo(sellado o inscripción) de la persona pasada por parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con última fecha soib y tipo
'---------------------------------------------------------------------------------
Public Function SOIB(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    strSql = " SELECT fecha, sellado" & _
             " FROM t_datossoib" & _
             " WHERE fkPersona = " & idPersona & _
             " ORDER BY fecha DESC;"
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        str = IIf(rs!sellado = 0, "Inscripción", "Sellado") & " " & Format(rs!fecha, "dd/mm/yyyy")
    Else
        str = "Sin especificar"
    End If
    
    'Close recordset
    rs.Close
    Set rs = Nothing
    
    SOIB = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/08/2010 - Actualización:  16/08/2010
'   Name:   nacionalidad
'   Desc:   Devuelve 1ª nacionalidad de la persona pasada por parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con nacionalidad
'---------------------------------------------------------------------------------
Public Function nacionalidad(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    strSql = " SELECT pais" & _
             " FROM t_datospersona LEFT JOIN a_pais ON t_datospersona.fkPaisNacionalidad = a_pais.id" & _
             " WHERE fkPersona = " & idPersona & ";"
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        str = IIf(IsNull(rs!pais), "Sin especificar", rs!pais)
    Else
        str = "Sin especificar"
    End If
    
    'Close recordset
    rs.Close
    Set rs = Nothing
    
    nacionalidad = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/08/2010 - Actualización:  16/08/2010
'   Name:   paisNacimiento
'   Desc:   Devuelve pais de nacimiento de la persona pasada por parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con país de nacimiento
'---------------------------------------------------------------------------------
Public Function paisNacimiento(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    strSql = " SELECT pais" & _
             " FROM t_datospersona LEFT JOIN a_pais ON t_datospersona.fkPaisNacimiento = a_pais.id" & _
             " WHERE fkPersona = " & idPersona & ";"
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        str = IIf(IsNull(rs!pais), "Sin especificar", rs!pais)
    Else
        str = "Sin especificar"
    End If
    
    'Close recordset
    rs.Close
    Set rs = Nothing
    
    paisNacimiento = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/08/2010 - Actualización:  16/08/2010
'   Name:   enCalviaDesde
'   Desc:   Devuelve pais de nacimiento de la persona pasada por parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con país de nacimiento
'---------------------------------------------------------------------------------
Public Function enCalviaDesde(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    strSql = " SELECT fechaResidenciaMunicipio" & _
             " FROM t_datospersona" & _
             " WHERE fkPersona = " & idPersona & ";"
    
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        str = IIf(IsNull(rs!fechaResidenciaMunicipio), "Sin especificar", rs!fechaResidenciaMunicipio)
    Else
        str = "Sin especificar"
    End If
    
    'Close recordset
    rs.Close
    Set rs = Nothing
    
    enCalviaDesde = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/03/2010 - Actualización:  15/03/2010
'   Name:   direccion
'   Desc:   Devuelve la dirección de la persona pasada por parámetro
'           Tenemos en cuenta la 1ªresidencia y en caso de no tener la 2ª res.
'           (esta consulta debe coincidir con la de direccion)
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con 1ª residencia ()
'---------------------------------------------------------------------------------
Public Function direccion(idPersona As Long) As String
    Dim str As String
    Dim numDir As Integer
    Dim strN As String
    Dim strSql As String
    
    strSql = " SELECT a_tipovia.tipoVia, t_direccion.direccion, t_direccion.bis, t_direccion.numero, t_direccion.bloque, t_direccion.escalera, t_direccion.piso, t_direccion.puerta" & _
             " FROM t_direccion LEFT JOIN a_tipovia ON t_direccion.fkTipoVia = a_tipovia.id" & _
             " WHERE t_direccion.fkPersona = " & idPersona & _
             " ORDER BY t_direccion.fkTipoDireccion ASC"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    numDir = rs.RecordCount
    If (numDir = 0) Then
        strN = ""
    Else
        strN = "(" & rs.RecordCount & ")"
    End If
    
    'Close rs
    rs.Close
    
    strSql = " SELECT a_tipovia.tipoVia, t_direccion.direccion, t_direccion.bis, t_direccion.numero, t_direccion.bloque, t_direccion.escalera, t_direccion.piso, t_direccion.puerta" & _
             " FROM t_direccion LEFT JOIN a_tipovia ON t_direccion.fkTipoVia = a_tipovia.id" & _
             " WHERE t_direccion.fkPersona = " & idPersona & " AND fkTipoDireccion = 1" & _
             " ORDER BY t_direccion.fkTipoDireccion ASC"
        
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    str = ""
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dirección
        str = IIf(Not IsNull(rs!TipoVia), rs!TipoVia, "")
        str = str & IIf(Not IsNull(rs!direccion), " " & rs!direccion, "")
        str = str & IIf(rs!numero <> 0, ", num. " & rs!numero, "")
        str = str & IIf(Not IsNull(rs!bis), ", bis: " & rs!bis, "")
        str = str & IIf(Not IsNull(rs!bloque), ", bloque: " & rs!bloque, "")
        str = str & IIf(Not IsNull(rs!escalera), ", esc: " & rs!escalera, "")
        str = str & IIf(Not IsNull(rs!piso), ", piso: " & rs!piso, "")
        str = str & IIf(Not IsNull(rs!puerta), ", puerta: " & rs!puerta, "")
    Else
        str = "Sin especificar"
    End If
    
    str = strN & " " & str
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    direccion = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/03/2010 - Actualización:  15/03/2010
'   Name:   direccion
'   Desc:   Devuelve la dirección de la persona pasada por parámetro
'           Tenemos en cuenta la 1ªresidencia y en caso de no tener la 2ª res.
'           (esta consulta debe coincidir con la de direccion)
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con cp, localidad, provincia
'---------------------------------------------------------------------------------
Public Function localidad(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    
    strSql = " SELECT t_direccion.codigoPostal, a_poblacion.poblacion, a_provincia.provincia, t_direccion.fkTipoDireccion" & _
             " FROM ((t_direccion LEFT JOIN a_tipovia ON t_direccion.fkTipoVia = a_tipovia.id) LEFT JOIN a_poblacion ON t_direccion.fkPoblacion = a_poblacion.id) LEFT JOIN a_provincia ON t_direccion.fkProvincia = a_provincia.id" & _
             " WHERE t_direccion.fkPersona = " & idPersona & "" & _
             " ORDER BY t_direccion.fkTipoDireccion ASC"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string localidad
        str = IIf(Not IsNull(rs!codigoPostal), rs!codigoPostal, "")
        str = str & IIf(Not IsNull(rs!poblacion), " " & rs!poblacion, "")
        str = str & IIf(Not IsNull(rs!provincia), ", " & rs!provincia, "")
    Else
        str = "Sin especificar"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    localidad = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/03/2010 - Actualización:  15/03/2010
'   Name:   telefonos
'   Desc:   Devuelve los teléfonos
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con teléfonos separados por comas.
'---------------------------------------------------------------------------------
Public Function telefonos(idPersona As Long) As String
    Dim str As String
    Dim strN As String
    Dim strSql As String
    Dim i As Integer
    Dim numTels As Integer
    Dim rs As ADODB.Recordset
    
    'Sólo miramos telefonos personales (fijo y móvil)
    strSql = " SELECT t_telefono.fkPersona, t_telefono.telefono" & _
             " FROM t_telefono" & _
             " WHERE (t_telefono.fkPersona =" & idPersona & ");"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    numTels = rs.RecordCount
    
    'Cerramos rs
    rs.Close
    
    'Sólo miramos telefonos personales (fijo y móvil)
    strSql = " SELECT t_telefono.fkPersona, t_telefono.telefono" & _
             " FROM t_telefono" & _
             " WHERE (t_telefono.fkTelefonoTipo = 1) And (t_telefono.fkTipoTelefono1 = 1) And (t_telefono.fkTipoTelefono2 <> 2)  And (t_telefono.fkPersona =" & idPersona & ")" & _
             " ORDER BY  t_telefono.fkTipoTelefono2 DESC;"
    
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    If (numTels = 0) Then
        strN = ""
    Else
        strN = "(" & numTels & ") "
    End If
    
    If Not rs.EOF Then
        rs.MoveFirst
        For i = 1 To 2 Step 1
            'Montamos string localidad
            If Not rs.EOF Then
                str = str & IIf(Len(str) < 8, rs!telefono, ", " & rs!telefono)
                rs.MoveNext
            End If
        Next i
    Else
        str = "Sin especificar"
    End If
    
    str = strN & str
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    telefonos = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  27/09/2011 - Actualización:  27/09/2011
'   Name:   strTelefonosMovilPersonal
'   Desc:   Devuelve los teléfonos moviles personales de las personas si los tienen
'   Param:  strIdPersonas, string idpersonas separados por comas
'   Retur:  String con teléfonos separados por comas.
'---------------------------------------------------------------------------------
Public Function strTelefonosMovilPersonal(ByRef strIdPersonas As String) As String
    Dim strTelefonos As String
    Dim idPersonas As Variant
    Dim idPersona
    Dim telefono As String
    Dim limit As Integer
    Dim newStrIdPersonas As String
    
    idPersonas = Split(strIdPersonas, ",")
    limit = UBound(idPersonas)
    
    newStrIdPersonas = ""
    For Each idPersona In idPersonas
        telefono = Nz(DLookup("telefono", "t_telefono", "fkPersona=" & CStr(idPersona) & " AND fkTipoTelefono1=1 AND fkTipoTelefono2=3"), "")
        
        If (Len(telefono) = 9) Then
            newStrIdPersonas = concatString(newStrIdPersonas, CStr(idPersona), ",")
            strTelefonos = concatString(strTelefonos, telefono, ",")
        End If
    Next
    
    strIdPersonas = newStrIdPersonas
    strTelefonosMovilPersonal = strTelefonos

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/03/2010 - Actualización:  17/03/2010
'   Name:   emails
'   Desc:   Devuelve los emails de la persona pasada por parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con emails separados por comas.
'---------------------------------------------------------------------------------
Public Function emails(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT email" & _
             " FROM t_email" & _
             " WHERE not isnull(t_email.fkPersona) AND (t_email.fkPersona =" & idPersona & ");"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = "(" & rs.RecordCount & ") "
    
    If Not rs.EOF Then
        rs.MoveFirst
        For i = 1 To 2 Step 1
            'Montamos string localidad
            If Not rs.EOF Then
                str = str & IIf(Len(str) < 8, rs!email, ", " & rs!email)
                rs.MoveNext
            End If
        Next i
    Else
        str = "Sin especificar"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    emails = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/03/2010 - Actualización:  17/03/2010
'   Name:   carnesConducir
'   Desc:   Devuelve los carnes profesionales relacionados con área Tráfico
'           idCarneProfesionalArea(2)
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con carnes de conducir separados por comas.
'---------------------------------------------------------------------------------
Public Function carnesConducir(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    'Sólo miramos los carnes profesionales relacionados con Tráfico (idCarneProfesionalArea=2)
    strSql = " SELECT a_carneprofesional.carneProfesional" & _
             " FROM t_carneprofesional INNER JOIN a_carneprofesional ON t_carneprofesional.fkCarneProfesional = a_carneprofesional.id" & _
             " WHERE (a_carneprofesional.fkCarneProfesionalArea=2) AND (t_carneprofesional.fkPersona=" & idPersona & ")" & _
             " ORDER BY a_carneprofesional.carneprofesional ASC;"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    Dim numRec As Integer
    numRec = rs.RecordCount
    
    If Not rs.EOF Then
        rs.MoveFirst
        For i = 1 To numRec Step 1
            'Montamos string localidad
            If Not rs.EOF Then
                str = str & IIf(str = "", rs!carneProfesional, ", " & rs!carneProfesional)
                rs.MoveNext
            End If
        Next i
    Else
        str = "Sin especificar"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    carnesConducir = str & IIf(disponeVehiculo(idPersona) = "", "", " con " & disponeVehiculo(idPersona))

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/03/2010 - Actualización:  17/03/2010
'   Name:   disponeVehiculo
'   Desc:   Devuelve los vehiculos que dispone la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con vehículos que dispone la persona separados por comas.
'---------------------------------------------------------------------------------
Private Function disponeVehiculo(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    'Sólo miramos los carnes profesionales relacionados con Tráfico (idCarneProfesionalArea=2)
    strSql = " SELECT disponeMoto, disponeCoche, disponeFurgoneta, disponeCamion" & _
             " FROM t_datospersona" & _
             " WHERE (t_datospersona.fkPersona=" & idPersona & ");"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone vehiculo/s
        If rs!disponeMoto = -1 Then
            str = str & IIf(Len(str) = 0, "moto", ", moto")
        End If
        If rs!disponeCoche = -1 Then
            str = str & IIf(Len(str) = 0, "coche", ", coche")
        End If
        If rs!disponeFurgoneta = -1 Then
            str = str & IIf(Len(str) = 0, "furgoneta", ", furgoneta")
        End If
        If rs!disponeCamion = -1 Then
            str = str & IIf(Len(str) = 0, "camión ", ", camión")
        End If
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    disponeVehiculo = str

End Function

'--------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  18/08/2010 - Actualización:  18/08/2010
'   Name:   estadoSituacionLaboral
'   Desc:   Te devuelve la situación laboral(estado)
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con estado de la situación laboral(actual)'
'           Devuelve el estado de la situación laboral de una persona de la que se recibe el id
'           Estados (a_situacionLaboralEstados)
'           1- Empleado: Con algun fechaFin de t_insercion a null o anterior a hoy
'           2- Desempleado: Sin ningun fechaFin de t_insercion a null
'           e.g. Si trabaja en 2 sitios a 1/2 joranada tendrá 2 trabajos activos a la vez
'           con fechafin = null y no es erroneo¿?
'--------------------------------------------------------------------------------------------
Public Function estadoSituacionLaboral(idPersona As Long) As String
    Dim rsTrabaja As ADODB.Recordset
    Dim strSql As String
    Dim idEstado As String
    
    strSql = " SELECT id" & _
             " FROM t_insercion " & _
             " WHERE fkPersona=" & idPersona & " AND fechaInicio < now() AND (fechaFin > now() OR isnull(FechaFin));"
    
    Set rsTrabaja = New ADODB.Recordset
    rsTrabaja.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    If Not rsTrabaja.EOF Then 'no tiene inserciones con FechaFin a Null
        'empleado
        idEstado = 1
    Else ' tiene inserciones con FechaFin a Null
        'desempleado
        idEstado = 2
    End If
    
    rsTrabaja.Close
    Set rsTrabaja = Nothing

    estadoSituacionLaboral = DLookup("[situacion]", "a_situacionlaboral", "[id]=" & idEstado)

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/08/2010 - Actualización:  17/08/2010
'   Name:   ultimaCita
'   Desc:   Te devuelve str de la última cita de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la última cita
'---------------------------------------------------------------------------------
Public Function ultimaCita(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT t_cita.fecha, a_gestiontipo.tipo, IIf([acudenoacudeanula]=-1,'Acude',IIf([acudenoacudeanula]=0,'No acude','Anula')) AS Asist" & _
             " FROM (t_cita INNER JOIN r_citausuario ON t_cita.id = r_citausuario.fkCita) LEFT JOIN a_gestiontipo ON t_cita.fkIfocAmbito = a_gestiontipo.id" & _
             " WHERE (((r_citausuario.fkPersona) =  " & idPersona & "))" & _
             " ORDER BY t_cita.fecha DESC;"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone vehiculo/s
        str = rs!asist & " " & rs!fecha & " (" & rs!tipo & ")"
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    ultimaCita = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/08/2010 - Actualización:  17/08/2010
'   Name:   ultimaCita
'   Desc:   Te devuelve str de la última cita de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la última cita
'---------------------------------------------------------------------------------
Public Function tareasPendientes(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT count(t_tarea.id) as numTareas, min(t_tarea.fechaLimite) as minFecha" & _
             " FROM t_tarea" & _
             " WHERE (t_tarea.fkPersona = " & idPersona & ") AND (t_tarea.realizado=0)" & _
             " GROUP BY t_tarea.id;"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone vehiculo/s
        str = rs!numTareas & " límite: " & Format(rs!minfecha, "dd/mm/yyyy") & ""
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    tareasPendientes = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/06/2011 - Actualización:  16/06/2011
'   Name:   getInfoUltimaCita
'   Desc:   Te devuelve str de la última cita de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la última cita
'---------------------------------------------------------------------------------
Public Function getInfoUltimaCita(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT t_cita.fecha, a_ifocambito.ambito, a_citasesion.sesion, IIf([acudenoacudeanula]=-1,'Acude',IIf([acudenoacudeanula]=0,'No acude','Anula')) AS Asist, t_cita.observacion as cita, r_citausuario.observacion as citapersona" & _
             " FROM ((t_cita INNER JOIN r_citausuario ON t_cita.id = r_citausuario.fkCita)" & _
             " LEFT JOIN a_ifocambito ON t_cita.fkIfocAmbito = a_ifocambito.id)" & _
             " LEFT JOIN a_citasesion ON t_cita.fkCitaSesion = a_citasesion.id" & _
             " WHERE (((r_citausuario.fkPersona) =  " & idPersona & "))" & _
             " ORDER BY t_cita.fecha DESC;"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone vehiculo/s
        str = rs!asist & " " & rs!fecha & " " & UCase(rs!sesion) & " (" & rs!Ambito & ")" & vbNewLine & _
              " CITA: " & rs!cita & vbNewLine & _
              " PERSONA: " & rs!citapersona
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    getInfoUltimaCita = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/08/2010 - Actualización:  17/08/2010
'   Name:   ultimaGestion
'   Desc:   Te devuelve str de la última gestion de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la última gestion
'---------------------------------------------------------------------------------
Public Function ultimaGestion(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT t_gestion.fecha, a_gestiontipo.tipo" & _
             " FROM (t_gestion INNER JOIN r_gestionusuario ON t_gestion.id = r_gestionusuario.fkGestion) LEFT JOIN a_gestiontipo ON t_gestion.fkIfocAmbito = a_gestiontipo.id" & _
             " WHERE (r_gestionusuario.fkPersona = " & idPersona & ")" & _
             " ORDER BY t_gestion.fecha DESC;"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string gestion
        str = rs!fecha & " (" & rs!tipo & ")"
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    ultimaGestion = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/06/2011 - Actualización:  16/06/2011
'   Name:   getInfoUltimaGestion
'   Desc:   Te devuelve str de la última gestion de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la última gestion
'---------------------------------------------------------------------------------
Public Function getInfoUltimaGestion(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT t_gestion.fecha, a_ifocambito.ambito, t_gestion.gestion, r_gestionusuario.observacion as gestionpersona" & _
             " FROM (t_gestion INNER JOIN r_gestionusuario ON t_gestion.id = r_gestionusuario.fkGestion) LEFT JOIN a_ifocambito ON t_gestion.fkIfocAmbito = a_ifocambito.id" & _
             " WHERE (r_gestionusuario.fkPersona = " & idPersona & ")" & _
             " ORDER BY t_gestion.fecha DESC;"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone gestion/s
        str = rs!fecha & " (" & rs!Ambito & ")" & vbNewLine & _
              "GESTION: " & rs!gestion & vbNewLine & _
              "PERSONA: " & rs!gestionpersona
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    getInfoUltimaGestion = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/06/2011 - Actualización:  16/06/2011
'   Name:   getInfoUltimaInsercion
'   Desc:   Te devuelve str de la última inserción de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la última gestion
'---------------------------------------------------------------------------------
Public Function getInfoUltimaInsercion(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    'Sólo miramos los carnes profesionales relacionados con Tráfico (idCarneProfesionalArea=2)
    strSql = " SELECT fechaInicio, fechaFin, cargo, ocupacion" & _
             " FROM t_insercion" & _
             " LEFT JOIN a_Cno2011 ON t_insercion.fkCno2011 = a_Cno2011.id" & _
             " WHERE (t_insercion.fkPersona = " & idPersona & ")" & _
             " ORDER BY fechaInicio DESC, fechaFin DESC;"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone gestion/s
        str = rs!fechaInicio & " - " & Nz(rs!fechaInicio, "¿?") & " " & IIf(IsNull(rs!cargo), rs!ocupacion, rs!cargo)
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    getInfoUltimaInsercion = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  20/08/2010 - Actualización:  20/08/2010
'   Name:   experiencia
'   Desc:   Te devuelve meses de experiencia de la persona
'           en la ocupacion y nivel, pasadas por parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  string (experiencia en meses)
'---------------------------------------------------------------------------------
Public Function experienciaCno(idPersona As Long, _
                               idCno As Integer, _
                               nivel As Integer) As String
    Dim str As String
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    strSql = " SELECT T_Insercion.fkPersona, T_Insercion.fkCno2011, T_Insercion.nivel, Sum(DateDiff('m',[fechaInicio],iif(isnull([fechaFin]),now(),[fechaFin]))) AS Experiencia" & _
             " FROM T_Insercion" & _
             " WHERE fkPersona = " & idPersona & " AND fkCno2011 = " & idCno & " AND nivel = " & nivel & _
             " GROUP BY T_Insercion.fkPersona, T_Insercion.fkCno2011, T_Insercion.nivel;"
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        str = rs!experiencia & " meses"
    Else
        str = "Sin experiencia"
    End If
    
    'Close recordset
    rs.Close
    Set rs = Nothing
    
    experienciaCno = str
End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  24/08/2010 - Actualización:  24/08/2010
'   Name:   ultimoCurso
'   Desc:   Te devuelve datos del último curso de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos del último curso
'---------------------------------------------------------------------------------
Public Function ultimoCurso(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT t_curso.nombre, a_tipocurso.aka as tipoCurso, t_curso.fechainicio" & _
             " FROM (t_curso INNER JOIN r_cursoalumno ON t_curso.id = r_cursoalumno.fkCurso) LEFT JOIN a_tipocurso ON t_curso.fkTipoCurso = a_tipocurso.id" & _
             " WHERE ((r_cursoalumno.fkPersona) =  " & idPersona & ")" & _
             " ORDER BY t_curso.fechaInicio DESC;"
   
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone vehiculo/s
        str = Format(rs!fechaInicio, "yyyy") & "-" & rs!tipoCurso & " " & rs!nombre
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    ultimoCurso = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  24/08/2010 - Actualización:  24/08/2010
'   Name:   ultimaOferta
'   Desc:   Te devuelve datos de la última oferta de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la última oferta
'---------------------------------------------------------------------------------
Public Function ultimaOferta(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT T_Oferta.id AS Oferta, T_Organizacion.nombre AS Empresa, T_Oferta.fechaOferta AS Fecha, T_Oferta.puesto AS Puesto, a_ofertasegestado.estado AS PreseleccionIFOC, a_ofertasegestadodetalle.estado AS [Estado Detalle]" & _
             " FROM (((R_OfertaCandidatos INNER JOIN T_Oferta ON R_OfertaCandidatos.fkOferta = T_Oferta.id) INNER JOIN T_Organizacion ON T_Oferta.fkOrganizacion = T_Organizacion.id) LEFT JOIN a_ofertasegestado ON R_OfertaCandidatos.fkOfertaSegEstado = a_ofertasegestado.id) LEFT JOIN a_ofertasegestadodetalle ON R_OfertaCandidatos.fkofertaSegEstadoDetalle = a_ofertasegestadodetalle.id" & _
             " WHERE ((R_OfertaCandidatos.fkPersona)=" & idPersona & ")" & _
             " ORDER BY T_Oferta.fechaOferta DESC , R_OfertaCandidatos.fkPersona;"
   
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone vehiculo/s
        str = Format(rs!fecha, "yyyy/mm") & "-" & rs!puesto & " (" & rs!PreseleccionIFOC & ")"
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    ultimaOferta = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  26/08/2010 - Actualización:  26/08/2010
'   Name:   ultimoProyEmprendedor
'   Desc:   Te devuelve datos de la último proyecto emprendedor de la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la último proyecto emprendedor
'---------------------------------------------------------------------------------
Public Function ultimoProyEmprendedor(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT t_proyectoemprendedor.nombreProy, t_proyectoemprendedor.fechaInicioProy AS fechaInicio, r_proyectoemprendedorpersona.socio" & _
             " FROM t_proyectoemprendedor INNER JOIN r_proyectoemprendedorpersona ON t_proyectoemprendedor.id = r_proyectoemprendedorpersona.fkProyectoEmprendedor" & _
             " WHERE (((r_proyectoemprendedorpersona.fkPersona)=" & idPersona & "))" & _
             " ORDER BY fechaInicioProy DESC"
   
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone vehiculo/s
        str = Format(rs!fechaInicio, "yyyy/mm") & "-" & rs!nombreProy & IIf(rs!socio = -1, "(socio)", "")
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    ultimoProyEmprendedor = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  26/08/2010 - Actualización:  26/08/2010
'   Name:   ctoEmpresa
'   Desc:   Te devuelve datos de cto de la empresa
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con datos de la última oferta
'---------------------------------------------------------------------------------
Public Function ctoEmpresa(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT t_organizacion.nombre, r_organizacionpersona.cargo" & _
             " FROM r_organizacionpersona INNER JOIN t_organizacion ON r_organizacionpersona.fkOrganizacion = t_organizacion.id" & _
             " WHERE (((r_organizacionpersona.fkPersona)=" & idPersona & "));"
   
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dispone vehiculo/s
        str = rs!nombre & " (" & rs!cargo & ")"
    Else
        str = "Sin información"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    ctoEmpresa = str

End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  11/04/2011 - Actualización:  11/04/2011
'   Name:   personaBajaServicios
'   Desc:   Da de baja de todos los servicios activos de la persona
'           con el motivo de baja, tecnico y observaciones pasadas por param.
'   Param:  idPersona(long), identificador de persona
'           fechaBaja (date), fecha de baja de servicios
'           idMotivo(int), motivo baja usuario
'           idIfocUsuario(long), id de ifoc usuario que da la baja
'           obs(texto), observaciones
'
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
Public Function personaBajaServicios(idPersona As Long, _
                                     fechaBaja As Date, _
                                     idMotivo As Integer, _
                                     idIfocUsuario As Integer, _
                                     Optional OBS As String) As Integer
On Error GoTo TratarError
    Dim fechaB As Date
    Dim strSql As String
    fechaB = Format(fechaBaja, "mm/dd/yyyy hh:mm:ss")
    
    strSql = " UPDATE r_serviciousuario" & _
             " SET r_serviciousuario.fechaFin =#" & fechaB & "#, r_serviciousuario.fkMotivoBaja = " & idMotivo & ", r_serviciousuario.fkIfocUsuarioBaja = " & idIfocUsuario & _
             " WHERE (((r_serviciousuario.fkPersona)=" & idPersona & ") AND ((r_serviciousuario.fechaFin) Is Null));"
    
    CurrentDb.Execute strSql
    
    personaBajaServicios = 0
    
SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error al dar de baja servicios" & Err.description, , "Modulo SIFOC_Persona"
    personaBajaServicios = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/06/2011 - Actualización:  15/06/2011
'   Name:   openFrmGestionPersona_instantanea
'   Desc:   Abre formulario GestionPersona_instantanea
'   Param:  idPersona(long), identificador de persona
'
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
'Public Function openFrmGestionPersona_instantanea(idPersona As Long)
'
'    'Dim db As database
'    Dim rst As DAO.Recordset
'    Dim strSql As String
'
'    Dim frm As New Form_GestionPersona_instantanea
'
'    If (idPersona <> 0) Then
'        frm.setIdPersona (idPersona)
'        frm.actualizaForm
'    End If
'
'    'frm.NavigationButtons = False
'    frm.visible = True
'
'    LigaFormulario.LigaFormulario frm
'
'End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/06/2011 - Actualización:  15/06/2011
'   Name:   updIdioma
'   Desc:   Actualizamos Idioma de persona
'   Param:  idPersona(long), identificador de persona
'   Retur:   0, OK
'           -1, KO
'---------------------------------------------------------------------------
Public Function updIdioma(idIdioma As Integer, _
                           idPersona As Long, _
                           idIdiomaNivelSimple As Long, _
                           idIfocUsuario As Long, _
                           OBS As String, _
                           Optional idCertificado As Integer = 0, _
                           Optional nivIfoc As Integer = 0, _
                           Optional lengMaterna As Integer = 0) As Integer
On Error GoTo TratarError
    Dim strSql As String
    
    strSql = " UPDATE t_idioma" & _
             " SET fkIdiomaNivelSimple = " & idIdiomaNivelSimple & _
             ", observacion = """ & OBS & """" & _
             IIf(nivIfoc = 0, "", ", nivelacionIfoc = " & nivIfoc) & _
             IIf(lengMaterna = 0, "", ", lenguaMaterna = " & lengMaterna) & _
             IIf(idCertificado = 0, "", ", fkCertificado = " & idCertificado) & _
             ", fkIfocUsuario = " & idIfocUsuario & _
             ", updDate = now()" & _
             " WHERE fkPersona = " & idPersona & " AND fkIdioma = " & idIdioma

'Debug.Print strSql

    CurrentDb.Execute strSql
    
SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error actualización de idioma." & Err.description, , "Modulo SIFOC_Persona"
    updIdioma = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/06/2011 - Actualización:  15/06/2011
'   Name:   openFrmGestionPersona
'   Desc:   Abre formulario GestionPersona
'   Param:  idPersona(long), identificador de persona
'
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
Public Function openFrmGestionPersona(idPersona As Long)

    'Dim db As database
    Dim rst As dao.Recordset
    Dim strSql As String
    
    Dim frm As New Form_GestionPersona
    
    If (idPersona <> 0) Then
        frm.setIdPersona (idPersona)
        frm.Filter = "id=" & idPersona
        frm.FilterOn = True
        frm.actualizaForm
    End If
    
    'frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
End Function


Public Function tieneNacionalidad(fkPersona As Long) As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim respuesta As Boolean
    
    str = " SELECT fkPersona, fkPaisNacimiento" & _
          " FROM t_datospersona" & _
          " WHERE (fkPersona=" & fkPersona & ");"
          
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        If (Len(Nz(rs!fkPaisNacimiento, "")) > 0) Then
            respuesta = True
        Else
            respuesta = False
        End If
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    tieneNacionalidad = respuesta
    
End Function

Public Function tieneDireccion(fkPersona As Long) As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim respuesta As Boolean
    
    str = " SELECT fkPersona, direccion, numero, fkCp" & _
          " FROM t_direccion" & _
          " WHERE (fkPersona=" & fkPersona & ");"
          
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        If (Len(Nz(rs!direccion, "")) > 0) And _
            (Len(Nz(rs!numero, "")) > 0) And _
            (Len(Nz(rs!fkCp, "")) > 0) Then
            respuesta = True
        Else
            respuesta = False
        End If
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    tieneDireccion = respuesta
    
End Function

'--------------------------------------------------------------------------------------------
'       Comoprobar datos obligatorios de persona
'--------------------------------------------------------------------------------------------
Public Function datosPersonalesRellenos(fkPersona As Long) As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim respuesta As Boolean
    
    str = " SELECT id, nombre, apellido1, apellido2, fechaNacimiento, fkSexo" & _
          " FROM t_persona" & _
          " WHERE (id=" & fkPersona & ");"
          
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        If (Len(Nz(rs!nombre, "")) > 0) And _
            (Len(Nz(rs!apellido1, "")) > 0) And _
            (Len(Nz(rs!fechaNacimiento, "")) > 0) And _
            (Len(Nz(rs!fkSexo, "")) > 0) Then
            respuesta = True
        Else
            respuesta = False
        End If
        
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    datosPersonalesRellenos = respuesta
    
End Function

Public Function tieneEstudios(fkPersona As Long) As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim respuesta As Boolean
    
    str = " SELECT fkPersona, fkTitulacion, fkEstadoFormacion, fechaFin" & _
          " FROM t_formacionReglada" & _
          " WHERE (fkPersona=" & fkPersona & ");"
          
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        If (Len(Nz(rs!fkTitulacion, "")) > 0) And _
            (Len(Nz(rs!fkEstadoFormacion, "")) > 0) And _
            (Len(Nz(rs!fechaFin, "")) > 0) Then
            respuesta = True
        Else
            respuesta = False
        End If
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    tieneEstudios = respuesta
    
End Function

Public Function tieneTelefono(fkPersona As Long) As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim respuesta As Boolean
    
    str = " SELECT fkPersona, telefono" & _
          " FROM t_telefono" & _
          " WHERE (fkPersona=" & fkPersona & ") AND fkTelefonoTipo=1;" 'fkTelefonoTipo=1(persona)
    
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        respuesta = True
    Else
        respuesta = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    tieneTelefono = respuesta
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  30/08/2012 - Actualización:  30/08/2012
'   Name:   isEmailNewsletterActivo
'   Desc:   Nos indica si algún mail tiene desactivado la opción de
'           newsletter
'   Param:  idPersona(long), identificador de persona
'   Retur:   0, no quiere newsletters
'           -1, quiere newsletters
'---------------------------------------------------------------------------
Public Function isEmailNewsletterActivo(idPersona As Long) As Integer
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim news As Integer
    Dim respuesta As Boolean
    
    str = " SELECT fkPersona, email, newsletter" & _
          " FROM t_email" & _
          " WHERE (fkPersona=" & idPersona & ") AND newsletter=0 AND not isnull(fkPersona)" & _
          " ORDER BY newsletter ASC"
    
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        news = rs!newsletter
    Else
        news = -1
    End If
    isEmailNewsletterActivo = news
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  30/08/2012 - Actualización:  30/08/2012
'   Name:   setEmailNewsletter
'   Desc:   Actualiza el campo newsletter de emails de una persona
'   Param:  idPersona(long), identificador de persona
'           newsletter(boolean), indica si quiere newsletters
'   Retur:   -
'---------------------------------------------------------------------------
Public Function setEmailNewsletter(idPersona As Long, _
                                   newsletter As Boolean)
    Dim str As String
    Dim news As Integer
    
    If (newsletter = True) Then
        news = -1
    Else
        news = 0
    End If
    
    str = "UPDATE t_email SET newsletter=" & news & " WHERE fkPersona=" & idPersona
    CurrentDb.Execute str
    
End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  04/04/2014 - Actualización: 04/04/2014
'   Name:   isPersonCompleted
'   Desc:   Check mandatory fields of person
'   Param:  idPersona(long), identificador de persona
'   Retur:  Boolean,    true,  ok
'                       false, ko
'---------------------------------------------------------------------------------
Public Static Function isPersonCompleted(idPersona As Long) As Boolean
    Dim isCompleted As Boolean
    isCompleted = True
    
    If isCompleted And Not hasPersonalMail(idPersona) Then
        isCompleted = False
    End If
    
    isPersonCompleted = isCompleted
End Function

