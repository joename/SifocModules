Attribute VB_Name = "SIFOC_PersonaCV"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  23/03/2010 - Actualización:  23/03/2010
'   Name:   insCarneProfesional
'   Desc:   Insertamos los carnés profesionales en la tabla local
'           L_CVPersonalizado con el tipo correspondiente (1-carne prof.)
'   Param:  -
'   Retur:  0, ok
'          -1, ko
'---------------------------------------------------------------------------
Public Function insCVCarneProfesional(idPersona As Long) As Integer
    Dim str As String
    Dim resultado As Integer

On Error GoTo TratarError
    
    resultado = 0
    
    str = " INSERT INTO L_CVPersonalizado (fkPersona, fkTipo, fechaFin, lugarConocimiento)" & _
          " SELECT t_carneprofesional.fkPersona, 1, t_carneprofesional.fechaCaducidad, [area] & ' - ' & [carneProfesional] & '  ' & [nivel] AS carne" & _
          " FROM (a_carneprofesionalnivel RIGHT JOIN (t_carneprofesional LEFT JOIN a_carneprofesional ON t_carneprofesional.fkCarneProfesional = a_carneprofesional.id) ON a_carneprofesionalnivel.id = t_carneprofesional.fkCarneProfesionalNivel) LEFT JOIN a_carneprofesionalarea ON a_carneprofesional.fkCarneProfesionalArea = a_carneprofesionalarea.id" & _
          " WHERE (a_carneprofesional.fkCarneProfesionalArea <> 2) AND t_carneprofesional.fkPersona = " & idPersona
    
    CurrentDb.Execute str
    
    insCVCarneProfesional = resultado
SalirTratarError:
    Exit Function
TratarError:
    Debug.Print Err.description
    insCVCarneProfesional = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  23/03/2010 - Actualización:  23/03/2010
'   Name:   insFormacionReglada
'   Desc:   Insertamos los formación reglada en la tabla local
'           L_CVPersonalizado con el tipo correspondiente (2-Form. reg)
'   Param:  -
'   Retur:  0, ok
'          -1, ko
'---------------------------------------------------------------------------
Public Function insCVFormacionReglada(idPersona As Long) As Integer
    Dim str As String
    Dim resultado As Integer

On Error GoTo TratarError
    
    resultado = 0
    
    str = " INSERT INTO L_CVPersonalizado (fkPersona, fkTipo, fechaFin, lugarConocimiento, tituloCurso, estadoFormacion)" & _
          " SELECT T_FormacionReglada.fkPersona, 2, T_FormacionReglada.fechaFin, T_FormacionReglada.centro, titulacion AS titulo, a_estadoformacion.descripcion AS estadoFormacion" & _
          " FROM (a_nivelformacion LEFT JOIN a_nivelformacionsoib ON a_nivelformacion.fkNivelFormacionSoib = a_nivelformacionsoib.id) INNER JOIN ((T_FormacionReglada LEFT JOIN T_Titulacion ON T_FormacionReglada.fkTitulacion = T_Titulacion.id) LEFT JOIN a_estadoformacion ON T_FormacionReglada.fkEstadoFormacion = a_estadoformacion.id) ON a_nivelformacion.id = T_FormacionReglada.fkNivelFormacion" & _
          " WHERE t_formacionreglada.fkPersona = " & idPersona & " AND (t_formacionreglada.fkEstadoformacion = 1 OR t_formacionreglada.fkEstadoformacion = 2) " & _
          " ORDER BY T_FormacionReglada.fechaFin DESC;"
    
    CurrentDb.Execute str
    
    insCVFormacionReglada = resultado
    
SalirTratarError:
    Exit Function
TratarError:
    insCVFormacionReglada = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  23/03/2010 - Actualización:  23/03/2010
'   Name:   insFormacionOcupacional
'   Desc:   Insertamos los formación no reglada en la tabla local
'           L_CVPersonalizado con el tipo correspondiente (3-Form. reg)
'   Param:  -
'   Retur:  0, ok
'          -1, ko
'---------------------------------------------------------------------------
Public Function insCVFormacionOcupacional(idPersona As Long) As Integer
    Dim str As String
    Dim resultado As Integer

On Error GoTo TratarError
    
    resultado = 0
    
    str = " INSERT INTO L_CVPersonalizado (fkPersona, fkTipo, fechaFin, lugarConocimiento, estadoFormacion, tituloCurso, horas, CertificadoProfesionalidad, funciones)" & _
          " SELECT T_FormacionNoReglada.fkPersona, 3, T_FormacionNoReglada.fechaFin,  T_FormacionNoReglada.centro,  a_estadoformacion.descripcion AS estadoFormacion, T_FormacionNoReglada.curso, T_FormacionNoReglada.horas, a_grupoformacion2.grupo2 as CertificadoProfesionalidad, T_FormacionNoReglada.contenidos" & _
          " FROM ((T_FormacionNoReglada LEFT JOIN a_estadoformacion ON T_FormacionNoReglada.fkEstadoFormacion = a_estadoformacion.id) LEFT JOIN a_grupoformacion2 ON t_formacionnoreglada.fkGrupoFormacion2 = a_grupoformacion2.id)" & _
          " WHERE (t_formacionnoreglada.fkPersona= " & idPersona & ") AND (T_FormacionNoReglada.fkEstadoFormacion = 1 Or T_FormacionNoReglada.fkEstadoFormacion = 2)" & _
          " ORDER BY T_FormacionNoReglada.fechaFin DESC;"
    
    CurrentDb.Execute str
Debug.Print str
    insCVFormacionOcupacional = resultado
SalirTratarError:
    Exit Function
TratarError:
    insCVFormacionOcupacional = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/03/2010 - Actualización:  24/03/2010
'   Name:   insExperiencia
'   Desc:   Insertamos la experiencia en la tabla local
'           L_CVPersonalizado con el tipo correspondiente (4-Experiencia)
'   Param:  -
'   Retur:  0, ok
'          -1, ko
'---------------------------------------------------------------------------
Public Function insCVExperiencia(idPersona As Long) As Integer
    Dim str As String
    Dim resultado As Integer

On Error GoTo TratarError
    
    resultado = 0
    
    str = " INSERT INTO L_CVPersonalizado (fkPersona, fkTipo, fechaInicio, fechaFin, lugarConocimiento, cargoCno, funciones)" & _
          " SELECT T_Insercion.fkPersona, 4, T_Insercion.fechaInicio, T_Insercion.fechaFin, T_Insercion.empresa, IIf([cargo] Is Null,[ocupacion],[cargo]) AS puesto, T_Insercion.funcion" & _
          " FROM T_Insercion INNER JOIN A_CNO2011 ON T_Insercion.fkCno2011 = A_CNO2011.id" & _
          " WHERE T_Insercion.fkPersona = " & idPersona & _
          " ORDER BY T_Insercion.fechaInicio DESC , T_Insercion.fechaFin DESC;"
    
    CurrentDb.Execute str
    
    insCVExperiencia = resultado
SalirTratarError:
    Exit Function
TratarError:
    insCVExperiencia = -1
    Resume SalirTratarError
End Function


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/03/2010 - Actualización:  24/03/2010
'   Name:   insIdioma
'   Desc:   Insertamos la experiencia en la tabla local
'           L_CVPersonalizado con el tipo correspondiente (5-Idioma)
'   Param:  -
'   Retur:  0, ok
'          -1, ko
'---------------------------------------------------------------------------
Public Function insCVIdioma(idPersona As Long) As Integer
    Dim str As String
    Dim resultado As Integer

On Error GoTo TratarError
    
    resultado = 0
    
    str = " INSERT INTO L_CVPersonalizado (fkPersona, fkTipo, lugarConocimiento, cargoCno, tituloCurso, funciones)" & _
          " SELECT T_Idioma.fkPersona, 5, A_Idioma.idioma, a_idiomanivelsimple.nivel, a_certificado.certificado, iif([lenguaMaterna]=-1,'Lengua materna','') as materna" & _
          " FROM a_certificado RIGHT JOIN (a_idiomanivelsimple RIGHT JOIN (T_Idioma LEFT JOIN A_Idioma ON T_Idioma.fkIdioma = A_Idioma.id) ON a_idiomanivelsimple.id = T_Idioma.fkIdiomaNivelSimple) ON a_certificado.id = T_Idioma.fkCertificado" & _
          " WHERE t_idioma.fkPersona = " & idPersona
    
    CurrentDb.Execute str
    
    insCVIdioma = resultado
SalirTratarError:
    Exit Function
TratarError:
    insCVIdioma = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/03/2010 - Actualización:  24/03/2010
'   Name:   insIdioma
'   Desc:   Insertamos la experiencia en la tabla local
'           L_CVPersonalizado con el tipo correspondiente (6-Informatica)
'   Param:  -
'   Retur:  0, ok
'          -1, ko
'---------------------------------------------------------------------------
Public Function insCVInformatica(idPersona As Long) As Integer
    Dim str As String
    Dim resultado As Integer

On Error GoTo TratarError
    
    resultado = 0
    
    str = " INSERT INTO L_CVPersonalizado (fkPersona, fkTipo,  lugarConocimiento, cargoCno)" & _
          " SELECT t_informatica.fkPersona, 6, a_informatica.informatica, a_nivel.nivel" & _
          " FROM a_informatica RIGHT JOIN (t_informatica LEFT JOIN a_nivel ON t_informatica.fkNivel = a_nivel.id) ON a_informatica.id = t_informatica.fkInformatica" & _
          " WHERE t_informatica.fkPersona = " & idPersona
    
    CurrentDb.Execute str
    
    insCVInformatica = resultado
SalirTratarError:
    Exit Function
TratarError:
    insCVInformatica = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  25/03/2010 - Actualización:  25/03/2010
'   Name:   delTablaLocalCVPersonalizado
'   Desc:   Eliminamos los datos de la tabla local L_CVPersonalizado
'   Param:  -
'   Retur:  0, ok
'          -1, ko
'---------------------------------------------------------------------------
Public Function delTablaLocalCVPersonalizado() As Integer
    Dim str As String
    Dim resultado As Integer

On Error GoTo TratarError
    
    resultado = 0
    
    str = " DELETE * FROM L_CVPersonalizado;"
          
    CurrentDb.Execute str
    
    delTablaLocalCVPersonalizado = resultado
SalirTratarError:
    Exit Function
TratarError:
    delTablaLocalCVPersonalizado = -1
    Resume SalirTratarError
End Function


'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/03/2010 - Actualización:  15/03/2010
'   Name:   direccionCV
'   Desc:   Devuelve la dirección de la persona pasada por parámetro
'           Tenemos en cuenta la 1ªresidencia y en caso de no tener la 2ª res.
'           (esta consulta debe coincidir con la de direccion)
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con
'---------------------------------------------------------------------------------
Public Function direccionCV(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    
    strSql = " SELECT a_tipovia.tipoVia, t_direccion.direccion, t_direccion.bis, t_direccion.numero, t_direccion.bloque, t_direccion.escalera, t_direccion.piso, t_direccion.puerta" & _
             " FROM t_direccion LEFT JOIN a_tipovia ON t_direccion.fkTipoVia = a_tipovia.id" & _
             " WHERE t_direccion.fkPersona = " & idPersona & "" & _
             " ORDER BY t_direccion.fkTipoDireccion ASC"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    str = ""
    If Not rs.EOF Then
        rs.MoveFirst
        'Montamos string dirección
        str = IIf(Not IsNull(rs!TipoVia), rs!TipoVia, "")
        str = str & IIf(Not IsNull(rs!direccion), " " & rs!direccion, "")
        str = str & IIf(rs!numero <> 0, ", num. " & rs!numero, "")
        str = str & IIf(Not IsNull(rs!bis), " bis: " & rs!bis, "")
        str = str & IIf(Not IsNull(rs!bloque), " bloque: " & rs!bloque, "")
        str = str & IIf(Not IsNull(rs!escalera), " esc: " & rs!escalera, "")
        str = str & IIf(Not IsNull(rs!piso), " piso: " & rs!piso, "")
        str = str & IIf(Not IsNull(rs!puerta), " puerta: " & rs!puerta, "")
    Else
        str = "No especificada"
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    direccionCV = str

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
Public Function localidadCV(idPersona As Long) As String
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
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    localidadCV = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/03/2010 - Actualización:  15/03/2010
'   Name:   telefonosCV
'   Desc:   Devuelve los teléfonos
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con teléfonos separados por comas.
'---------------------------------------------------------------------------------
Public Function telefonosCV(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    'Sólo miramos telefonos personales (fijo y móvil)
    strSql = " SELECT t_telefono.fkPersona, t_telefono.telefono" & _
             " FROM t_telefono" & _
             " WHERE (t_telefono.fkTelefonoTipo = 1) And (t_telefono.fkTipoTelefono1 = 1) And (t_telefono.fkTipoTelefono2 <> 2)  And (t_telefono.fkPersona =" & idPersona & ")" & _
             " ORDER BY  t_telefono.fkTipoTelefono2 DESC;"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        For i = 1 To 2 Step 1
            'Montamos string localidad
            If Not rs.EOF Then
                str = str & IIf(str = "", rs!telefono, ", " & rs!telefono)
                rs.MoveNext
            End If
        Next i
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    telefonosCV = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/03/2010 - Actualización:  17/03/2010
'   Name:   emailsCV
'   Desc:   Devuelve los emails de la persona pasada por parámetro
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con emails separados por comas.
'---------------------------------------------------------------------------------
Public Function emailsCV(idPersona As Long) As String
    Dim str As String
    Dim strSql As String
    Dim i As Integer
    
    strSql = " SELECT t_email.email" & _
             " FROM t_email" & _
             " WHERE not isnull(t_email.fkPersona) AND (t_email.fkPersona =" & idPersona & ");"
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    str = ""
    
    If Not rs.EOF Then
        rs.MoveFirst
        For i = 1 To 2 Step 1
            'Montamos string localidad
            If Not rs.EOF Then
                str = str & IIf(str = "", rs!email, ", " & rs!email)
                rs.MoveNext
            End If
        Next i
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    emailsCV = str

End Function

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/03/2010 - Actualización:  17/03/2010
'   Name:   carnesConducirCV
'   Desc:   Devuelve los carnes profesionales relacionados con área Tráfico
'           idCarneProfesionalArea(2)
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con carnes de conducir separados por comas.
'---------------------------------------------------------------------------------
Public Function carnesConducirCV(idPersona As Long) As String
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
    End If
    
    'Cerramos rs
    rs.Close
    Set rs = Nothing
    
    carnesConducirCV = str & IIf(disponeVehiculoCV(idPersona) = "", "", " con " & disponeVehiculoCV(idPersona))

End Function


'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/03/2010 - Actualización:  17/03/2010
'   Name:   disponeVehiculoCV
'   Desc:   Devuelve los vehiculos que dispone la persona
'   Param:  idPersona(long), identificador de persona
'   Retur:  String con vehículos que dispone la persona separados por comas.
'---------------------------------------------------------------------------------
Private Function disponeVehiculoCV(idPersona As Long) As String
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
    
    disponeVehiculoCV = str

End Function

