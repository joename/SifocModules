Attribute VB_Name = "SIFOC_WebPersona"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  25/11/2009 - Actualización:
'   Name:   insDatosPersonales
'   Desc:   Function para insertar datos personales
'   Param:  nombre As String,apellido1 As String,apellido2 As String,idSexo As Integer
'                                   DNI As String
'                                  fechaNac
'   Retur:  insDatosPersonales(long)
'---------------------------------------------------------------------------
Public Function insDatosPersonales(idWifoc As Long, _
                                   nombre As String, _
                                   apellido1 As String, _
                                   apellido2 As String, _
                                   idSexo As Integer, _
                                   dni As String, _
                                   fechaNac As Date, _
                                   idIfocUsuario As Long) As Integer
    Dim strSql As String
    Dim fields, values As String
On Error GoTo TratarError
    
    If (idWifoc = 0 Or nombre = "" Or apellido1 = "" Or idSexo = 0 Or dni = "" Or Not IsDate(fechaNac)) Then
        Exit Function
    End If
    
    fields = " idWifoc" & _
             ", nombre" & _
             ", apellido1" & _
             IIf(apellido2 <> "", ", apellido2", "") & _
             ", fkSexo" & _
             ", dni" & _
             ", fechanacimiento" & _
             ", fechaAlta" & _
             ", fkIfocUsuario"
    values = idWifoc & _
             ", '" & PrimeraLetraPalabraMayuscula(nombre) & "'" & _
             ", '" & PrimeraLetraPalabraMayuscula(apellido1) & "'" & _
             IIf(apellido2 <> "", ", '" & PrimeraLetraPalabraMayuscula(apellido2) & "'", "") & _
             ", " & idSexo & _
             ", '" & dni & "'" & _
             ", '" & fechaNac & "'" & _
             ", now()" & _
             ", " & idIfocUsuario
             
    strSql = " INSERT INTO t_persona (" & fields & ")" & _
             " VALUES (" & values & ");"
    
    CurrentDb.Execute strSql
    Debug.Print strSql
SalirError:
    insDatosPersonales = 0
    Exit Function
TratarError:
    insDatosPersonales = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  27/11/2009 - Actualización:
'   Name:   insDatosSeguridadSocial
'   Desc:   Inserta seguridadSocial en la tabla t_email de la persona pasada en los paramentros
'   Param:  fkPersona,seguridadSocial(String)
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosSeguridadSocial(fkPersona As Long, _
                                        seguridadSocial As String) As Integer
    Dim strSql As String
    Dim id As Integer
    
On Error GoTo TratarError

    strSql = " INSERT INTO t_seguridadsocial (fkPersona, seguridadSocial)" & _
             " VALUES (" & fkPersona & ", """ & filterSQL(seguridadSocial) & """);"
    
    CurrentDb.Execute strSql
     Debug.Print strSql
 
SalirError:
    insDatosSeguridadSocial = 0
    Exit Function
TratarError:
    insDatosSeguridadSocial = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P. Actualización: José Manuel Sánchez
'   Fecha:  25/11/2009 - Actualización: 25/03/2011
'   Name:   insDatosPesona
'   Desc:   Insertamos registro en datos persona
'   Param:
'   Retur:  0, Ok
'           -1, error
'---------------------------------------------------------------------------
Public Function insDatosPersona(idPersona As Long, _
                                idPaisNacimiento As Integer, _
                                idPaisNacionalidad As Integer, _
                                fechaResidenciaMunicipio As Date, _
                                Optional lstVehiculos As String = "") As Integer
    Dim strSql As String
    Dim fields As String
    Dim values As String
    
    Dim vehiculo, vehiculos
    
    Dim id As Integer
    
On Error GoTo TratarError
    
    fields = " fkPersona," & _
             IIf(idPaisNacimiento <> 0, " fkPaisNacimiento,", "") & _
             IIf(idPaisNacionalidad <> 0, " fkPaisNacionalidad ,", "") & _
             IIf(IsDate(fechaResidenciaMunicipio), " fechaResidenciaMunicipio", "")
    
    values = idPersona & ", " & _
             IIf(idPaisNacimiento <> 0, idPaisNacimiento & ", ", "") & _
             IIf(idPaisNacionalidad <> 0, idPaisNacionalidad & ", ", "") & _
             IIf(IsDate(fechaResidenciaMunicipio), "'" & Format(fechaResidenciaMunicipio, "dd/mm/yyyy") & "'", "")
    
    If (Len(lstVehiculos) > 0) Then
        vehiculos = Split(lstVehiculos, "#")
        For Each vehiculo In vehiculos
            fields = fields & ", dispone" & vehiculo
            values = values & ", -1"
        Next
    End If
    
    strSql = " INSERT INTO t_datospersona (" & fields & ")" & _
             " VALUES (" & values & ");"
    
'Debug.Print strSql
    CurrentDb.Execute strSql

 
SalirError:
    insDatosPersona = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosPersona = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  25/11/2009 - Actualización:
'   Name:   insDatosDireccion
'   Desc:   Obtenemos id de la persona con dni pasado por parametro
'   Param:  DNI (string)
'   Retur:  idPersona(long) si existe, sino 0
'---------------------------------------------------------------------------
Public Function insDatosDireccion(idPersona As Long, _
                                  idTipoVia As Integer, _
                                  direccion As String, _
                                  idTipoDireccion As Integer, _
                                  codigoPostal As String, _
                                  numero As Integer, _
                                  idPoblacion As Integer, _
                                  idProvincia As Integer) As Integer
                                
    Dim strSql As String
    Dim fields As String
    Dim values As String
    
    Dim vehiculo, vehiculos
    
    Dim id As Integer
    
On Error GoTo TratarError
    
    'strSql = " INSERT INTO t_direccion (fkPersona, " & _
             " fkTipoVia, direccion, fkTipoDireccion," & _
             " codigoPostal,numero,fkPoblacion,fkProvincia)" & _
             " VALUES (" & fkPersona & ", " & fkTipoVia & ",'" & direccion & "'," & fkTipoDireccion & _
             ",'" & codigoPostal & "'," & numero & "," & fkPoblacion & "," & fkProvincia & ");"
             
    fields = " fkPersona" & _
             ", fkTipoVia" & _
             ", direccion" & _
             ", fkTipoDireccion" & _
             ", codigoPostal" & _
             ", numero" & _
             ", fkPoblacion" & _
             IIf(idProvincia <> 0, ", fkProvincia", "")
    
    values = idPersona & _
             ", " & idTipoVia & _
             ", '" & filterSQL(direccion) & "'" & _
             ", " & idTipoDireccion & _
             ", '" & codigoPostal & "'" & _
             ", " & numero & _
             ", " & idPoblacion & _
             IIf(idProvincia <> 0, ", " & idProvincia, "")
    
    strSql = " INSERT INTO t_direccion (" & fields & ")" & _
             " VALUES (" & values & ");"
                 
Debug.Print strSql
    CurrentDb.Execute strSql
    
 
SalirError:
    insDatosDireccion = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosDireccion = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  25/11/2009 - Actualización:
'   Name:   insDatosTelefono
'   Desc:   Obtenemos id de la persona con dni pasado por parametro
'   Param:  DNI (string)
'   Retur:  idPersona(long) si existe, sino 0
'---------------------------------------------------------------------------
Public Function insDatosTelefono(idPersona As Long, _
                                 telefono As Long, _
                                 idIfocUsuario As Long, _
                                 Optional idTelefonoTipo As Integer = 0, _
                                 Optional idTipoTelefono1 As Integer = 0, _
                                 Optional idTipoTelefono2 As Integer = 0) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
On Error GoTo TratarError

    If esTelefono(Trim(str(telefono))) Then
        fields = " fkPersona" & _
                 ", telefono" & _
                 IIf(idTelefonoTipo <> 0, ", fkTelefonoTipo", "") & _
                 IIf(idTipoTelefono1 <> 0, ", fkTipoTelefono1", "") & _
                 IIf(idTipoTelefono2 <> 0, ", fkTipoTelefono2", "") & _
                 ", fkIfocUsuario"
        values = idPersona & _
                 ", " & telefono & _
                 IIf(idTelefonoTipo <> 0, ", " & idTelefonoTipo, "") & _
                 IIf(idTipoTelefono1 <> 0, ", " & idTipoTelefono1, "") & _
                 IIf(idTipoTelefono2 <> 0, ",  " & idTipoTelefono2, "") & _
                 ", " & idIfocUsuario

        strSql = " INSERT INTO t_telefono (" & fields & ")" & _
                 " VALUES (" & values & ");"
    
        CurrentDb.Execute strSql
    End If
    
'Debug.Print strSql
 
SalirError:
    insDatosTelefono = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosTelefono = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  27/11/2009 - Actualización:
'   Name:   insDatosTelefono
'   Desc:   Inserta el email en la tabla t_email de la persona pasada en los paramentros
'   Param:  fkPersona,fkEmailTipo,fkIfocUsuario(Integer),email(String)
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosEmail(idPersona As Long, _
                              idEmailTipo As Integer, _
                              email As String, _
                              idIfocUsuario As Long) As Integer
    Dim strSql As String
    Dim id As Integer
    
On Error GoTo TratarError
    
    If esEmail(email) Then
        strSql = " INSERT INTO t_email (fkPersona, fkEmailTipo, email, fkIfocUsuario)" & _
                 " VALUES (" & idPersona & ", " & idEmailTipo & ", '" & filterSQL(email) & "', " & idIfocUsuario & ");"
    
        CurrentDb.Execute strSql
    End If
'Debug.Print strSql
 
SalirError:
    insDatosEmail = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosEmail = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   autor upd: Jose Manuel Sanchez
'   Fecha:  27/11/2009 - Actualización: 20/4/2011
'   Name:   insDatosFormacionReglada
'   Desc:
'   Param:
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosFormacionReglada(idPersona As Long, _
                                        idNivelFormacion As Integer, _
                                        idEstadoFormacion As Integer, _
                                        idIfocUsuario As Long, _
                                        Optional idTitulacion As Integer = 0, _
                                        Optional idComunidadAutonoma As Integer = 0, _
                                        Optional idPais As Integer = 0, _
                                        Optional idMunicipio As Integer = 0, _
                                        Optional centro As String = "", _
                                        Optional homologado As Integer = 0, _
                                        Optional fechaFin As Date = "01/01/1900") As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    fields = " fkPersona" & _
             ", fkNivelFormacion" & _
             ", fkEstadoFormacion" & _
             ", fkIfocUsuario" & _
             IIf(idTitulacion <> 0, ", fkTitulacion", "") & _
             IIf(idComunidadAutonoma <> 0, ", fkComunidadAutonoma", "") & _
             IIf(idPais <> 0, ", fkPais", "") & _
             IIf(idMunicipio <> 0, ", fkMunicipio", "") & _
             IIf(centro <> "", ", centro", "") & _
             IIf(homologado <> 0, ", homologado", "") & _
             IIf(fechaFin <> "01/01/1900", ", fechaFin", "")
    
    values = idPersona & _
             ", " & idNivelFormacion & _
             ", " & idEstadoFormacion & _
             ", " & idIfocUsuario & _
             IIf(idTitulacion <> 0, ", " & idTitulacion, "") & _
             IIf(idComunidadAutonoma <> 0, ", " & idComunidadAutonoma, "") & _
             IIf(idPais <> 0, ", " & idPais, "") & _
             IIf(idMunicipio <> 0, ", " & idMunicipio, "") & _
             IIf(centro <> "", ", '" & filterSQL(centro) & "'", "") & _
             IIf(homologado <> 0, ", " & homologado, "") & _
             IIf(fechaFin <> "01/01/1900", ", '" & fechaFin & "'", "")
    
    strSql = " INSERT INTO t_formacionreglada (" & fields & ")" & _
             " VALUES (" & values & ");"
    
Debug.Print strSql
    CurrentDb.Execute strSql
 
SalirError:
    insDatosFormacionReglada = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosFormacionReglada = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   autor upd: Jose Manuel Sanchez
'   Fecha:  27/11/2009 - Actualización: 20/4/2011
'   Name:   insDatosFormacionNoReglada
'   Desc:
'   Param:
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosFormacionNoReglada(idPersona As Long, _
                                           idIfocUsuario As Long, _
                                           curso As String, _
                                           Optional centro As String = "", _
                                           Optional horas As String = "", _
                                           Optional fechaFin As Date = "01/01/1900", _
                                           Optional idComunidadAutonoma As Integer = 0, _
                                           Optional idPais As Integer = 0, _
                                           Optional idEstadoFormacion = 0) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    If (idPersona = 0 Or idIfocUsuario = 0 Or curso = "") Then
        Exit Function
    End If
    
    fields = " fkPersona" & _
             ", fkIfocUsuario" & _
             ", curso" & _
             IIf(centro <> "", ", centro", "") & _
             IIf(fechaFin <> "01/01/1900", ", fechaFin", "") & _
             IIf(horas <> "", ", horas", "") & _
             IIf(idComunidadAutonoma <> 0, ", fkComunidadAutonoma", "") & _
             IIf(idPais <> 0, ", fkPais", "") & _
             IIf(idEstadoFormacion <> 0, ", fkEstadoFormacion", "")
    
    values = idPersona & _
             ", " & idIfocUsuario & _
             ", '" & filterSQL(curso) & "'" & _
             IIf(centro <> "", ", '" & filterSQL(centro) & "'", "") & _
             IIf(fechaFin <> "01/01/1900", ", '" & fechaFin & "'", "") & _
             IIf(horas <> "", ", " & horas, "") & _
             IIf(idComunidadAutonoma <> 0, ", " & idComunidadAutonoma, "") & _
             IIf(idPais <> 0, ", " & idPais, "") & _
             IIf(idEstadoFormacion <> 0, ", " & idEstadoFormacion, "")
    
    strSql = " INSERT INTO t_formacionnoreglada (" & fields & ")" & _
             " VALUES (" & values & ");"
    
Debug.Print strSql
    CurrentDb.Execute strSql
    
SalirError:
    insDatosFormacionNoReglada = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosFormacionNoReglada = -1
End Function

'--------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   autor upd: Jose Manuel Sanchez
'   Fecha:  27/11/2009 - Actualización: 20/4/2011
'   Name:   insDatosCarneProfesional
'   Desc:
'   Param:
'   Retur: 0 si todo es correcto
'---------------------------------------------------------------------------
Public Function insDatosCarneProfesional(idPersona As Long, _
                                         idCarneProfesional As Integer, _
                                         idIfocUsuario As Long) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    If (idPersona = 0 Or idCarneProfesional = 0) Then
        insDatosCarneProfesional = -1
        Exit Function
    End If
    
    fields = " fkPersona" & _
             IIf(idCarneProfesional <> 0, ", fkCarneProfesional", "") & _
             IIf(idIfocUsuario <> 0, ", fkIfocUsuario", "")
    
    values = idPersona & _
             IIf(idCarneProfesional <> 0, ", " & idCarneProfesional, "") & _
             IIf(idIfocUsuario <> 0, ", " & idIfocUsuario, "")
    
    strSql = " INSERT INTO t_carneprofesional (" & fields & ")" & _
             " VALUES (" & values & ");"
    
    CurrentDb.Execute strSql
     Debug.Print strSql
 
SalirError:
    insDatosCarneProfesional = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosCarneProfesional = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   autor upd: Jose Manuel Sanchez
'   Fecha:  27/11/2009 - Actualización: 20/4/2011
'   Name:   insDatosIdiomas
'   Desc:
'   Param:
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosIdiomas(idPersona As Long, _
                                idIdioma As Integer, _
                                idIdiomaNivelSimple As Integer, _
                                idIfocUsuario As Long, _
                                Optional idCertificado As Integer, _
                                Optional idNivelIdiomaComprenderOral As Integer, _
                                Optional idNivelIdiomaComprenderEscrito As Integer, _
                                Optional idNivelIdiomaHablarInteraccion As Integer, _
                                Optional idNivelIdiomaHablarExpresion As Integer, _
                                Optional idNivelIdiomaEscribir As Integer, _
                                Optional observacion As String) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    If (idPersona = 0 Or idIdioma = 0 Or idIdiomaNivelSimple = 0 Or idIfocUsuario = 0) Then
        Exit Function
    End If
    
    fields = " fkPersona" & _
             ", fkIdioma" & _
             ", fkIdiomaNivelSimple" & _
             ", fkIfocUsuario" & _
             ", updDate" & _
             IIf(idCertificado <> 0, ", fkCertificado", "") & _
             IIf(idNivelIdiomaComprenderOral <> 0, ", fkNivelIdiomaComprenderOral", "") & _
             IIf(idNivelIdiomaComprenderEscrito <> 0, ", fkNivelIdiomaComprenderEscrito", "") & _
             IIf(idNivelIdiomaHablarInteraccion <> 0, ", fkNivelIdiomaHablarInteraccion", "") & _
             IIf(idNivelIdiomaHablarExpresion <> 0, ", fkNivelIdiomaHablarExpresion", "") & _
             IIf(observacion <> "", ", observacion", "")
    
    values = idPersona & _
             ", " & idIdioma & _
             ", " & idIdiomaNivelSimple & _
             ", " & idIfocUsuario & _
             ", now()" & _
             IIf(idCertificado <> 0, ", " & idCertificado, "") & _
             IIf(idNivelIdiomaComprenderOral <> 0, ", " & idNivelIdiomaComprenderOral, "") & _
             IIf(idNivelIdiomaComprenderEscrito <> 0, ", " & idNivelIdiomaComprenderEscrito, "") & _
             IIf(idNivelIdiomaHablarInteraccion <> 0, ", " & idNivelIdiomaHablarInteraccion, "") & _
             IIf(idNivelIdiomaHablarExpresion <> 0, ", " & idNivelIdiomaHablarExpresion, "") & _
             IIf(observacion <> "", ", '" & filterSQL(observacion) & "'", "")
    
    strSql = " INSERT INTO t_idioma (" & fields & ")" & _
             " VALUES (" & values & ");"
    
Debug.Print strSql
    CurrentDb.Execute strSql
 
SalirError:
    insDatosIdiomas = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosIdiomas = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   autor upd: Jose Manuel Sanchez
'   Fecha:  27/11/2009 - Actualización: 20/4/2011
'   Name:   insDatosInformatica
'   Desc:   Inserta los datos de informatica en la tabla t_informatica de la persona pasada en los paramentros
'   Param:  fkPersona,fkInformatica,fkNivel,fkIfocUsuario
'   Retur: 0 si todo es correcto
'---------------------------------------------------------------------------
Public Function insDatosInformatica(idPersona As Long, _
                                    idInformatica As Integer, _
                                    idNivel As Integer, _
                                    idIfocUsuario As Long) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    fields = " fkPersona" & _
             IIf(idInformatica <> 0, ", fkInformatica", "") & _
             IIf(idNivel <> 0, ", fkNivel", "") & _
             IIf(idIfocUsuario <> 0, ", fkIfocUsuario", "")
    
    values = idPersona & _
             IIf(idInformatica <> 0, ", " & idInformatica, "") & _
             IIf(idNivel <> 0, ", " & idNivel, "") & _
             IIf(idIfocUsuario <> 0, ", " & idIfocUsuario, "")
    
    strSql = " INSERT INTO t_informatica (" & fields & ")" & _
             " VALUES (" & values & ");"

Debug.Print strSql
    CurrentDb.Execute strSql
 
SalirError:
    insDatosInformatica = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosInformatica = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Autor upd: Jose Manuel Sanchez
'   Fecha:  27/11/2009 - Actualización: 16/05/2011
'   Name:   insDatosInformatica
'   Desc:   Inserta los datos de informatica en la tabla t_informatica de la persona pasada en los paramentros
'   Param:  fkPersona,fkInformatica,fkNivel,fkIfocUsuario
'   Retur: 0 si todo es correcto
'---------------------------------------------------------------------------
Public Function insDatosInsercion(idPersona As Long, _
                                  idIfocUsuario As Long, _
                                  fechaInicio As Date, _
                                  Optional fechaFin As Date = "01/01/1900", _
                                  Optional cargo As String = "", _
                                  Optional empresa As String = "", _
                                  Optional idComunidadAutonoma As Integer = 0, _
                                  Optional idPais As Integer = 0, _
                                  Optional idCno As Integer = 0) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    If (idPersona = 0 Or idIfocUsuario = 0 Or fechaInicio = "01/01/1900") Then
        insDatosInsercion = -1
        Exit Function
    End If
    
    fields = " fkPersona" & _
             ", fechaInicio" & _
             ", fkIfocUsuario" & _
             IIf(fechaFin <> "01/01/1900", ", fechaFin", "") & _
             IIf(cargo <> "", ", cargo", "") & _
             IIf(empresa <> "", ", empresa", "") & _
             IIf(idComunidadAutonoma <> 0, ", fkComunidadAutonoma", "") & _
             IIf(idPais <> 0, ", fkPais", "") & _
             IIf(idCno <> 0, ", fkCno2011", "")
    
    values = idPersona & _
             ", '" & fechaInicio & "'" & _
             ", " & idIfocUsuario & _
             IIf(fechaFin <> "01/01/1900", ", '" & fechaFin & "'", "") & _
             IIf(cargo <> "", ", '" & filterSQL(cargo) & "'", "") & _
             IIf(empresa <> "", ", '" & filterSQL(empresa) & "'", "") & _
             IIf(idComunidadAutonoma <> 0, ", " & idComunidadAutonoma, "") & _
             IIf(idPais <> 0, ", " & idPais, "") & _
             IIf(idCno <> 0, ", " & idCno, "")
    
    strSql = " INSERT INTO t_insercion (" & fields & ")" & _
             " VALUES (" & values & ");"

Debug.Print strSql
    CurrentDb.Execute strSql
    
SalirError:
    insDatosInsercion = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosInsercion = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   autor upd: Jose Manuel Sanchez
'   Fecha:  27/11/2009 - Actualización: 20/4/2011
'   Name:   insDatosConocimientos
'   Desc:   Inserta los datos de conocimiento en la tabla t_informatica de la persona pasada en los paramentros
'   Param:  fkPersona,fkInformatica,fkNivel,fkIfocUsuario
'   Retur: 0 si todo es correcto
'---------------------------------------------------------------------------
Public Function insDatosConocimiento(idPersona As Long, _
                                     idIfocUsuario As Long, _
                                     folletoIfoc As Integer, _
                                     folletoIfocDonde As String, _
                                     derivadoOtroServicio As Integer, _
                                     derivadoOtroServicioCual As String, _
                                     prensa As Integer, _
                                     internet As Integer, _
                                     familiaresamigos As Integer, _
                                     radio As Integer, _
                                     revistacalviaaldia As Integer, _
                                     otros As Integer, _
                                     otrosDescripcion As String) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    fields = " fkPersona" & _
             ", fkIfocUsuario" & _
             ", fecha" & _
             IIf(folletoIfoc <> 0, ", folletoifoc", "") & _
             IIf(folletoIfocDonde <> "", ", folletoIfocDonde", "") & _
             IIf(derivadoOtroServicio <> 0, ", derivadoOtroServicio", "") & _
             IIf(derivadoOtroServicioCual <> "", ", derivadoOtroServicioCual", "") & _
             IIf(prensa <> 0, ", prensa", "") & _
             IIf(internet <> 0, ", internet", "") & _
             IIf(familiaresamigos <> 0, ", familiaresamigos", "") & _
             IIf(radio <> 0, ", radio", "") & _
             IIf(revistacalviaaldia <> 0, ", revistacalviaaldia", "") & _
             IIf(otros <> 0, ", otros", "") & _
             IIf(otrosDescripcion <> "", ", otrosDescripcion", "")
    
    values = idPersona & _
             ", " & idIfocUsuario & _
             ", now()" & _
             IIf(folletoIfoc <> 0, "," & folletoIfoc, "") & _
             IIf(folletoIfocDonde <> "", ", '" & filterSQL(folletoIfocDonde) & "'", "") & _
             IIf(derivadoOtroServicio <> 0, ", " & derivadoOtroServicio, "") & _
             IIf(derivadoOtroServicioCual <> "", ", '" & derivadoOtroServicioCual & "'", "") & _
             IIf(prensa <> 0, ", " & prensa, "") & _
             IIf(internet <> 0, ", " & internet, "") & _
             IIf(familiaresamigos <> 0, ", " & familiaresamigos, "") & _
             IIf(radio <> 0, ", " & radio, "") & _
             IIf(revistacalviaaldia <> 0, ", " & revistacalviaaldia, "") & _
             IIf(otros <> 0, ", " & otros, "") & _
             IIf(otrosDescripcion <> "", ", '" & filterSQL(otrosDescripcion) & "'", "")

    strSql = " INSERT INTO t_conocimientoifoc (" & fields & ")" & _
             " VALUES (" & values & ");"
    
    'strSql = " INSERT INTO t_conocimientoifoc (fkPersona, fkIfocUsuario,fecha ,folletoIfoc, folletoIfocDonde, " & _
            " derivadoOtroServicio, derivadoOtroServicioCual,prensa, internet,familiaresAmigos,radio,revistaCalviaAldia,otros,otrosDescripcion)" & _
             " VALUES (" & fkPersona & ", " & fkIfocUsuario & ", now(), " & folletoIfoc & ", '" & folletoIfocDonde & "', " & derivadoOtroServicio & ", '" & derivadoOtroServicioCual & "', " & prensa & ", " & internet & ", " & familiaresamigos & ", " & radio & ", " & revistacalviaaldia & ", " & otros & ", '" & otrosDescripcion & "');"
    
Debug.Print strSql
    CurrentDb.Execute strSql
    
    
SalirError:
    insDatosConocimiento = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosConocimiento = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Autor upd: Jose Manuel Sanchez
'   Fecha:  26/04/2011 - Actualización: 26/04/2011
'   Name:   insDatosDemandaAyuda
'   Desc:
'   Param:  idPersona (long), identificador de persona
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosDemandaAyuda(idPersona As Long, _
                                     idPrestacionTipo As Integer, _
                                     Optional idPrestacionSubtipo As Integer = 0, _
                                     Optional fechaInicio As Date = "01/01/1900", _
                                     Optional fechaFin As Date = "01/01/1900", _
                                     Optional cantidad As Integer = 0, _
                                     Optional OBS As String = "") As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    If (idPersona = 0 Or idPrestacionTipo = 0) Then
        Exit Function
    End If
    
    fields = " fkPersona" & _
             ", fkPrestacionTipo" & _
             IIf(idPrestacionSubtipo <> 0, ", fkPrestacionSubtipo", "") & _
             IIf(fechaInicio > "01/01/1990", ", fechaInicio", "") & _
             IIf(fechaFin > "01/01/1990", ", fechaFin", "") & _
             IIf(cantidad <> 0, ", cantidad", "") & _
             IIf(OBS <> "", ", observaciones", "")
    
    values = idPersona & _
             ", " & idPrestacionTipo & _
             IIf(idPrestacionSubtipo <> 0, ", " & idPrestacionSubtipo, "") & _
             IIf(fechaInicio > "01/01/1990", ", '" & fechaInicio & "'", "") & _
             IIf(fechaFin > "01/01/1990", ", '" & fechaFin & "'", "") & _
             IIf(cantidad <> 0, ", " & cantidad, "") & _
             IIf(OBS <> "", ", '" & filterSQL(OBS) & "'", "")
    
    strSql = " INSERT INTO t_prestaciones (" & fields & ")" & _
             " VALUES (" & values & ");"
    
Debug.Print strSql
    CurrentDb.Execute strSql
    
SalirError:
    insDatosDemandaAyuda = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosDemandaAyuda = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Autor upd: Jose Manuel Sanchez
'   Fecha:  26/04/2011 - Actualización: 26/04/2011
'   Name:   insDatosInteresFormativo
'   Desc:
'   Param:  idPersona (long), identificador de persona
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosInteresFormacion(idPersona As Long, _
                                        idGrupoFormacion2 As Integer, _
                                        idIfocUsuario As Long, _
                                        fecha As Date, _
                                        Optional idCursoNivel As Integer = 0) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    If (idPersona = 0 Or idGrupoFormacion2 = 0 Or idIfocUsuario = 0 Or fecha = "01/01/1900") Then
        insDatosInteresFormacion = -1
        Exit Function
    End If
    
    fields = " fkPersona" & _
             ", fkIfocUsuario" & _
             ", fkGrupoFormacion2" & _
             ", fecha" & _
             IIf(idCursoNivel <> 0, ", fkCursoNivel", "")

    values = idPersona & _
             ", " & idIfocUsuario & _
             ", " & idGrupoFormacion2 & _
             ", '" & fecha & "'" & _
             IIf(idCursoNivel <> 0, ", " & idCursoNivel, "")

    strSql = " INSERT INTO t_interesformacion (" & fields & ")" & _
             " VALUES (" & values & ");"
    
Debug.Print strSql
    CurrentDb.Execute strSql
    
SalirError:
    insDatosInteresFormacion = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosInteresFormacion = -1
End Function


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Autor upd: Jose Manuel Sanchez
'   Fecha:  26/04/2011 - Actualización: 26/04/2011
'   Name:   insDatosOcupacionesDemandadas
'   Desc:
'   Param:  idPersona (long), identificador de persona
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosOcupacionesDemandadas(idPersona As Long, _
                                              idCno As Integer, _
                                              nivel As Integer, _
                                              Optional principal As Integer = 0, _
                                              Optional OBS As String = "", _
                                              Optional DEMANDA As Integer = 0, _
                                              Optional SOIB As Integer = 0, _
                                              Optional op As Integer = 0) As Integer
    Dim strSql As String
    Dim fields, values As String
    
On Error GoTo TratarError
    
    If (idPersona = 0 Or idCno = 0) Then
        insDatosOcupacionesDemandadas = -1
        Exit Function
    End If
    
    fields = " fkPersona" & _
             ", fkCno2011" & _
             ", nivel" & _
             IIf(OBS <> "", ", observacion", "") & _
             IIf(principal <> 0, ", principal", "") & _
             IIf(DEMANDA <> 0, ", demanda", "") & _
             IIf(SOIB <> 0, ", soib", "") & _
             IIf(op <> 0, ", objetivoProfesional", "")

    values = idPersona & _
             ", " & idCno & _
             ", " & nivel & _
             IIf(OBS <> "", ", '" & filterSQL(OBS) & "'", "") & _
             IIf(principal <> 0, ", " & principal, "") & _
             IIf(DEMANDA <> 0, ", " & DEMANDA, "") & _
             IIf(SOIB <> 0, ", " & SOIB, "") & _
             IIf(op <> 0, ", " & op, "")

    strSql = " INSERT INTO t_cnodebusqueda (" & fields & ")" & _
             " VALUES (" & values & ");"
    
Debug.Print strSql
    CurrentDb.Execute strSql
    
SalirError:
    insDatosOcupacionesDemandadas = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosOcupacionesDemandadas = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Autor upd: Jose Manuel Sanchez
'   Fecha:  26/04/2011 - Actualización: 26/04/2011
'   Name:   insDatosDisponibilidad
'   Desc:
'   Param:  idPersona (long), identificador de persona
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosDisponibilidad(idPersona As Long, _
                                       idDisponibilidadJornada As Integer, _
                                       idDisponibilidadDistribucionJornada As Integer, _
                                       Optional OBS As String = "") As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    If (idPersona = 0) Then
        insDatosDisponibilidad = -1
        Exit Function
    End If
    
    fields = " fkPersona" & _
             IIf(idDisponibilidadJornada <> 0, ", fkDisponibilidadJornada", "") & _
             IIf(idDisponibilidadDistribucionJornada <> 0, ", fkDisponibilidadDistribucionJornada", "") & _
             IIf(OBS <> "", ", observaciones", "")

    values = idPersona & _
             IIf(idDisponibilidadJornada <> 0, ", " & idDisponibilidadJornada, "") & _
             IIf(idDisponibilidadDistribucionJornada <> 0, ", " & idDisponibilidadDistribucionJornada, "") & _
             IIf(OBS <> "", ", '" & filterSQL(OBS) & "'", "")

    strSql = " INSERT INTO t_disponibilidad (" & fields & ")" & _
             " VALUES (" & values & ");"
    
Debug.Print strSql
    CurrentDb.Execute strSql
    
SalirError:
    insDatosDisponibilidad = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosDisponibilidad = -1
End Function


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Autor upd: Jose Manuel Sanchez
'   Fecha:  26/04/2011 - Actualización: 26/04/2011
'   Name:   insDatosDiscapacidad
'   Desc:
'   Param:  idPersona (long), identificador de persona
'   Retur:
'---------------------------------------------------------------------------
Public Function insDatosDiscapacidad(idPersona As Long, _
                                     idDiscapacidad As Integer) As Integer
    Dim strSql As String
    Dim id As Integer
    Dim fields, values As String
    
On Error GoTo TratarError
    
    If (idPersona = 0 Or idDiscapacidad = 0) Then
        insDatosDiscapacidad = -1
        Exit Function
    End If
    
    fields = " fkPersona" & _
             IIf(idDiscapacidad <> 0, ", fkDiscapacidad", "") & _
             ", principal"

    values = idPersona & _
             IIf(idDiscapacidad <> 0, ", " & idDiscapacidad, "") & _
             ", -1"

    strSql = " INSERT INTO t_discapacidad (" & fields & ")" & _
             " VALUES (" & values & ");"
    
Debug.Print strSql
    CurrentDb.Execute strSql
    
SalirError:
    insDatosDiscapacidad = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    insDatosDiscapacidad = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  27/11/2009 - Actualización:
'   Name:   upDateTWebPersona
'   Desc:   Actualiza la tabla T_WEBORGANIZACION y agrega el id dado en sifoc en wifoc
'   Param: idOrganizacionSIFOC,idOrganizacionWifoc
'   Retur: 0 si todo es correcto
'---------------------------------------------------------------------------
Public Function updTWebPersona(idPersonaSIFOC As Long, idPersonaWifoc As Integer) As Integer
    Dim strSql As String
    Dim id As Integer
    
On Error GoTo TratarError
    
    strSql = " UPDATE t_webpersona" & _
             " SET idPersonaSIFOC=" & idPersonaSIFOC & ", estadoPersona=-1" & _
             " WHERE id=" & idPersonaWifoc & ""
    
    'CurrentDb.Execute strSql
    
    Debug.Print strSql
    CurrentDb.Execute strSql
SalirError:
    updTWebPersona = 0
    Exit Function
TratarError:
    Debug.Print Err.description
    updTWebPersona = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  25/11/2009 - Actualización:
'   Name:   getIdUsuarioConDni
'   Desc:   Obtenemos id de la persona con dni pasado por parametro
'   Param:  DNI (string)
'   Retur:  idPersona(long) si existe, sino 0
'---------------------------------------------------------------------------
Public Function getIdUsuarioConDni(dni As String) As Long
    
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim idPersona As Long
On Error GoTo TratarError
    str = " SELECT id, dni" & _
          " FROM t_persona" & _
          " WHERE dni='" & dni & "';"
  
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
       
    If Not (rs.EOF) Then 'existe dni persona en sifoc
        rs.MoveFirst
        getIdUsuarioConDni = rs!id
    Else 'no existe persona
        getIdUsuarioConDni = 0
    End If

    rs.Close
    Set rs = Nothing
SalirError:
    Exit Function
TratarError:
    Debug.Print Err.description
    getIdUsuarioConDni = 0
End Function
