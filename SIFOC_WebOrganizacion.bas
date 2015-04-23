Attribute VB_Name = "Sifoc_WebOrganizacion"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  30/12/2009 - Actualización:
'   Name:   insDatosEmpresa
'   Desc:   Function para insertar datos insDatosEmpresa
'   Param:  nombre,razonSocial,fechaCreacion,fkActividadEmpresarial,cif,telefono
'           fax,email,web,fkPlantillaMedia,esAmbitoMunicipal,esAmbitoAutonomico
'           esAmbitoEstatal,esAmbitoInternacional,numCentrosCalvia,numCentrosPalma
'           numCentrosMallorca,numCentrosBaleares,numCentrosEspana,numCentrosInternacional
'   Retur:  insDatosEmpresa(long)
'---------------------------------------------------------------------------

Public Function insDatosEmpresa(nombre As String, razonSocial As String, _
                                   fechaCreacion As Date, fkActividadEmpresarial As Integer, _
                                   cif As String, telefono As Long, fax As Long, _
                                   email As String, _
                                   web As String, fkPlantillaMedia As Integer, _
                                   esAmbitoMunicipal As Integer, esAmbitoAutonomico As Integer, _
                                   esAmbitoEstatal As Integer, esAmbitoInternacional As Integer, _
                                   numCentrosCalvia As Integer, numCentrosPalma As Integer, _
                                   numCentrosMallorca As Integer, numCentrosBaleares As Integer, _
                                   numCentrosEspana As Integer, numCentrosInternacional As Integer, _
                                   SS As String, _
                                   fkIFOCUsuario As Integer, domicilio As String, _
                                   numero As Integer, codigoPostal As Integer, _
                                   fkPoblacion As Integer, fkMunicipio As Integer, _
                                   fkProvincia As Integer, fkTipoCentroTrabajo As Integer, _
                                   bis As String, bloque As String, piso As String, _
                                   fkTipoVia As Integer, escalera As String, _
                                   puerta As String) As Integer
                                   
                                 
    
    Dim strSql As String
    Dim strValues As String
    Dim sql As String
On Error GoTo TratarError
    
     strSql = " INSERT INTO t_organizacion (nombre,razonSocial,CIF,fkIfocUsuario,fechaInscripcion" & _
                IIf(fkActividadEmpresarial = 0, "", ",fkActividadEmpresarial") & _
                IIf(fechaCreacion < 1, "", ",FechaCreacion") & IIf(telefono = 0, "", ",telefono") & _
                IIf(fax = 0, "", ",fax") & IIf(email = "", "", ",email") & IIf(web = "", "", ",web") & IIf(fkPlantillaMedia = 0, "", ",fkPlantillaMedia") & _
                IIf(esAmbitoMunicipal = 0, "", ",esAmbitoMunicipal") & IIf(esAmbitoAutonomico = 0, "", ",esAmbitoAutonomico") & _
                IIf(esAmbitoEstatal = 0, "", ",esAmbitoEstatal") & IIf(esAmbitoInternacional = 0, "", ",esAmbitoInternacional") & _
                IIf(numCentrosCalvia = 0, "", ",numCentrosCalvia") & IIf(numCentrosPalma = 0, "", ",numCentrosPalma") & _
                IIf(numCentrosMallorca = 0, "", ",numCentrosMallorca") & IIf(numCentrosBaleares = 0, "", ",numCentrosBaleares") & _
                IIf(numCentrosEspana = 0, "", ",numCentrosEspana") & IIf(numCentrosInternacional = 0, "", ",numCentrosInternacional") & _
                IIf(SS = "", "", ",SS") & IIf(domicilio = "", "", ",domicilio") & _
                IIf(numero = 0, "", ",numero") & IIf(codigoPostal = 0, "", ",codigoPostal") & _
                IIf(fkPoblacion = 0, "", ",fkPoblacion") & IIf(fkMunicipio = 0, "", ",fkMunicipio") & _
                IIf(fkProvincia = 0, "", ",fkProvincia") & IIf(fkTipoCentroTrabajo = 0, "", ",fkTipoCentroTrabajo") & _
                IIf(bis = "", "", ",bis") & IIf(bloque = "", "", ",bloque") & IIf(piso = "", "", ",piso") & _
                IIf(fkTipoVia = 0, "", ",fkTipoVia") & IIf(escalera = "", "", ",escalera") & _
                IIf(puerta = "", "", ",puerta") & ")"
                
    strValues = " VALUES ('" & nombre & "', '" & razonSocial & "','" & cif & "'," & fkIFOCUsuario & ",'" & Format(Date, "dd/MM/yyyy") & "'" & _
                IIf(fkActividadEmpresarial = 0, "", ", " & fkActividadEmpresarial & " ") & _
                IIf(fechaCreacion < 1, "", ", '" & Format(fechaCreacion, "yyyy/mm/dd") & "' ") & IIf(telefono = 0, "", ", " & telefono & " ") & _
                IIf(fax = 0, "", ", " & fax & " ") & IIf(email = "", "", ", '" & email & "' ") & IIf(web = "", "", ", '" & web & "' ") & IIf(fkPlantillaMedia = 0, "", ", " & fkPlantillaMedia & " ") & _
                IIf(esAmbitoMunicipal = 0, "", ", " & esAmbitoMunicipal & " ") & IIf(esAmbitoAutonomico = 0, "", ", " & esAmbitoAutonomico & " ") & _
                IIf(esAmbitoEstatal = 0, "", ", " & esAmbitoEstatal & " ") & IIf(esAmbitoInternacional = 0, "", ", " & esAmbitoInternacional & " ") & _
                IIf(numCentrosCalvia = 0, "", ", " & numCentrosCalvia & " ") & IIf(numCentrosPalma = 0, "", ", " & numCentrosPalma & " ") & _
                IIf(numCentrosMallorca = 0, "", ", " & numCentrosMallorca & " ") & IIf(numCentrosBaleares = 0, "", ", " & numCentrosBaleares & " ") & _
                IIf(numCentrosEspana = 0, "", ", " & numCentrosEspana & " ") & IIf(numCentrosInternacional = 0, "", ", " & numCentrosInternacional & " ") & _
                IIf(SS = "", "", ", '" & SS & "' ") & IIf(domicilio = "", "", ", '" & domicilio & "' ") & _
                IIf(numero = 0, "", ", " & numero & " ") & IIf(codigoPostal = 0, "", ", " & codigoPostal & " ") & _
                IIf(fkPoblacion = 0, "", ", " & fkPoblacion & " ") & IIf(fkMunicipio = 0, "", ", " & fkMunicipio & " ") & _
                IIf(fkProvincia = 0, "", ", " & fkProvincia & " ") & IIf(fkTipoCentroTrabajo = 0, "", ", " & fkTipoCentroTrabajo & " ") & _
                IIf(bis = "", "", ", '" & bis & "' ") & IIf(bloque = "", "", ", '" & bloque & "' ") & IIf(piso = "", "", ", '" & piso & "' ") & _
                IIf(fkTipoVia = 0, "", ", " & fkTipoVia & " ") & IIf(escalera = "", "", ", '" & escalera & "' ") & _
                IIf(puerta = "", "", ", '" & puerta & "' ") & ");"


    
    sql = strSql & strValues
    Debug.Print sql
    CurrentDb.Execute sql
    
    
SalirError:
    insDatosEmpresa = 0
    Exit Function
TratarError:
    insDatosEmpresa = -1
    Debug.Print Err.description
    Debug.Print Err.Number
 
    

End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  30/12/2009 - Actualización:
'   Name:   insDatosActividadEconomicaOrganizacion
'   Desc:   Function para insertar el cnae
'   Param:  fkOrganizacion,fkACnae2009
'   Retur:  insDatosActividadEconomicaOrganizacion(long)
'---------------------------------------------------------------------------
Public Function insDatosActividadEconomicaOrganizacion(fkOrganizacion As Integer, _
                                   fkCnae2009 As Integer) As Integer
    Dim strSql As String
On Error GoTo TratarError
    
    strSql = " INSERT INTO T_ActividadEconomicaOrganizacion (fkOrganizacion, fkCnae2009,principal)" & _
             " VALUES (" & fkOrganizacion & ", " & fkCnae2009 & ",-1);"
             
     
    CurrentDb.Execute strSql
    Debug.Print strSql
SalirError:
    insDatosActividadEconomicaOrganizacion = 0
    Exit Function
TratarError:
    insDatosActividadEconomicaOrganizacion = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  25/11/2009 - Actualización:
'   Name:   insDatosPersonaContacto
'   Desc:   Function para insertar datos persona de contacto de empresa
'   Param:  nombre As String,apellido1 As String,apellido2 As String,
'
'   Retur:  insDatosPersonaContacto(long)
'---------------------------------------------------------------------------
Public Function insDatosPersonaContacto(nombre As String, _
                                   apellido1 As String, _
                                   apellido2 As String, _
                                   fkSexo As String) As Integer
    Dim strSql As String
   
On Error GoTo TratarError
    
    strSql = " INSERT INTO t_persona (nombre, apellido1, apellido2,fkSexo)" & _
             " VALUES ('" & nombre & "', '" & apellido1 & "','" & apellido2 & fkSexo & "') &" _
        & vbCrLf & " SELECT @@IDENTITY"
    
     Debug.Print strSql
    'CurrentDb.Execute strSql
   
    
SalirError:
    insDatosPersonaContacto = 0
    Exit Function
TratarError:
    insDatosPersonaContacto = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  25/11/2009 - Actualización:
'   Desc:   Function para insertar datos persona de contacto de empresa
'   Param:  nombre As String,apellido1 As String,apellido2 As String,
'
'   Retur:  insDatosPersonales(long)
'---------------------------------------------------------------------------
Public Function insDatosPersonaContactoOrganizacion(fkOrganizacion As Integer, _
                                   fkPersona As Integer, _
                                   cargo As String, _
                                   departamento As String, _
                                   telefonoFijo As Long, _
                                   telefonoMovil As Long, _
                                   email As String) As Integer
    Dim strSql As String
On Error GoTo TratarError
    
    strSql = " INSERT INTO r_organizacionpersona (fkOrganizacion, fkPersona, cargo, departamento, " & _
             " telefonoFijo,telefonoMovil,email)" & _
             " VALUES ('" & fkOrganizacion & "', '" & fkPersona & "','" & cargo & "','" & departamento & _
             "'," & telefonoFijo & "," & telefonoMovil & "," & email & ");"
    
    
    'CurrentDb.Execute strSql
    Debug.Print strSql
SalirError:
    insDatosPersonaContactoOrganizacion = 0
    Exit Function
TratarError:
    insDatosPersonaContactoOrganizacion = -1
End Function

'   Autor:  Nelson A. Hernández P.
'   Fecha:  27/11/2009 - Actualización:
'   Name:   insDatosConocimientosOrganizacion
'   Desc:   Inserta los datos de conocimiento en la tabla t_informatica de la persona pasada en los paramentros
'   Param:  fkPersona,fkInformatica,fkNivel,fkIfocUsuario
'   Retur: 0 si todo es correcto
'---------------------------------------------------------------------------
Public Function insDatosConocimientosOrganizacion(fkOrganizacion As Integer, _
                                   fkIFOCUsuario As Integer, _
                                   fecha As Date, _
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
    
On Error GoTo TratarError
    
    strSql = " INSERT INTO t_conocimientoifoc (fkOrganizacion, fkIfocUsuario,fecha,folletoIfoc, folletoIfocDonde, " & _
            " derivadoOtroServicio, derivadoOtroServicioCual,prensa, internet,familiaresAmigos,radio,revistaCalviaAldia,otros,otrosDescripcion)" & _
             " VALUES (" & fkOrganizacion & ", " & fkIFOCUsuario & ", '" & fecha & "', " & folletoIfoc & ", '" & folletoIfocDonde & "', " & derivadoOtroServicio & ", '" & derivadoOtroServicioCual & "', " & prensa & ", " & internet & ", " & familiaresamigos & ", " & radio & ", " & revistacalviaaldia & ", " & otros & ", '" & otrosDescripcion & "');"
    
    Debug.Print strSql
    CurrentDb.Execute strSql
    
     Debug.Print strSql
    
 
SalirError:
    insDatosConocimientosOrganizacion = 0
    Exit Function
TratarError:
    insDatosConocimientosOrganizacion = -1
    Debug.Print Err.description
        
    
End Function

'   Autor:  Nelson A. Hernández P.
'   Fecha:  27/11/2009 - Actualización:
'   Name:   upDateTWebOrganizacion
'   Desc:   Actualiza la tabla T_WEBORGANIZACION y agrega el id dado en sifoc en wifoc
'   Param: idOrganizacionSIFOC,idOrganizacionWifoc
'   Retur: 0 si todo es correcto
'---------------------------------------------------------------------------
Public Function upDateTWebOrganizacion(idOrganizacionSIFOC As Integer, idOrganizacionWifoc As Integer) As Integer
    Dim strSql As String
    Dim id As Integer
    
On Error GoTo TratarError
    
    strSql = "UPDATE T_WEBORGANIZACION SET idOrganizacionSIFOC=" & idOrganizacionSIFOC & " WHERE ID=" & idOrganizacionWifoc & ""
    
    CurrentDb.Execute strSql
    
     Debug.Print strSql
    
SalirError:
    upDateTWebOrganizacion = 0
    Exit Function
TratarError:
    upDateTWebOrganizacion = -1
    Debug.Print Err.description
        
    
End Function


'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  25/11/2009 - Actualización:
'   Name:   isExisteCif
'   Desc:   Function para verificar si existe alguna empresa con ese cif
'   Param:  CIF (string)
'   Retur:  true si existe,
'---------------------------------------------------------------------------
Public Function isExisteCif(cif As String) As Boolean
    Dim idem As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    str = " SELECT cif" & _
          " FROM t_organizacion" & _
          " WHERE cif='" & cif & "';"
  
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
   

    idem = False
    If Not (rs.EOF) Then
        rs.MoveFirst
        
       idem = True
       
    Else
        idem = False
    End If

    rs.Close
    Set rs = Nothing
    
    isExisteCif = idem
End Function

Public Function prueba10() As Boolean
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim lngRecs As String
    
    Set cn = CurrentProject.Connection
    
    cn.Execute ("INSERT INTO t_persona (fkSexo,nombre, apellido1 ) " _
                & "VALUES (1,'Nelson','Hernandez')"), lngRecs
    Debug.Print lngRecs
    'rs.Open "SELECT @@identity AS id FROM t_persona", cn
    rs.Open "SELECT max(id) AS NewID FROM t_persona", cn
    'Debug.Print rs!id
End Function


'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernández P.
'   Fecha:  25/11/2009 - Actualización:
'   Name:   getIdOrganizacionConCif
'   Desc:   Obtenemos id de la persona con dni pasado por parametro
'   Param:  Cif (string)
'   Retur:  idOrganizacion(long) si existe, sino 0
'---------------------------------------------------------------------------
Public Function getIdOrganizacionConCif(cif As String) As Long
    
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim idPersona As Long
On Error GoTo TratarError
    str = " SELECT id, Cif" & _
          " FROM t_organizacion" & _
          " WHERE Cif='" & cif & "';"
  
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
       
    If Not (rs.EOF) Then 'existe cif organizacion en sifoc
        rs.MoveFirst
        getIdOrganizacionConCif = rs!id
    Else 'no existe organizacion
        getIdOrganizacionConCif = 0
    End If

    rs.Close
    Set rs = Nothing
SalirError:
    Exit Function
TratarError:
    getIdOrganizacionConCif = 0
End Function


