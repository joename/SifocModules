Attribute VB_Name = "SIFOC_IfocUsuario"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  30/11/2011
'   Name:   insIfocUsuario
'   Desc:   Inserta periodo de alta de IFOC usuario, no alta histórico
'           la persona debe estar creada previamente
'   Param:  idIfocUsuario
'           aka
'           user
'           password
'           fechaPassword
'   Retur:  string, nivel del usuario ifoc
'---------------------------------------------------------------------------
Public Function insIfocUsuario(idIfocUsuario As Long, _
                               aka As String, _
                               user As String, _
                               password As String) As Integer
On Error GoTo TratarError

    Dim strSql As String
    Dim strFields As String
    Dim strValues As String
    
    strFields = "(fkPersona" & _
                ", aka" & _
                ", username" & _
                ", password" & _
                ", fechaPassword" & _
                ")"
    
    strValues = "(" & idIfocUsuario & _
                ", """ & filterSQL(aka) & """" & _
                ", """ & filterSQL(user) & """" & _
                ", """ & filterSQL(password) & """" & _
                ", now()" & _
                ")"
    
    strSql = " INSERT INTO t_ifocusuario" & _
             strFields & _
             " VALUES " & _
             strValues & ";"
    
    CurrentDb.Execute strSql
    insIfocUsuario = 0
    
Debug.Print strSql

SalirTratarError:
    Exit Function
TratarError:
    Debug.Print "Error (Inserción Ifoc Usuario): " & vbNewLine & _
            Err.description & "Alert: SIFOC IfocUsuario"
    insIfocUsuario = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  30/11/2011
'   Name:   insIfocUsuarioHistorico
'   Desc:   Inserta periodo de alta de IFOC usuario histórico,
'           implica que Ifoc usuario tiene que estar de alta
'   Param:
'   Retur:  string, nivel del usuario ifoc
'---------------------------------------------------------------------------
Public Function insIfocUsuarioHistorico(idIfocUsuario As Long, _
                                        idUnidad As Integer, _
                                        idIfocNivel As Integer, _
                                        fechaInicio As Date, _
                                        Optional fechaFin As Date = "01/01/1900", _
                                        Optional idIfocSubArea As Integer = 0) As Integer
    Dim strSql As String
    Dim strFields As String
    Dim strValues As String
    
    
    
    strFields = "(fkIfocUsuario" & _
                ", fkIfocUnidad" & _
                ", fkIfocNivel" & _
                ", fechaInicio" & _
                IIf(fechaFin = "01/01/1900", "", ", fechaFin") & _
                IIf(idIfocSubArea = 0, "", ", fkIfocSubArea") & _
                ", updDate" & _
                ", fkIfocUsuarioUpd" & _
                ")"
    
    strValues = "(" & idIfocUsuario & _
                ", " & idUnidad & _
                ", " & idIfocNivel & _
                ", #" & Format(fechaInicio, "mm/dd/yyyy hh:nn:ss") & "#" & _
                IIf(fechaFin = "01/01/1900", "", ", #" & Format(fechaFin, "mm/dd/yyyy hh:nn:ss") & "#") & _
                IIf(idIfocSubArea = 0, "", ", " & idIfocSubArea) & _
                ", now()" & _
                ", " & U_idIfocUsuarioActivo & _
                ")"
    
    strSql = " INSERT INTO t_ifocusuariohistorico" & _
             strFields & _
             " VALUES " & _
             strValues & ";"
    
'Debug.Print strSql
    CurrentDb.Execute strSql
    insIfocUsuarioHistorico = 0

SalirTratarError:
    Exit Function
TratarError:
    Debug.Print "Error (Inserción Ifoc Usuario Historico): " & vbNewLine & _
            Err.description & "Alert: SIFOC IfocUsuario"
    insIfocUsuarioHistorico = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  30/11/2011
'   Name:   insIfocUsuarioHistorico
'   Desc:   Inserta periodo de alta de IFOC usuario histórico,
'           implica que Ifoc usuario tiene que estar de alta
'   Param:
'   Retur:  string, nivel del usuario ifoc
'---------------------------------------------------------------------------
Public Function updIfocUsuarioHistorico(idIUHistorico As Long, _
                                        idIfocUsuario As Long, _
                                        Optional idUnidad As Integer = 0, _
                                        Optional idIfocNivel As Integer = 0, _
                                        Optional fechaInicio As Date = "01/01/1900", _
                                        Optional fechaFin As Date = "01/01/1900", _
                                        Optional idIfocSubArea As Integer = 0) As Integer
    Dim strSql As String
    Dim strFields As String
    Dim strValues As String
    
    strFields = "fkIfocUsuario = " & idIfocUsuario & _
                IIf(idUnidad = 0, "", ", fkIfocUnidad = " & idUnidad) & _
                IIf(idIfocNivel = 0, "", ", fkIfocNivel = " & idIfocNivel) & _
                IIf(idIfocSubArea = 0, ", fkIfocSubArea = null", ", fkIfocSubArea = " & idIfocSubArea) & _
                IIf(fechaInicio = "01/01/1900", "", ", fechaInicio = #" & Format(fechaInicio, "mm/dd/yyyy hh:mm:nn") & "#") & _
                IIf(fechaFin = "01/01/1900", ", fechafin = null", ", fechaFin = #" & Format(fechaFin, "mm/dd/yyyy hh:mm:nn") & "#") & _
                ", updDate = now()" & _
                ", fkIfocUsuarioUpd = " & usuarioIFOC()
    
    strSql = " UPDATE t_ifocusuariohistorico" & _
             " SET " & strFields & _
             " WHERE id = " & idIUHistorico & ";"
    
'Debug.Print strSql
    CurrentDb.Execute strSql
    updIfocUsuarioHistorico = 0
SalirTratarError:
    Exit Function
TratarError:
    Debug.Print "Error (Update Ifoc Usuario Historico): " & vbNewLine & _
            Err.description & "Alert: SIFOC IfocUsuario"
    updIfocUsuarioHistorico = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  30/11/2011
'   Name:   insIfocUsuarioHistorico
'   Desc:   Inserta periodo de alta de IFOC usuario histórico,
'           implica que Ifoc usuario tiene que estar de alta
'   Param:
'   Retur:  string, nivel del usuario ifoc
'---------------------------------------------------------------------------
Public Function isIfocUsuarioActivo(idIfocUsuario As Long, _
                                    Optional fechaInicio As Date = "01/01/1900", _
                                    Optional fechaFin As Date = "01/01/1900") As Boolean
    Dim strSql As String
    Dim FECHAF As Date
    Dim activo As Boolean
    
    activo = False
    
    FECHAF = IIf(fechaFin = "01/01/1900", now(), fechaFin)
    
    If (fechaInicio < FECHAF) Then
        strSql = " SELECT fkIfocUsuario" & _
                 " FROM t_ifocUsuarioHistorico" & _
                 " WHERE (fechaInicio < #" & Format(FECHAF, "mm/dd/yyyy hh:nn:ss") & "# AND (fechaFin > #" & Format(fechaInicio, "mm/dd/yyyy hh:nn:ss") & "# OR fechaFin is null))" & _
                 " AND (fkIfocUsuario = " & idIfocUsuario & ")"
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
'Debug.Print strSql
        rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        If Not rs.EOF Then
            activo = True
        End If
    Else
        activo = False
    End If
    
    isIfocUsuarioActivo = activo
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  23/01/2014
'   Name:   ifocUsuarioName
'   Desc:   Nos devuelve nombro y apellidos del ifoc usuario
'   Param:
'   Retur:  string, nombre completo de usuario ifoc
'---------------------------------------------------------------------------
Public Function ifocUsuarioName(idIfocUsuario As Long) As String
    ifocUsuarioName = Nz(DLookup("[name]", "[v_ifocusuario]", "[fkPersona]=" & Nz(idIfocUsuario, 0)), "")
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  18/11/2013
'   Name:   ifocUsuarioActivoNow
'   Desc:   devuelve si el usuario ifoc está activo en sifoc
'   Param:
'   Return:  True, si usuario está activo como usuario ifoc
'            False, otherwise
'---------------------------------------------------------------------------
Public Function isIfocUsuarioActivoNow(idIfocUsuario As Long)
    Dim strSql As String
    Dim FECHAF As Date
    Dim activo As Boolean
    
    activo = False
    
    strSql = " SELECT idIfocUsuario" & _
             " FROM v_ifocusuario_activo" & _
             " WHERE idIfocUsuario = " & idIfocUsuario
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

'Debug.Print strSql
        
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
    If Not rs.EOF Then
        activo = True
    Else
        activo = False
    End If
    
    isIfocUsuarioActivoNow = activo
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/04/2009
'   Name:   ifocUsuarioNivel
'   Desc:   devuelve el nivel del usuario ifoc que se pasa por parámetro
'   Param:
'   Retur:  string, nivel del usuario ifoc
'---------------------------------------------------------------------------
Public Function ifocUsuarioIdNivel(idIfocUsuario As Long) As Integer
    ifocUsuarioIdNivel = Nz(DLookup("idIfocNivel", "v_ifocusuario_activo", "[idIfocUsuario]=" & idIfocUsuario), 0)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  01/6/2007   Actualización: 18/03/2010
'   Name:   usuarioIFOC
'   Desc:   devuelve el usuario ifoc que se logeo en la base
'   Param:  null
'   Retur:  Long, id del usuario ifoc
'---------------------------------------------------------------------------
Public Function usuarioIFOC() As Long
    Dim usr As Long
    
    If IsFormLoaded("Inicio") Then 'And Not IsNull(Nz(Forms!Inicio!cbx_usuario, Null)) Then
        Dim idIfocUsuario As Long
        Dim fechaSesion As Date
        
        idIfocUsuario = U_idIfocUsuarioActivo
        
        fechaSesion = Nz(DLookup("[logInSession]", "t_ifocusuario", "[fkPersona]=" & idIfocUsuario), now() - 1)
        
        If (fechaSesion < now) And (fechaSesion > Date) Then
            usr = idIfocUsuario
        Else
            usr = 0
        End If
    Else
        usr = 0
    End If
    usuarioIFOC = usr
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/04/2009
'   Name:   ifocUsuarioUnidad
'   Desc:   devuelve el usuario ifoc que se logeo en la base
'   Param:  null
'   Retur:  string, unidad del usuario ifoc
'---------------------------------------------------------------------------
Public Function ifocUsuarioUnidad(idIfocUsuario As Long) As String
    ifocUsuarioUnidad = Nz(DLookup("unidad", "v_ifocusuario_activo", "[idIfocUsuario]=" & idIfocUsuario), 0)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/04/2009
'   Name:   usuarioIfocSubArea
'   Desc:   devuelve el usuario ifoc que se logeo en la base
'   Param:  null
'   Retur:  string, subArea del usuario ifoc
'---------------------------------------------------------------------------
Public Function ifocUsuarioSubArea(idIfocUsuario As Long) As String
    ifocUsuarioSubArea = Nz(DLookup("subarea", "v_ifocusuario_activo", "[idIfocUsuario]=" & idIfocUsuario), 0)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/04/2009
'   Name:   usuarioIfocNivel
'   Desc:   devuelve el nivel del usuario ifoc que se pasa por parámetro
'   Param:  null
'   Retur:  string, nivel del usuario ifoc
'---------------------------------------------------------------------------
Public Function ifocUsuarioNivel(idIfocUsuario As Long) As String
    ifocUsuarioNivel = Nz(DLookup("nivel", "v_ifocusuario_activo", "[idIfocUsuario]=" & idIfocUsuario), 0)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  16/04/2009
'   Name:   actualizaDatosUsuario
'   Desc:   Actualiza nombre de usuario y ordenador en el que inicia sesión
'   Param:  idPersona(long), identificador ifoc usuario(persona)
'   Retur:  -
'---------------------------------------------------------------------------
Function actualizaDatosUsuario(idPersona As Long) As Integer
    Dim str As String
    
    str = " UPDATE t_ifocusuario" & _
          " SET t_ifocusuario.computerName ='" & computerName() & "'," & _
          " t_ifocusuario.username ='" & userName() & "'," & _
          " t_ifocusuario.logInSession ='" & now() & "'," & _
          " t_ifocusuario.logOutSession = null" & _
          " WHERE (((t_ifocusuario.fkPersona)=" & idPersona & "));"

    CurrentDb.Execute str
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  30/11/2011
'   Name:   updIfocUsuarioPassword
'   Desc:   Actualizamos el password del usuario pasado por parámetro
'   Param:
'   Retur:  integer, -1 ko
'                     0 ok
'---------------------------------------------------------------------------
Public Function updIfocUsuarioPassword(idIfocUsuario As Long, password As String) As Integer
    Dim strSql As String
    Dim pass As String
    Dim fecha As String
On Error GoTo TratarError
    
    fecha = now() + 365

    strSql = " UPDATE t_ifocusuario SET password = '" & password & "', fechaPassword = now()+365" & _
             " WHERE fkPersona = " & idIfocUsuario
    CurrentDb.Execute strSql
    updIfocUsuarioPassword = 0
    
SalirTratarError:
    Exit Function
TratarError:
    Debug.Print "Error: " & Err.description, , "Alert: Sifoc_ifocUsuario"
    updIfocUsuarioPassword = -1
End Function
