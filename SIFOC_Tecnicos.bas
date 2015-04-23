Attribute VB_Name = "SIFOC_Tecnicos"
Option Explicit
Option Compare Database

'--------------------------------------------------------------------------------------------
'   Name:   numTecnicosAsignadosAUsuario
'   Autor:  Asunción Huertas
'   Fecha:  25/02/2010  Actualización:
'   Desc:   Devuelve el numero de técnicos asignados a la persona u organizacion
'           que se pasa por parámetro
'   Return: Número de técnicos asignados a la persona o organizacion del ifoc
'--------------------------------------------------------------------------------------------
Public Function checkAndAssignTR(idPersona As Long, idTR As Long, fecha As Date) As Boolean
    Dim respuesta

    'Check TR USUARIO > Si la persona no tiene como TR el tecnico de la cita, se le asigna
    If Not isTRDePersona(idTR, idPersona) Then
        respuesta = MsgBox("La persona no tiene el TR de la cita. Si no eres el TR no se realizarán altas automáticas de servicios." & vbNewLine & _
                           "Se le va a asignar a " & ifocUsuarioName(idTR) & " ¿Es correcto?", vbYesNo, "Alert: SIFOC_Tecnicos")
        If respuesta = vbYes Then
            If asignaTecnicoAPersona(idPersona, idTR, fecha) = 0 Then
                MsgBox "La asignación de TR se ha realizado correctamente", vbOKOnly, "Alert: SIFOC_Tecnicos"
                checkAndAssignTR = True
            Else
                MsgBox "La asignación de TR no se ha podido realizar", vbOKOnly, "Alert: SIFOC_Tecnicos"
                checkAndAssignTR = False
            End If
        End If
    End If
End Function
'--------------------------------------------------------------------------------------------
'   Name:   numTecnicosAsignadosAUsuario
'   Autor:  Asunción Huertas
'   Fecha:  25/02/2010  Actualización:
'   Desc:   Devuelve el numero de técnicos asignados a la persona u organizacion
'           que se pasa por parámetro
'   Return: Número de técnicos asignados a la persona o organizacion del ifoc
'--------------------------------------------------------------------------------------------
Public Function numTecnicosAsignadosAUsuario(fechaActivo As Date, _
                                             Optional idPersona As Long = 0, _
                                             Optional idOrganizacion As Long = 0) As Integer
On Error GoTo Error
    
    Dim fecha As Date
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim num As Integer
    
    fecha = Format(fechaActivo, "mm/dd/yyyy hh:nn:ss")

    sql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
          " FROM " & IIf((idPersona <> 0), "r_personaifocusuario", "r_organizacionifocusuario") & _
          " WHERE " & IIf((idPersona <> 0), "fkPersona = " & idPersona, "fkOrganizacion = " & idOrganizacion) & _
          " AND ((fechaBaja >= #" & fecha & "#) Or (fechaBaja Is Null))"

Debug.Print sql

    Set rs = New ADODB.Recordset
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        num = rs.RecordCount
    Else
        num = 0
    End If
    
    rs.Close
    Set rs = Nothing
    numTecnicosAsignadosAUsuario = num

    Exit Function

Error:
    debugando "Error: " & Err.description
    numTecnicosAsignadosAUsuario = -1
End Function

'--------------------------------------------------------------------------------------------
'   Name:   tecnicoAsignadoAUsuario
'   Autor:  Asunción Huertas - Update: José Manuel Sánchez Báez
'   Fecha:  26/03/2010 Asunción Huertas Actualización: 15/04/2014 Jose M. Sánchez
'   Desc:   Devuelve el primer técnico asignado a la persona/empresa en el ámbito
'           que se pasa por parámetro
'   Return: idTecnico - del técnico
'           -1        - si no tiene asignado TR
'--------------------------------------------------------------------------------------------
Public Function tecnicoAsignadoAUsuario(fechaActivo As Date, _
                                        Optional idPersona As Long = 0, _
                                        Optional idOrganizacion As Long = 0) As Long
On Error GoTo Error
       
    Dim fecha As Date
    Dim sql As String
    Dim rs As ADODB.Recordset

    fecha = Format(fechaActivo, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
    
        sql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
              " FROM " & IIf((idPersona <> 0), "r_personaifocusuario", "r_organizacionifocusuario") & _
              " WHERE " & IIf((idPersona <> 0), "fkPersona = " & idPersona, "fkOrganizacion = " & idOrganizacion) & _
              " AND ((fechaBaja >= #" & fecha & "#) Or (fechaBaja Is Null));"
           
        'Debug.Print sql
        Set rs = New ADODB.Recordset
        rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            rs.MoveFirst
            tecnicoAsignadoAUsuario = rs!fkIFOCUsuario
        Else
            tecnicoAsignadoAUsuario = -1
        End If
        
        rs.Close
        Set rs = Nothing
    
        Exit Function
    End If
    
Error:
    debugando "Error: " & Err.description
    tecnicoAsignadoAUsuario = -1
End Function

'--------------------------------------------------------------------------------------------
'   Name:   asignaTecnicoAPersona
'   Autor:  Jose Manuel Sanchez
'   Fecha:  26/03/2010 Asunción Huertas Actualización: 15/04/2014 Jose M. Sanchez
'   Desc:   Se asigna un técnico de referencia a una persona en un ámbito del ifoc
'           Crea un registro en r_personaifocusuario
'   Return: 0 -> si es correcto
'           -1 -> si es erróneo
'--------------------------------------------------------------------------------------------
Public Function asignaTecnicoAPersona(idPersona As Long, _
                                      idTecnico As Long, _
                                      fechaAlta) As Integer
On Error GoTo Error
    
    Dim fecha As Date
    Dim nivelTec As Integer
    Dim sql As String

    fecha = Format(fechaAlta, "mm/dd/yyyy hh:nn:ss")
    
    'Sólo permitimos asignar TR a persona a técnicos(No aux, No resp, No gerente)
    nivelTec = DLookup("[fkIfocNivel]", "[t_ifocusuariohistorico]", "[fkIfocUsuario]=" & idTecnico & " AND (fechaInicio <= now()) AND ((fechaFin is null) OR (fechaFin > now()))")
    
    If nivelTec < 4 Then
        sql = " INSERT INTO r_personaifocusuario (fkPersona, fkIfocUsuario, fkIfocUsuarioAlta, fechaAlta) " & _
              " VALUES (" & idPersona & ", " & idTecnico & ", " & U_idIfocUsuarioActivo & ", #" & fecha & "#)"

'Debug.Print sql
        CurrentDb.Execute sql
        asignaTecnicoAPersona = 0
    Else
        MsgBox "Solo se puede asignar como TR al personal técnico", vbOKOnly, "Alert: Módulo de Técnicos"
        asignaTecnicoAPersona = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    asignaTecnicoAPersona = -1
End Function

'--------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Name:   bajaPersonaDeTecnico
'   Fecha:  26/03/2010 Asunción Huertas Actualización: 15/04/2014 Jose M. Sanchez
'   Actualización: 10/03/2010 Asunción Huertas
'   Desc:   Pone fecha de baja a la relación de persona/técnico de referencia
'           pasados por parámetro
'   Return: 0 -> si es correcto
'           -1 -> si es erróneo
'--------------------------------------------------------------------------------------------
Public Function bajaPersonaDeTecnico(idPersona As Long, _
                                     idTecnico As Long, _
                                     fechaBaja As Date) As Integer
On Error GoTo Error
    
    Dim fecha As Date
    Dim sql As String

    fecha = Format(fechaBaja, "mm/dd/yyyy hh:nn:ss")
    
    sql = " UPDATE r_personaifocusuario" & _
          " SET fechaBaja = #" & fecha & "#" & _
          ", fkIfocUsuarioBaja = " & U_idIfocUsuarioActivo & _
          " WHERE (((fkPersona)=" & idPersona & _
          ") AND (fkIfocUsuario=" & idTecnico & _
          ") AND ((fechaBaja >= #" & fecha & "#) OR (fechaBaja Is Null)));"
    
    'Debug.Print sql
    CurrentDb.Execute sql
    bajaPersonaDeTecnico = 0
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    bajaPersonaDeTecnico = -1
End Function

'---------------------------------------------------------------------------
'   Name:   reasignaPersonaDeTecnico
'   Autor:  Jose Manuel Sanchez
'   Fecha:  26/03/2010 Asunción Huertas Actualización: 15/04/2014 Jose M. Sanchez
'   Desc:   Reasigna la persona del tecnico anterior al técnico nuevo en el mismo ámbito
'           Guardando el histórico de TR
'   Param:  idPersona(long), identificador de persona
'           idTecAnterior(long), identificador de tecnico anterior
'           idTecNuevo(long), identificador de tecnico nuevo
'   Retur:   0 -> si es correcto
'           -1 -> si es erróneo
'---------------------------------------------------------------------------
Public Function reasignaPersonaDeTecnico(idPersona As Long, _
                                         idTecAnterior As Long, _
                                         idTecNuevo As Long, _
                                         fechaCambio As Date) As Integer
On Error GoTo Error
   
    'Damos de baja del tecnico actual
    If bajaPersonaDeTecnico(idPersona, idTecAnterior, fechaCambio) = 0 Then
        'Damos de alta con nuevo técnico
        If asignaTecnicoAPersona(idPersona, idTecNuevo, fechaCambio) = -1 Then  'si hay error en la asignación
            reasignaPersonaDeTecnico = -1
        Else
            reasignaPersonaDeTecnico = 0
        End If
    Else 'si hay error en la baja
        reasignaPersonaDeTecnico = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    reasignaPersonaDeTecnico = -1
End Function

'--------------------------------------------------------------------------------------------
'   Name:   asignaTecnicoAEmpresa
'   Autor:  Jose Manuel Sanchez
'   Fecha:  26/03/2010 Asunción Huertas Actualización: 15/04/2014 Jose M. Sanchez
'   Desc:   Se asigna un técnico de referencia a una empresa en un ámbito del ifoc
'           Crea un registro en r_organizacionifocusuario
'   Return: 0 -> si es correcto
'           -1 -> si es erróneo
'--------------------------------------------------------------------------------------------
Public Function asignaTecnicoAEmpresa(idEmpresa As Long, _
                                      idTecnico As Long, _
                                      fechaAlta) As Integer
On Error GoTo Error
    
    Dim fecha As Date
    Dim nivelTec As Integer
    Dim sql As String

    fecha = Format(fechaAlta, "mm/dd/yyyy hh:nn:ss")
           
    'Sólo permitimos asignar TR a persona a técnicos(No aux, No resp, No gerente)
    nivelTec = DLookup("[fkIfocNivel]", "[t_ifocusuariohistorico]", "[fkIfocUsuario]=" & idTecnico & " AND (fechaInicio <= now()) AND ((fechaFin is null) OR (fechaFin > now()))")
    
    If nivelTec < 4 Then
        sql = " INSERT INTO r_organizacionifocusuario (fkOrganizacion, fkIfocUsuario, fkIfocUsuarioAlta, fechaAlta) " & _
              " VALUES (" & idEmpresa & ", " & idTecnico & ", " & U_idIfocUsuarioActivo & ", #" & fecha & "#);"
        
        'Debug.Print sql
        CurrentDb.Execute sql
        asignaTecnicoAEmpresa = 0
    Else
        MsgBox "Solo se puede asignar como TR al personal técnico", vbOKOnly, "Alert: Módulo de Técnicos"
        asignaTecnicoAEmpresa = -1
    End If
     
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    asignaTecnicoAEmpresa = -1
End Function
    
'--------------------------------------------------------------------------------------------
'   Name:   BajaEmpresaDeTecnico
'   Autor:  Jose Manuel Sanchez
'   Fecha:  26/03/2010 Asunción Huertas Actualización: 15/04/2014 Jose M. Sanchez
'   Desc:   Pone fecha de baja a la relación de empresa/técnico de referencia
'           pasados por parámetro
'   Return: 0 -> si es correcto
'           -1 -> si es erróneo
'--------------------------------------------------------------------------------------------
Public Function bajaEmpresaDeTecnico(idEmpresa As Long, _
                                     idTecnico As Long, _
                                     fechaBaja As Date) As Integer
On Error GoTo Error
    
    Dim fecha As Date
    Dim sql As String

    fecha = Format(fechaBaja, "mm/dd/yyyy hh:nn:ss")
    
    sql = " UPDATE r_organizacionifocusuario" & _
          " SET fechaBaja = #" & fecha & "#" & _
          ", fkIfocUsuarioBaja = " & U_idIfocUsuarioActivo & _
          " WHERE ((fkOrganizacion=" & idEmpresa & _
          ") AND (fkIfocUsuario=" & idTecnico & _
          ") AND ((fechaBaja >= #" & fecha & "#) OR (fechaBaja Is Null)))"
    
    'Debug.Print sql
    CurrentDb.Execute sql
    bajaEmpresaDeTecnico = 0
    
    Exit Function
        
Error:
    debugando "Error: " & Err.description
    bajaEmpresaDeTecnico = -1
End Function

'--------------------------------------------------------------------------------------------
'   Name:   bajasTecnicosDeEmpresa
'   Autor:  Jose Manuel Sanchez
'   Fecha:  26/03/2010 Asunción Huertas Actualización: 15/04/2014 Jose M. Sanchez
'   Desc:   Pone fecha de baja a la relación de técnico de referencia con empresa
'           pasados por parámetro
'   Return: 0 -> si es correcto
'           -1 -> si es erróneo
'--------------------------------------------------------------------------------------------
Public Function bajasTecnicosDeEmpresa(idEmpresa As Long, _
                                       fechaBaja As Date) As Integer
On Error GoTo Error
    
    Dim fecha As Date
    Dim sql As String

    fecha = Format(fechaBaja, "mm/dd/yyyy hh:nn:ss")
    
    sql = " UPDATE r_organizacionifocusuario" & _
          " SET fechaBaja = #" & fecha & "#" & _
          ", fkIfocUsuarioBaja = " & U_idIfocUsuarioActivo & _
          " WHERE (((fkOrganizacion)=" & idEmpresa & ") AND (((fechaBaja) >= #" & fecha & "#) OR ((fechaBaja) Is Null)));"
    
    'Debug.Print sql
    CurrentDb.Execute sql
    bajasTecnicosDeEmpresa = 0
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    bajasTecnicosDeEmpresa = -1
End Function

'---------------------------------------------------------------------------
'   Name:   reasignaEmpresaDeTecnico
'   Autor:  Asunción Huertas
'   Fecha:  26/03/2010 Asunción Huertas Actualización: 15/04/2014 Jose M. Sanchez
'   Desc:   Reasigna la empresa del tecnico anterior al técnico nuevo en el mismo ámbito
'           Guardando el histórico de TR
'   Param:  idEmpresa(long), identificador de empresa
'           idTecAnterior(long), identificador de tecnico anterior
'           idTecNuevo(long), identificador de tecnico nuevo
'   Retur:   0 -> si es correcto
'           -1 -> si es erróneo
'---------------------------------------------------------------------------
Public Function reasignaEmpresaDeTecnico(idEmpresa As Long, _
                                         idTecAnterior As Long, _
                                         idTecNuevo As Long, _
                                         fechaCambio As Date) As Integer
On Error GoTo Error
   
    'Damos de baja del tecnico actual
    If bajaEmpresaDeTecnico(idEmpresa, idTecAnterior, fechaCambio) = 0 Then
        'Damos de alta con nuevo técnico
        If asignaTecnicoAEmpresa(idEmpresa, idTecNuevo, fechaCambio) = -1 Then  'si hay error en la asignación
            reasignaEmpresaDeTecnico = -1
        Else
            reasignaEmpresaDeTecnico = 0
        End If
    Else 'si hay error en la baja
        reasignaEmpresaDeTecnico = -1
    End If
    
    Exit Function
    
Error:
    debugando "Error: " & Err.description
    reasignaEmpresaDeTecnico = -1
End Function


'--------------------------------------------------------------------------------------------
'   Autor:  José Manuel Sánchez Báez
'   Fecha:  26/03/2010 Actualización: 07/05/2014 Jose M. Sanchez
'   Name:   isTRDePersona
'   Desc:   Indica si el técnico pasado por parámetro
'           es TR de la persona pasada por parametro
'   Param:  idIfocUsuario(long),identificador de organización (si no persona)
'           idPersona(long)     ,identificador de persona (si no organización)
'   Return: True - tr de persona
'           False -no tr de persona
'--------------------------------------------------------------------------------------------
Public Function isTRDePersona(idIfocUsuario As Long, idPersona As Long) As Boolean
    
    On Error GoTo TratarError
       
    Dim fecha As Date
    Dim sql As String
    Dim rs As ADODB.Recordset

    fecha = Format(now(), "mm/dd/yyyy hh:nn:ss")

        sql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
              " FROM r_personaifocusuario" & _
              " WHERE fkPersona = " & idPersona & " AND fkIfocUsuario = " & idIfocUsuario & _
              " AND (fechaAlta < #" & fecha & "#)" & _
              " AND ((fechaBaja >= #" & fecha & "#) Or (fechaBaja Is Null))"
    
        'Debug.Print sql
        Set rs = New ADODB.Recordset
        rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            rs.MoveFirst
            isTRDePersona = True
        Else
            isTRDePersona = False
        End If
    
    rs.Close
    Set rs = Nothing

SalirTratarError:
    Exit Function
TratarError:
    debugando "Error: " & Err.description
    isTRDePersona = False
End Function

'--------------------------------------------------------------------------------------------
'   Autor:  José Manuel Sánchez Báez
'   Fecha:  07/05/2014 Actualización: 07/05/2014 Jose M. Sanchez
'   Name:   isTRDeOrganizacion
'   Desc:   Indica si el técnico pasado por parámetro
'           es TR de la persona pasada por parametro
'   Param:  idIfocUsuario(long),identificador de organización (si no persona)
'           idPersona(long)     ,identificador de persona (si no organización)
'   Return: True - tr de organizacion
'           False -no tr de organizacion
'--------------------------------------------------------------------------------------------
Public Function isTRDeOrganizacion(idIfocUsuario As Long, idOrganizacion As Long) As Boolean
    
    On Error GoTo TratarError
       
    Dim fecha As Date
    Dim sql As String
    Dim rs As ADODB.Recordset

    fecha = Format(now(), "mm/dd/yyyy hh:nn:ss")

        sql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
              " FROM r_organizacionifocusuario" & _
              " WHERE fkOrganizacion = " & idOrganizacion & _
              " AND (fkIfocUsuario = " & idIfocUsuario & ")" & _
              " AND (fechaAlta <= #" & fecha & "#)" & _
              " AND ((fechaBaja >= #" & fecha & "#) Or (fechaBaja Is Null))"

'Debug.Print sql
        Set rs = New ADODB.Recordset
        rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            rs.MoveFirst
            isTRDeOrganizacion = True
        Else
            isTRDeOrganizacion = False
        End If
    
    rs.Close
    Set rs = Nothing

SalirTratarError:
    Exit Function
TratarError:
    debugando "Error: " & Err.description
    isTRDeOrganizacion = False
End Function

'--------------------------------------------------------------------------------------------
'   Autor:  José Manuel Sánchez Báez
'   Fecha:  26/03/2010 Actualización: 15/04/2014 Jose M. Sanchez
'   Name:   trDePersona
'   Desc:   Devuelve técnico(primero en caso de error) asignado a la persona/empresa en el ámbito
'           que se pasa por parámetro
'   Param:  *fechaActivo(date)
'           idPersona(long)     ,identificador de persona (si no organización)
'   Return: idTecnico - del técnico
'           -1        - si no tiene asignado TR
'--------------------------------------------------------------------------------------------
Public Function trDePersona(fechaActivo As Date, _
                            idPersona As Long) As Long
On Error GoTo TratarError
       
    Dim fecha As Date
    Dim sql As String
    Dim rs As ADODB.Recordset

    fecha = Format(fechaActivo, "mm/dd/yyyy hh:nn:ss")

        sql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
              " FROM r_personaifocusuario" & _
              " WHERE fkPersona = " & idPersona & _
              " AND ((fechaBaja >= #" & fecha & "#) Or (fechaBaja Is Null));"
    
        'Debug.Print sql
        Set rs = New ADODB.Recordset
        rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
        If Not rs.EOF Then
            rs.MoveFirst
            trDePersona = rs!fkIFOCUsuario
        Else
            trDePersona = -1
        End If
    
    rs.Close
    Set rs = Nothing

SalirTratarError:
    Exit Function
TratarError:
    debugando "Error: " & Err.description
    trDePersona = -1
End Function

'--------------------------------------------------------------------------------------------
'   Autor:  José Manuel Sánchez Báez - Update: José Manuel Sánchez Báez
'   Fecha:  14/04/2014 - Act: 14/04/2014
'   Name:   trDeOrganizacion
'   Desc:   Devuelve técnico(primero en caso de error) asignado a la persona/empresa en el ámbito
'           que se pasa por parámetro
'   Param:  *fechaActivo(date)
'           idOrganizacion(long),identificador de organización (si no persona)
'   Return: idTecnico - del técnico
'           -1        - si no tiene asignado TR
'--------------------------------------------------------------------------------------------
Public Function trDeOrganizacion(fechaActivo As Date, _
                                 idOrganizacion As Long) As Long
On Error GoTo TratarError
    
    Dim fecha As Date
    Dim sql As String
    Dim rs As ADODB.Recordset

    fecha = Format(fechaActivo, "mm/dd/yyyy hh:nn:ss")

    sql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
          " FROM r_organizacionifocusuario" & _
          " WHERE fkOrganizacion = " & idOrganizacion & _
          " AND ((fechaBaja >= #" & fecha & "#) Or (fechaBaja Is Null));"

    'Debug.Print sql
    Set rs = New ADODB.Recordset
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        trDeOrganizacion = rs!fkIFOCUsuario
    Else
        trDeOrganizacion = -1
    End If
    
    rs.Close
    Set rs = Nothing

SalirTratarError:
    Exit Function
TratarError:
    debugando "Error: " & Err.description
    trDeOrganizacion = -1
End Function

'--------------------------------------------------------------------------------------------
'   Autor:  José Manuel Sánchez Báez - Update: José Manuel Sánchez Báez
'   Fecha:  26/05/2010 - Act: 26/05/2010
'   Name:   trDeUsuario
'   Desc:   Devuelve técnico(primero en caso de error) asignado a la persona/empresa en el ámbito
'           que se pasa por parámetro
'   Param:  *idAmbito(integer)
'           *fechaActivo(date)
'           idPersona(long)     ,identificador de persona (si no organización)
'           idOrganizacion(long),identificador de organización (si no persona)
'   Return: idTecnico - del técnico
'           -1        - si no tiene asignado TR
'--------------------------------------------------------------------------------------------
'Public Function trDeUsuario(idAmbito As Integer, _
'                            fechaActivo As Date, _
'                            Optional idPersona As Long = 0, _
'                            Optional idOrganizacion As Long = 0) As Long
'On Error GoTo TratarError
'
'    Dim fecha As Date
'    Dim sql As String
'    Dim rs As ADODB.Recordset
'
'    fecha = Format(fechaActivo, "mm/dd/yyyy hh:nn:ss")
'
'    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
'        sql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
'              " FROM " & IIf((idPersona <> 0), "r_personaifocusuario", "r_organizacionifocusuario") & _
'              " WHERE " & IIf((idPersona <> 0), "fkPersona = " & idPersona, "fkOrganizacion = " & idOrganizacion) & _
'              " AND fkIfocAmbito = " & idAmbito & _
'              " AND ((fechaBaja >= #" & fecha & "#) Or (fechaBaja Is Null));"
'
'        'Debug.Print sql
'        Set rs = New ADODB.Recordset
'        rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
'
'        If Not rs.EOF Then
'            rs.MoveFirst
'            trDeUsuario = rs!fkIFOCUsuario
'        Else
'            trDeUsuario = -1
'        End If
'    End If
'
'    rs.Close
'    Set rs = Nothing
'
'SalirTratarError:
'    Exit Function
'TratarError:
'    debugando "Error: " & Err.description
'    trDeUsuario = -1
'End Function

'--------------------------------------------------------------------------------------------
'   Name:   esTRdePersona
'   Autor:  José Manuel Sánchez
'   Update: José Manuel Sánchez
'   Fecha:  14/04/2014 Actualización: 14/04/2014
'   Desc:   Comprueba si la persona tiene asignado ese TR
'   Param:  idPersona(long)
'           idIfocUsuario(long)
'   Return: 0 -> si es correcto
'           -1 -> si es erróneo
'--------------------------------------------------------------------------------------------
'Public Function esTRdePersona(idPersona As Long, _
'                              idIfocUsuario As Long) As Boolean
'On Error GoTo TratarError
'
'    Dim resultado As Boolean
'    Dim idTecnico As Long
'
'    idTecnico = trDePersona(Now(), idPersona)
'
'    If (idIfocUsuario = idTecnico) Then
'        resultado = True
'    Else
'        resultado = False
'    End If
'
'    esTRdePersona = resultado
'SalirTratarError:
'    Exit Function
'TratarError:
'    debugando "Error: " & Err.description
'    esTRdePersona = False
'End Function

'---------------------------------------------------------------------------
'   Name:   TRActivoUsuario
'   Autor:  José Manuel Sanchez - Upd: José Manuel Sanchez
'   Fecha:  15/04/2014
'   Desc:   Indica si el usuario está de alta
'           en el servicio indicado en la fecha indicada
'           Si no se le pasa un servicio, indica si el usuario está de alta
'           en algún servicio en la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  fecha(date), fecha de actividad
'           idServicio(long), identificador de servicio (OPCIONAL)
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:   1, usuario con TR asignado en fecha
'            0, usuario sin TR asignado en fecha
'           -1, ko ERROR
'---------------------------------------------------------------------------
Public Function TRActivoUsuario(fecha As Date, _
                                Optional idPersona As Long = 0, _
                                Optional idOrganizacion As Long = 0) As Integer
On Error GoTo TratarError
    
    Dim fechaQuery As Date
    Dim strSql As String
    Dim resultado As Integer
    Dim rs As ADODB.Recordset
    
    fechaQuery = Format(fecha, "mm/dd/yyyy hh:nn:ss")

    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
                 " FROM " & IIf((idPersona <> 0), "r_personaifocusuario", "r_organizacionifocusuario") & _
                 " WHERE " & IIf((idPersona <> 0), "fkPersona = " & idPersona, "fkOrganizacion = " & idOrganizacion) & _
                 " AND (fechaAlta<=#" & fechaQuery & "#) AND ((fechaBaja >= #" & fechaQuery & "#) Or (fechaBaja Is Null));"
    
'Debug.Print strSql

        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        If (rs.RecordCount > 0) Then
            resultado = 1
        Else
            resultado = 0
        End If
            
        rs.Close
        Set rs = Nothing
        
    Else
        resultado = -1
    End If
    
    TRActivoUsuario = resultado
    
    Exit Function
    
TratarError:
    debugando "Error: " & Err.description
    TRActivoUsuario = -1
End Function

'----------------------------------------------------------------------------------------------
'   Name:   TRPosteriorUsuario
'   Autor:  Jose Manuel Sanchez - Upd: Jose Manuel Sanchez
'   Fecha:  08/06/2010 - Update: 15/04/2014
'   Desc:   Indica si el usuario esta de alta en el servicio posteriormente a la fecha indicada
'           Si no se le pasa un servicio, indica si el usuario estará de alta en algún servicio
'           posteriormente a la fecha indicada
'           Sólo se le pasa idPersona o idOrganizacion a la vez
'   Param:  *fecha(date), fecha de actividad
'           idPersona(long), identificador de persona (OPCIONAL)
'           idOrganizacion(long), identificador de organizacion (OPCIONAL)
'   Retur:  Número de servicios activos del usuario con fecha posterior a la fecha indicada
'           (servicios futuros)
'           -1, ko
'----------------------------------------------------------------------------------------------
Public Function TRPosteriorUsuario(fecha As Date, _
                                   Optional idPersona As Long = 0, _
                                   Optional idOrganizacion As Long = 0) As Integer
On Error GoTo TratarError
    
    Dim fechaQuery As Date
    Dim strSql As String
    Dim resultado As Integer
    Dim rs As ADODB.Recordset
    
    fechaQuery = Format(fecha, "mm/dd/yyyy hh:nn:ss")
    
    If (idPersona <> 0 And idOrganizacion = 0) Or (idPersona = 0 And idOrganizacion <> 0) Then
        strSql = " SELECT fkIfocUsuario, fechaAlta, fechaBaja" & _
                 " FROM " & IIf((idPersona <> 0), "r_personaifocusuario", "r_organizacionifocusuario") & _
                 " WHERE " & IIf((idPersona <> 0), "fkPersona = " & idPersona, "fkOrganizacion = " & idOrganizacion) & _
                 " AND ((fechaAlta> #" & fechaQuery & "#)" & _
                 " AND ((fechaBaja > #" & fechaQuery & "#) OR (fechaBaja Is Null)));"
        
'Debug.Print strSql

        Set rs = New ADODB.Recordset
        rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
        
        If (rs.RecordCount > 0) Then
            resultado = 1
        Else
            resultado = 0
        End If
        
        rs.Close
        Set rs = Nothing
        
    Else
        resultado = -1
    End If
    
    TRPosteriorUsuario = resultado
    Exit Function
    
TratarError:
    debugando "Error: " & Err.description
    TRPosteriorUsuario = -1
End Function

