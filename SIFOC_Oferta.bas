Attribute VB_Name = "SIFOC_Oferta"
Option Explicit
Option Compare Database

Private strSql As String
Private strselect As String
Private strFrom As String
Private strWhere As String
Private strGroup As String
Private strHaving As String
Private strOrder As String

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  09/07/2009 - Actualizacion: 15/03/2010
'   Name:   esOfertaModificable
'   Desc:   Nos indica si una oferta es modificable o no
'           Oferta modificable si esta abierta o si es el técnico de la oferta
'   Param:  id de la oferta que queremos consultar si es modificable
'   Return: TRUE, si se puede modificar la oferta
'           FALSE, si no se puede modificar la oferta
'           (tenemos en cuenta si esta cerrada o no)
'---------------------------------------------------------------------------
Public Function esOfertaModificable(idOferta As Long, _
                                    idIfocUsuario As Long) As Boolean
    Dim estado As Integer
    Dim idUsuarioIfocOferta As Long
    Dim fechaestado As Date
On Error GoTo TratarError:

    If (idOferta = 0) Then
        esOfertaModificable = True
        Exit Function
    End If
    
    'Obtenemos el último estado
    estado = getidEstadoActualOferta(idOferta)
    
    'Obtenemos la fecha del último estado de la oferta
    fechaestado = Format(Nz(DLookup("max(fecha)", _
                            "t_ofertaestados", _
                            "[fkoferta] = " & idOferta), "00/00/0000"), "dd/mm/yyyy")
    
    'En caso de estar null fkUsuarioIfocAna se coje 0, para que los controles estén activos
    idUsuarioIfocOferta = Nz(DLookup("[fkUsuarioIfocAna]", "t_oferta", "[id]=" & idOferta), 0)
    If (estado <> 0) Then
        'Comprobamos que se pueda modificar en cualquier estado que no sea "Cerrada" o "Modificada"
        'con día diferente al de hoy
        If (estado < 5) Or ((estado = 6) And (fechaestado = Format(now, "dd/mm/yyyy"))) Then
            esOfertaModificable = True
        ElseIf (idUsuarioIfocOferta = idIfocUsuario) Then
            esOfertaModificable = True
        Else
            esOfertaModificable = False
        End If
    Else
        esOfertaModificable = False
    End If

SalirTratarError:
    Exit Function
TratarError:
    esOfertaModificable = False
End Function

'-----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  1/7/2007 -   Actualización:  15/03/2010
'   Descr:  Registar estado de la oferta y guarda estado, fecha en que se realiza el cambio y el usuarioIFOC
'   Param:  idOferta
'           fkEstado
'           fkIfocUsr
'   Retur:
'-----------------------------------------------------------------------------------------------------------
Public Function altaCambioEstadoOferta(idOferta As Long, _
                                       idEstado As Integer, _
                                       idIfocUsr As Long) As Integer
    Dim sql As String
    Dim estadoactual As Integer
    Dim motivoBD As String
    Dim motivo As String
    
    
    'Controlamos el caso especial del estado "Modificada" (Reabrir Oferta)
    If (idEstado = 6) Then
        'Pedimos al usuario que introduzca un motivo para reabrir la oferta
        motivo = InputBox("Introduzca el motivo por el que se desea reabrir la oferta cerrada. Mínimo 10 carácteres", _
                "Reapertura de la oferta en estado de cierre", "")
                
        'En el caso que el motivo introducido sea correcto y el estado de la oferta ya sea "Modificada",
        ' sólo realizaremos una actualización del estado "Modificada" agregando: el nuevo motivo, fecha, usuarioIFOC
        If (Len(motivo) > 10) And (getidEstadoActualOferta(idOferta) = 6) Then
                motivoBD = Nz(DLookup("[motivoApertura]", _
                                      "[t_ofertaestados]", _
                                      "[fkoferta] = " & idOferta & " And [fkofertaestado] = 6"), "")
                motivo = motivoBD + " --- " + Format(now, "dd/mm/yyyy") + _
                         " [" + DLookup("[aka]", "[t_ifocusuario]", "[fkpersona] = " & idIfocUsr) + "]: " + motivo
                sql = " UPDATE t_ofertaestados" & _
                      " SET fecha = Now(), fkusuarioIfoc = " & idIfocUsr & ",MotivoApertura = """ & motivo & """" & _
                      " WHERE (((t_ofertaestados.fkOferta) = " & idOferta & ") AND (t_ofertaestados.fkOfertaEstado = 6));"
                CurrentDb.Execute sql
                
        'En el caso que el motivo introducido sea correcto y el estado de la oferta no sea "Modificada",
        ' insertaremos en la tabla t_ofertaestados el nuevo estado de la oferta
        ElseIf (Len(motivo) > 10) And (getidEstadoActualOferta(idOferta) <> 6) Then
                motivo = Format(now, "dd/mm/yyyy") + " [" + DLookup("[aka]", _
                                                                    "[t_ifocusuario]", _
                                                                    "[fkpersona] = " & idIfocUsr) + "]: " + motivo
                insOfertaEstados idOferta, idEstado, idIfocUsr, motivo
                updOferta idOferta, idEstado
                
        'En caso de un motivo incorrecto mostraremos el siguiente mensaje
        Else
            MsgBox "No has introducido un motivo correcto. Mínimo 10 carácteres", vbOKOnly, "Alert: No se reabrió la oferta"
        End If
        
    'Trataremos los otros casos posibles
    Else
        'Obtenemos el estado actual de la oferta
        estadoactual = getidEstadoActualOferta(idOferta)
        
        'Miramos si la oferta esta cerrada
        If (estadoactual <> 5) And (estadoactual <> 6) Then
            
            insOfertaEstados idOferta, idEstado, idIfocUsr
            updOferta idOferta, idEstado
            
            If (idEstado = 5) Then
                sql = " UPDATE t_oferta SET fechacierre = Now()" & _
                      " WHERE ((t_oferta.id) = " & idOferta & ");"
                CurrentDb.Execute sql
            End If
        Else
            MsgBox "No se cambió el estado." & vbNewLine & _
                   "Una vez que la oferta está cerrada no se puede volver cambiar de estado." _
                   , vbOKOnly, "Alert: SIFOC_Empleo"
        End If
    End If
    
debugando sql

End Function

'-----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  1/7/2007 -   Actualización:  15/03/2010
'   Name:   insEstadoOfertaCandidato
'   Descr:  Registar estado de la oferta y guarda estado, fecha en que se realiza el cambio y el usuarioIFOC
'   Param:  idOferta
'           fkEstado
'           fkIfocUsr
'   Retur:
'-----------------------------------------------------------------------------------------------------------
Public Function insOfertaEstados(idOferta As Long, _
                                idEstado As Integer, _
                                idIfocUsr As Long, _
                                Optional motivoApertura As String = "") As Integer
    Dim sql As String
    
    sql = " INSERT INTO t_ofertaestados (fkOferta, fkOfertaEstado, fecha, fkUsuarioIFOC" & IIf(motivoApertura = "", "", ", MotivoApertura") & ") " & _
          " VALUES (" & idOferta & ", " & idEstado & ", Now(), " & idIfocUsr & IIf(motivoApertura = "", "", ", '" & filterSQL(motivoApertura) & "'") & ");"
    CurrentDb.Execute sql
End Function

'-----------------------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  1/7/2007 -   Actualización:  15/03/2010
'   Name:   updEstadoOferta
'   Descr:  Actualización estado oferta (motivo apertura o modificacion)
'   Param:  idOferta
'           fkEstado
'           fkIfocUsr
'   Retur:
'-----------------------------------------------------------------------------------------------------------
Public Function updOferta(idOferta As Long, _
                          Optional idEstado As Integer = 0) As Integer
    Dim sql As String
    
    sql = " UPDATE t_oferta" & _
          " SET t_oferta.fkOfertaEstado = " & idEstado & _
          " WHERE (t_oferta.id = " & idOferta & ");"
          
    CurrentDb.Execute sql
End Function

'--------------------------------------------------------------------------------------------
'       Añade persona a listado de candidatos bruto
'--------------------------------------------------------------------------------------------
Public Function addPersonaListadoCandidatoBruto(idOferta As Integer, _
                                                idPersona As Long, _
                                                Optional idUsuarioIfoc As Long = 0)
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    'Insertamos nuevo candidato en lista candidatos brutos
    strSql = " T_ListadoCandidatosBrutos"
    Set rs = New ADODB.Recordset

    rs.Open strSql, CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
    
    rs.AddNew
    rs!fkOferta = idOferta
    rs!fkCnoFiltro = Null
    rs!fkPersona = idPersona
    rs!fkUsuarioIFOC = idUsuarioIfoc
    'rs!fkEstado dejamos este campo vacio
    rs.update
    
    rs.Close
    Set rs = Nothing
End Function

'--------------------------------------------------------------------------------------------
'       Añade persona a seleccionadas oferta
'--------------------------------------------------------------------------------------------
Public Function addPersonaSeleccionadaOferta(idOferta As Integer, _
                                             idPersona As Long, _
                                             idIfocUsuario As Long, _
                                             Optional idCnoFiltro As Long = 0) As Integer
    Dim strSql As String
    Dim rs As ADODB.Recordset

On Error GoTo TratarError

    'Insertamos nuevo candidato en tabla de seleccionados
    strSql = " INSERT INTO r_ofertacandidatos (fkOferta, fkPersona, fkIFOCUsuario, fecha" & IIf(idCnoFiltro = 0, "", ", fkCnoFiltro") & ") " & _
             " VALUES (" & idOferta & ", " & idPersona & ", " & idIfocUsuario & ",now() " & IIf(idCnoFiltro = 0, "", ", " & idCnoFiltro) & ");"
    CurrentDb.Execute strSql
    
    'rs.Close
    'Set rs = Nothing
    
SalirTratarError:
    Exit Function
TratarError:
    addPersonaSeleccionadaOferta = -1
    debugando "SIFOC_Oferta - addPersonaSeleccionadaOferta" & Err.description
    Resume SalirTratarError
End Function

'------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  09/07/2009 - Actualizacion: 15/03/2010
'   Name:   actualizaSeguimientoEstadoPersonaSMS
'   Desc:   Actualiza Estado Candidatos de Oferta, por aviso de sms
'   Param:  ID de la oferta, personas, estado y el detalle del estado
'   Return: 0 -> OK
'           -1 -> Error
'------------------------------------------------------------------------------
Public Function actualizaSeguimientoEstadoPersonaSMS(idOferta As Long, _
                                                     strPersonas As String, _
                                                     idEstadoSeg As Integer, _
                                                     idEstadoSegDetalle As Integer) As Integer
    Dim observacion As String
    Dim numPersonas As Integer
    Dim i As Integer
    Dim persona As String
    Dim fkPersona As Long
    Dim idFormaContacto As Integer
    Dim nomEmpresa As String
    Dim puesto As String
    
On Error GoTo TratarError
    
    'Obtenemos el nombre de la empresa
    nomEmpresa = DLookup("[nombre]", _
                         "t_organizacion", _
                         "[id]=" & DLookup("[fkOrganizacion]", _
                                           "t_oferta", _
                                           "[id]=" & idOferta))
    'El puesto de la vacante
    puesto = DLookup("[puesto]", "t_oferta", "[id]=" & idOferta)
    
    'Se monta la observacion
    observacion = "En la oferta " & idOferta & " " & _
                  "de la empresa " & nomEmpresa & vbNewLine & _
                  "Puesto: " & puesto & vbNewLine & _
                  "El seguimiento de usuario cambia de estado a Contactado - Avisado."
    
    numPersonas = cuentaSubstrings(strPersonas)
    
    idFormaContacto = DLookup("[id]", "A_FormaContacto", "[descripcion]='sms'")
    
    'Se crea la gestion grupal tecnico-oferta-candidatos
    creaGestionGrupal 3, _
                      1, _
                      now, _
                      observacion, _
                      U_idIfocUsuarioActivo, _
                      8, _
                      Replace(strPersonas, ";", ","), _
                      , _
                      idOferta, _
                      , _
                      , _
                      idFormaContacto

    'Actualiza el estado de los candidatos de la oferta
    updEstadosPersonasOferta idOferta, _
                            strPersonas, _
                            idEstadoSeg, _
                            idEstadoSegDetalle 'Estado Comunicado(2) y Detalle Avisado(1)
    
    actualizaSeguimientoEstadoPersonaSMS = 0
    
SalirTratarError:
    Exit Function
TratarError:
    actualizaSeguimientoEstadoPersonaSMS = -1
    debugando "ActualizandoSeguimientoEstadoPersonaSMS" & Err.description
    Resume SalirTratarError
End Function

'------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  09/07/2009 - Actualizacion: 15/03/2010
'   Name:   updEstadoOfertaCandidato
'   Desc:   Actualiza el estado del candidato y se incluye en el histórico de estados
'   Param:  ID de la oferta, Id de la persona, estado, detalle del estado, servicio,
'           observaciones y tipo de técnico
'   Return: 0 -> OK
'           -1 -> Error
'------------------------------------------------------------------------------
Public Function updEstadoOfertaCandidato(idOferta As Long, _
                                         idPersona As Long, _
                                         idEstadoCandidato As Integer, _
                                         Optional idEstadoDetalleCandidato As Integer = 0, _
                                         Optional idServicio As Integer = 0, _
                                         Optional observaciones As String = "", _
                                         Optional tecnico As String = "", _
                                         Optional fecha As Date = "01/01/1900 00:00:00") As Integer
    Dim strSql As String
    Dim mifecha As Date
 On Error GoTo TratarError
 
    If (fecha = "01/01/1900 00:00:00") Then
        mifecha = now()
    Else
        mifecha = fecha
    End If
    'Si el estado de la oferta es Modificada (Reabierta), entonces los estados de los
    ' candidatos tendrán la fecha del cierre de la oferta
    If (getidEstadoActualOferta(idOferta) = 6) Then
        fecha = Nz(DLookup("[fechaCierre]", "[t_oferta]", "[id] = " & idOferta), Format(mifecha, "mm/dd/yyyy hh:mm:ss"))
    End If
    
    'SQL para la actualización del estado del candidato en la tabla r_ofertacandidatos
    strSql = " UPDATE r_ofertacandidatos" & _
                 " SET fkIFOCUsuario = " & usuarioIFOC() & ", observacion = """ & observaciones & _
                 """, r_ofertacandidatos.fkOfertaSegEstado = " & idEstadoCandidato & ", r_ofertacandidatos.fkOfertaSegEstadoDetalle = " & _
                 IIf(idEstadoDetalleCandidato = 0, "Null", idEstadoDetalleCandidato) & ", r_ofertacandidatos.fkServicio = " & _
                 IIf(idServicio = 0, "Null", idServicio) & ", r_ofertacandidatos.fecha =#" & Format(mifecha, "mm/dd/yyyy hh:mm:ss") & "#" & _
                 IIf(tecnico = "", ", r_ofertacandidatos.tipotecnico = Null", ", r_ofertacandidatos.tipotecnico = '" & tecnico & "'") & _
                 " WHERE (((r_ofertacandidatos.fkOferta) = " & idOferta & ") AND (r_ofertacandidatos.fkPersona = " & idPersona & "));"
    CurrentDb.Execute strSql
    
    'SQL para la inserción del nuevo cambio de estado en el histórico de estados del candidato
    strSql = " INSERT INTO r_ofertacandidatoshistorico (fkOferta,  fkPersona, fkIFOCUsuario, fechaHora, observacion, " & _
                 "fkOfertaSegEstado, fkOfertaSegEstadoDetalle, fkServicio, tipotecnico) VALUES ( " & idOferta & ", " & idPersona & _
                 ", " & usuarioIFOC() & ", #" & Format(mifecha, "mm/dd/yyyy hh:mm:ss") & "#, """ & observaciones & """, " & idEstadoCandidato & _
                 ", " & IIf(idEstadoDetalleCandidato = 0, "Null", idEstadoDetalleCandidato) & _
                 ", " & IIf(idServicio = 0, "Null", idServicio) & ", " & IIf(tecnico = "", "Null", "'" & tecnico & "'") & " );"
    CurrentDb.Execute strSql
    
    updEstadoOfertaCandidato = 0
    
SalirTratarError:
    Exit Function
TratarError:
    updEstadoOfertaCandidato = -1
    debugando Err.description
    Resume SalirTratarError
End Function

'------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  09/07/2009 - Actualizacion: 15/03/2010
'   Name:   updOfertaCandidatoSeguimiento
'   Desc:   cambio finSeguimento candidato
'   Param:  ID de la oferta, Id de la persona, finSeg
'   Return: 0 -> OK
'           -1 -> Error
'------------------------------------------------------------------------------
Public Function updOfertaCandidatoSeguimiento(idOferta As Long, _
                                            idPersona As Long, _
                                            finSeg As Boolean) As Integer
    Dim strSql As String
    Dim fecha As Date
 On Error GoTo TratarError
 
    'SQL para la actualización del estado del candidato en la tabla r_ofertacandidatos
    strSql = " UPDATE r_ofertacandidatos" & _
             " SET fkIFOCUsuario = " & usuarioIFOC() & _
             IIf(finSeg = True, ", finSeguimiento = -1", ", finSeguimiento = 0") & _
             " WHERE (((r_ofertacandidatos.fkOferta) = " & idOferta & ") AND (r_ofertacandidatos.fkPersona = " & idPersona & "));"
'Debug.Print strSql
    CurrentDb.Execute strSql
    
    updOfertaCandidatoSeguimiento = 0
    
SalirTratarError:
    Exit Function
TratarError:
    updOfertaCandidatoSeguimiento = -1
    debugando Err.description
    Resume SalirTratarError
End Function

'-------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  22/09/2009 Actualizado: 24/11/2009
'   Name:   getSqlListadoOfertas
'   Desc:   Calcula el sql con los listados de los ofertas abiertas en el periodo establecido
'   Param:  fechaInicio(date), fecha inicio del periodo del que quieren verse las ofertas
'           fechaFin(date), fecha fin del periodo del que quiere verse las ofertas
'           idIfocUsuario(long), identificador de usuario ifoc(tecnico que lleva oferta)
'   Return:  string, listado de ofertas con los criterios pasados por parametro
'-------------------------------------------------------------------------------------------
Public Function getSqlListadoOfertas(Optional fechaInicio As Date = "01/01/1900", _
                                     Optional fechaFin As Date, _
                                     Optional idOrganizacion As Long = 0, _
                                     Optional idIfocUsuario As Long = 0, _
                                     Optional order As Integer = 0, _
                                     Optional idEstado As Integer = 0) As String
    Dim strselect As String
    Dim strFrom As String
    Dim strWhere As String
    Dim strOrderBy As String
    
    Dim fechaI As Date
    Dim FECHAF As Date
    
    'Inicializamos variables
    strselect = " t_oferta.id," & _
                " t_organizacion.nombre AS Empresa," & _
                " v_ifocusuario.aka AS [Técnic@]," & _
                " t_oferta.numeroPuestos," & _
                " t_oferta.fechaOferta AS Fecha," & _
                " t_oferta.puesto," & _
                " t_oferta.horarioZona," & _
                " a_ofertaestado.descripcion AS Estado," & _
                " t_oferta.fkOrganizacion," & _
                " t_oferta.fechaOferta," & _
                " t_oferta.fkUsuarioIfocAna"
    strFrom = "(((t_oferta LEFT JOIN EstadoDeOferta ON t_oferta.id = EstadoDeOferta.fkOferta) LEFT JOIN t_organizacion ON t_oferta.fkOrganizacion = t_organizacion.id) LEFT JOIN v_ifocusuario ON t_oferta.fkUsuarioIFOCAna = v_ifocusuario.fkPersona) LEFT JOIN a_ofertaestado ON EstadoDeOferta.MaxEstadoOferta = a_ofertaestado.id"
    strWhere = ""
    strOrderBy = ""
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"
    
    If (fechaInicio <> "01/01/1900") Then
        strWhere = addConditionWhere(strWhere, _
                                     "t_oferta.fechaOferta Between #" & fechaI & "# AND #" & FECHAF & "#")
    End If
    
    If (idOrganizacion <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_oferta.fkOrganizacion = " & idOrganizacion)
    ElseIf (idIfocUsuario <> 0) Then
        strWhere = addConditionWhere(strWhere, "t_oferta.fkUsuarioIfocAna = " & idIfocUsuario)
    End If
    
    If (order = 1) Then
        strOrderBy = "fechaOferta ASC"
    ElseIf (order = 2) Then
        strOrderBy = "fechaOferta DESC"
    End If
    
    If (idEstado <> 0) Then
        strWhere = addConditionWhere(strWhere, "EstadoDeOferta.MaxEstadoOferta=" & idEstado)
    End If
    
    getSqlListadoOfertas = montarSQL(strselect, _
                                    strFrom, _
                                    strWhere, _
                                    , _
                                    , _
                                    strOrderBy)
End Function

'-------------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Modif:  Antonio Nadal Company
'   Fecha:  17/09/2009   --   Actualizado: 15/03/2010
'   Name:   updEstadosPersonasOferta
'   Desc:   Actualiza los estados de la personas de la oferta que se le pasan por parametro
'   Param:  idOferta(long), identificador de la oferta
'           personas(string), identificadores de la oferta separados por ','
'           idEstadoSeg(integer), identificador de estado de la oferta
'           idEstadoSegDetalle(integer), identificador del detalle del estado de la oferta
'   Return:  0, OK
'           -1, KO
'-------------------------------------------------------------------------------------------
Public Function updEstadosPersonasOferta(idOferta As Long, _
                                         Personas As String, _
                                         idEstadoSeg As Integer, _
                                         idEstadoSegDetalle As Integer) As Integer
    Dim numPersonas As Integer
    Dim args As Variant
    Dim idPersona As Long
    Dim i As Integer
On Error GoTo TratarError
    
    numPersonas = countSubStrings(Personas, ";")
    args = Split(Personas, ";")
        
    For i = 0 To numPersonas - 1
        idPersona = args(i)
        If updEstadoOfertaCandidato(idOferta, _
                                    idPersona, _
                                    idEstadoSeg, _
                                    idEstadoSegDetalle, _
                                    , _
                                    "El candidato ha sido avisado mediante SMS.") = -1 Then
            GoTo TratarError
        End If
    Next
    
    updEstadosPersonasOferta = 0
    
SalirTratarError:
    Exit Function
TratarError:
    updEstadosPersonasOferta = -1
    debugando Err.description
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  23/11/2009 - Actualización:  23/11/2009
'   Name:   insOfertaVacante
'   Desc:   Inserta vacante en oferta vacía
'   Param:  idOferta(long), identificador de oferta
'   Retur:  0, OK
'           -1, KO
'---------------------------------------------------------------------------
Public Function insOfertaVacante(idOferta As Long) As Integer
    Dim str As String
On Error GoTo TratarError

    str = " INSERT INTO t_ofertavacanteresultado ( fkOferta )" & _
          " VALUES (" & idOferta & ");"
    
    CurrentDb.Execute str

SalirTratarError:
    insOfertaVacante = 0
    Exit Function
TratarError:
    insOfertaVacante = -1
    debugando "SIFOC_Oferta.insOfertaVacante" & Err.description
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  23/11/2009 - Actualización:  23/11/2009
'   Name:   delOfertaVacante
'   Desc:   Elimina vacante de oferta
'   Param:  idOferta(long), identificador de oferta
'           idVacante(long), identificador de oferta vacante
'   Retur:   0, OK
'           -1, KO
'---------------------------------------------------------------------------
Public Function delOfertaVacante(idOferta As Long, _
                                 idOfertaVacante As Long) As Integer
    Dim str As String
On Error GoTo TratarError
    
    str = " DELETE id, fkOferta" & _
          " FROM t_ofertavacanteresultado" & _
          " WHERE ((id=" & idOfertaVacante & ") AND (fkOferta=" & idOferta & "));"

    CurrentDb.Execute str
    
SalirTratarError:
    delOfertaVacante = 0
    Exit Function
TratarError:
    delOfertaVacante = -1
    debugando "SIFOC_Oferta.delOfertaVacante" & Err.description
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  23/11/2009 - Actualización:  23/11/2009
'   Name:   getNumVacantes
'   Desc:   Devuelve el número de vacantes de la oferta
'   Param:  idOferta(long), identificador de oferta
'   Retur:  (integer)número de vacantes de la oferta
'---------------------------------------------------------------------------
Public Function getNumVacantes(idOferta As Long) As Integer
    Dim str As String
    Dim rs As ADODB.Recordset

    str = " SELECT fkOferta, count(id) as vacantes" & _
          " FROM t_ofertavacanteresultado" & _
          " WHERE fkOferta=" & idOferta & _
          " GROUP BY fkOferta;"
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    getNumVacantes = rs!vacantes
    
    rs.Close
    Set rs = Nothing
End Function

'---------------------------------------------------------------------------
'   Autor:  Antonio Nadal Company
'   Fecha:  11/13/2010
'   Name:   getidEstadoActualOferta
'   Desc:   Devuelve el estado actual en el que se encuentra la oferta
'   Param:  idOferta(long), identificador de oferta
'   Retur:  (integer)Id del estado de la oferta
'---------------------------------------------------------------------------
Public Function getidEstadoActualOferta(idOferta As Long) As Integer
    
    getidEstadoActualOferta = Nz(DLookup("[fkofertaestado]", "[t_oferta]", "[id] = " & idOferta), 0)

End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/06/2011
'   Name:   getOfertaEstado
'   Desc:   Devuelve el estado de la oferta
'   Param:  idOferta(long), identificador de oferta
'   Retur:  (integer)Id del estado de la oferta
'---------------------------------------------------------------------------
Public Function getOfertaEstado(idOferta As Long) As String
    
    getOfertaEstado = Nz(DLookup("descripcion", "a_ofertaestado", "[id] = " & getidEstadoActualOferta(idOferta)), 0)

End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/06/2011
'   Name:   getOfertaEstadoCandidato
'   Desc:   Devuelve el estado del candidato en la oferta
'   Param:  idOferta(long), identificador de oferta
'           idPersona(long), identificador de persona
'   Retur:  (integer)Id del estado de la oferta
'---------------------------------------------------------------------------
Public Function getOfertaEstadoCandidato(idOferta As Long, _
                                         idPersona As Long) As String
    Dim estado As String
    Dim Detalle As String
Dim ide As Integer
idOferta = Nz(DLookup("fkOfertaSegEstado", "r_ofertacandidatos", "fkOferta = " & idOferta & " AND fkPersona = " & idPersona), 0)
    estado = Nz(DLookup("estado", _
                        "a_ofertasegestado", _
                        "id=" & _
                        ide), "")
    Detalle = Nz(DLookup("estado", _
                         "a_ofertasegestadodetalle", _
                         "id=" & Nz(DLookup("fkOfertaSegEstadoDetalle", "r_ofertacandidatos", "fkOferta = " & idOferta & " AND fkPersona = " & idPersona), 0)), "")
    
    getOfertaEstadoCandidato = estado & " - " & Detalle

End Function

Public Function getEstadoActualOfertaAntigua(idOferta As Long) As Integer
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim estado As Integer
    
    str = "SELECT t_ofertaestados.fkOferta, t_ofertaestados.fecha, t_ofertaestados.fkOfertaEstado as estadooferta " & _
          "FROM t_ofertaestados WHERE (((t_ofertaestados.fkOferta) = " & idOferta & ")) " & _
          "ORDER BY t_ofertaestados.fecha DESC;"

    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
Debug.Print rs.RecordCount
    If Not rs.EOF Then
        rs.MoveFirst
        estado = rs!estadoOferta
        getEstadoActualOfertaAntigua = estado
    Else
        getEstadoActualOfertaAntigua = 0
    End If
    
    rs.Close
    Set rs = Nothing
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getFiltroDemandantesCNO
'   Desc:   Devuelve los CNO's seleccionados en el filtro, para cada CNO devuelve también:
'               Hasta 3 Niveles (los que se hayan seleccionado en el filtro)
'               Meses de Experiencia (si se hubiera seleccionado en el filtro)
'               Hasta 3 Nivels de Experiencia (si se hubieran seleccionado en el filtro)
'           Se aplica sobre el form FiltroDemandantes
'   Retur:  Devuelve un String que contiene la información previa, con un salto de línea después de cada CNO
'-------------------------------------------------------------------------------------------------------------------
Public Function getFiltroDemandantesCNO() As String
    Dim str As String
    Dim numReg As Integer
    Dim rs As New ADODB.Recordset
    getFiltroDemandantesCNO = ""
    str = " SELECT a_cno2011.ocupacion, t_cnofiltro.Principal, t_cnofiltro.mesesExp, " & _
          " t_cnofiltro.nivelCno1, t_cnofiltro.nivelExp1, t_cnofiltro.nivelCno2, " & _
          " t_cnofiltro.nivelExp2, t_cnofiltro.nivelCno3,  t_cnofiltro.nivelExp3 " & _
          " FROM t_cnofiltro INNER JOIN a_cno2011 ON t_cnofiltro.fkCno2011=a_cno2011.id " & _
          " WHERE t_cnofiltro.fkOferta = " & Forms!FiltroDemandantes!txt_fkOferta & _
          "   AND t_cnofiltro.fkPerfilFiltro = " & Forms!FiltroDemandantes!txt_idPerfil
    rs.Open str, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
    numReg = rs.RecordCount
    If numReg > 0 Then
        rs.MoveFirst 'como numreg > 0 es que hay algun registro CNO
        While Not rs.EOF
            If rs!principal = -1 Then
                getFiltroDemandantesCNO = getFiltroDemandantesCNO & "[CNO Pr]"
            Else
                getFiltroDemandantesCNO = getFiltroDemandantesCNO & "[CNO]"
            End If
            getFiltroDemandantesCNO = getFiltroDemandantesCNO & "(" & rs!ocupacion & ")"
            If Trim(rs!nivelCno1) <> "" Or Trim(rs!nivelCno2) <> "" Or Trim(rs!nivelCno2) <> "" Then
                getFiltroDemandantesCNO = getFiltroDemandantesCNO & " Niveles("
                If Trim(rs!nivelCno1) <> "" Then
                    getFiltroDemandantesCNO = getFiltroDemandantesCNO & Trim(rs!nivelCno1)
                End If
                If Trim(rs!nivelCno2) <> "" Then
                    getFiltroDemandantesCNO = getFiltroDemandantesCNO & ", " & Trim(rs!nivelCno2)
                End If
                If Trim(rs!nivelCno3) <> "" Then
                    getFiltroDemandantesCNO = getFiltroDemandantesCNO & ", " & Trim(rs!nivelCno3)
                End If
                getFiltroDemandantesCNO = getFiltroDemandantesCNO & ")"
            End If
            If Trim(rs!mesesExp) <> "" Then
                getFiltroDemandantesCNO = getFiltroDemandantesCNO & " Meses Exp(" & Trim(rs!mesesExp) & ")"
            End If
            If Trim(rs!nivelExp1) <> "" Or Trim(rs!nivelExp2) <> "" Or Trim(rs!nivelExp3) <> "" Then
                getFiltroDemandantesCNO = getFiltroDemandantesCNO & " Nivel Exp("
                If Trim(rs!nivelExp1) <> "" Then
                    getFiltroDemandantesCNO = getFiltroDemandantesCNO & Trim(rs!nivelExp1)
                End If
                If Trim(rs!nivelExp2) <> "" Then
                    getFiltroDemandantesCNO = getFiltroDemandantesCNO & ", " & Trim(rs!nivelExp2)
                End If
                If Trim(rs!nivelExp3) <> "" Then
                    getFiltroDemandantesCNO = getFiltroDemandantesCNO & ", " & Trim(rs!nivelExp3)
                End If
                getFiltroDemandantesCNO = getFiltroDemandantesCNO & ")"
            End If
            getFiltroDemandantesCNO = getFiltroDemandantesCNO & vbNewLine
            rs.MoveNext
        Wend
    End If
    rs.Close
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getCampoLeido
'   Desc:   Ejecuta un SELECT que devuelve el valor de un único campo
'           IMPORTANTE ===>>> Debe tener el alias de campoLeido <<< Ej.: SELECT campo AS campoLeido FROM nombreTabla ...
'   Param:  stringSelect(String) Contiene el SELECT a ejecutar.
'   Retur:  Devuelve el valor del campo leído, o un campo vacío si este no se hubiera leído
'-------------------------------------------------------------------------------------------------------------------
Public Function getCampoLeido(stringSelect As String) As String
    Dim rs As New ADODB.Recordset
    Dim numReg As Integer
    rs.Open stringSelect, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
    numReg = rs.RecordCount
    If numReg > 0 Then
        rs.MoveFirst
        getCampoLeido = rs!campoLeido
    Else
        getCampoLeido = ""
    End If
    rs.Close
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getFiltroDemandantesGeneral
'   Desc:   Lee la selección SEXO, EDAD MÍNIMA y EDAD MÁXIMA del filtro (si se hubiera seleccionado)
'           Se aplica sobre el form FiltroDemandantes
'   Retur:  Devuelve un String con esa información y un salto de línea al final, o un campo vacío si no hubiera selección
'-------------------------------------------------------------------------------------------------------------------
Public Function getFiltroDemandantesGeneral() As String
    Dim str As String
    getFiltroDemandantesGeneral = "[General]"
    If Forms!FiltroDemandantes!sexo <> "" Then
        str = "SELECT sexo AS campoLeido FROM a_sexo WHERE id = " & Forms!FiltroDemandantes!sexo
        getFiltroDemandantesGeneral = getFiltroDemandantesGeneral & "Sexo(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!txt_edadMin <> "" Then
        getFiltroDemandantesGeneral = getFiltroDemandantesGeneral & "Edad Mín(" & Forms!FiltroDemandantes!txt_edadMin & ")"
    End If
    If Forms!FiltroDemandantes!txt_edadMax <> "" Then
        getFiltroDemandantesGeneral = getFiltroDemandantesGeneral & "Edad Máx(" & Forms!FiltroDemandantes!txt_edadMax & ")"
    End If
    If getFiltroDemandantesGeneral = "[General]" Then
        getFiltroDemandantesGeneral = ""
    Else
        getFiltroDemandantesGeneral = getFiltroDemandantesGeneral & vbNewLine
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getFiltroDemandantesIdiomas
'   Desc:   Lee la selección, si se hubiera realizado, de hasta 3 idiomas y sus correspondientes niveles.
'           Se aplica sobre el form FiltroDemandantes
'   Retur:  Devuelve un String con esa información y un salto de línea al final, o un campo vacío si no hubiera selección
'-------------------------------------------------------------------------------------------------------------------
Public Function getFiltroDemandantesIdiomas() As String
    Dim str As String
    getFiltroDemandantesIdiomas = "[Idiomas]"
    If Forms!FiltroDemandantes!cbx_idioma1 <> "" Then
        str = "SELECT idioma AS campoLeido FROM a_idioma WHERE id = " & Forms!FiltroDemandantes!cbx_idioma1
        getFiltroDemandantesIdiomas = getFiltroDemandantesIdiomas & getCampoLeido(str)
    End If
    If Forms!FiltroDemandantes!cbx_nivelIdioma1 <> "" Then
        str = "SELECT nivel AS campoLeido FROM a_idiomanivelsimple WHERE id = " & Forms!FiltroDemandantes!cbx_nivelIdioma1
        getFiltroDemandantesIdiomas = getFiltroDemandantesIdiomas & "(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!cbx_idioma2 <> "" Then
        str = "SELECT idioma AS campoLeido FROM a_idioma WHERE id = " & Forms!FiltroDemandantes!cbx_idioma2
        getFiltroDemandantesIdiomas = getFiltroDemandantesIdiomas & getCampoLeido(str)
    End If
    If Forms!FiltroDemandantes!cbx_nivelIdioma2 <> "" Then
        str = "SELECT nivel AS campoLeido FROM a_idiomanivelsimple WHERE id = " & Forms!FiltroDemandantes!cbx_nivelIdioma2
        getFiltroDemandantesIdiomas = getFiltroDemandantesIdiomas & "(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!cbx_idioma3 <> "" Then
        str = "SELECT idioma AS campoLeido FROM a_idioma WHERE id = " & Forms!FiltroDemandantes!cbx_idioma3
        getFiltroDemandantesIdiomas = getFiltroDemandantesIdiomas & getCampoLeido(str)
    End If
    If Forms!FiltroDemandantes!cbx_nivelIdioma3 <> "" Then
        str = "SELECT nivel AS campoLeido FROM a_idiomanivelsimple WHERE id = " & Forms!FiltroDemandantes!cbx_nivelIdioma3
        getFiltroDemandantesIdiomas = getFiltroDemandantesIdiomas & "(" & getCampoLeido(str) & ")"
    End If
    If getFiltroDemandantesIdiomas = "[Idiomas]" Then
        getFiltroDemandantesIdiomas = ""
    Else
        getFiltroDemandantesIdiomas = getFiltroDemandantesIdiomas & vbNewLine
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getFiltroDemandantesInformatica
'   Desc:   Lee la selección, si se hubiera realizado, de hasta 3 conocimientos de informática y sus correspondientes niveles.
'           Se aplica sobre el form FiltroDemandantes
'   Retur:  Devuelve un String con esa información y un salto de línea al final, o un campo vacío si no hubiera selección
'-------------------------------------------------------------------------------------------------------------------
Public Function getFiltroDemandantesInformatica() As String
    Dim str As String
    getFiltroDemandantesInformatica = "[Informática]"
    If Forms!FiltroDemandantes!cbx_informatica1 <> "" Then
        str = "SELECT informatica AS campoLeido FROM a_informatica WHERE id = " & Forms!FiltroDemandantes!cbx_informatica1
        getFiltroDemandantesInformatica = getFiltroDemandantesInformatica & getCampoLeido(str)
    End If
    If Forms!FiltroDemandantes!cbx_nivelInformatica1 <> "" Then
        str = "SELECT nivel AS campoLeido FROM a_nivel WHERE id = " & Forms!FiltroDemandantes!cbx_nivelInformatica1
        getFiltroDemandantesInformatica = getFiltroDemandantesInformatica & "(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!cbx_informatica2 <> "" Then
        str = "SELECT informatica AS campoLeido FROM a_informatica WHERE id = " & Forms!FiltroDemandantes!cbx_informatica2
        getFiltroDemandantesInformatica = getFiltroDemandantesInformatica & getCampoLeido(str)
    End If
    If Forms!FiltroDemandantes!cbx_nivelInformatica2 <> "" Then
        str = "SELECT nivel AS campoLeido FROM a_nivel WHERE id = " & Forms!FiltroDemandantes!cbx_nivelInformatica2
        getFiltroDemandantesInformatica = getFiltroDemandantesInformatica & "(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!cbx_informatica3 <> "" Then
        str = "SELECT informatica AS campoLeido FROM a_informatica WHERE id = " & Forms!FiltroDemandantes!cbx_informatica3
        getFiltroDemandantesInformatica = getFiltroDemandantesInformatica & getCampoLeido(str)
    End If
    If Forms!FiltroDemandantes!cbx_nivelInformatica3 <> "" Then
        str = "SELECT nivel AS campoLeido FROM a_nivel WHERE id = " & Forms!FiltroDemandantes!cbx_nivelInformatica3
        getFiltroDemandantesInformatica = getFiltroDemandantesInformatica & "(" & getCampoLeido(str) & ")"
    End If
    If getFiltroDemandantesInformatica = "[Informática]" Then
        getFiltroDemandantesInformatica = ""
    Else
        getFiltroDemandantesInformatica = getFiltroDemandantesInformatica & vbNewLine
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getFiltroDemandantesFormacion
'   Desc:   Lee la selección, si se hubiera realizado, de:
'               Nivel Formativo
'               Hasta 2 Titulaciones
'               Formación complementaria
'               Fecha de inserción
'               Hasta 3 Carnés Profesionales y sus correspondientes niveles (si los hubiera)
'           Se aplica sobre el form FiltroDemandantes
'   Retur:  Devuelve un String con esa información y un salto de línea al final, o un campo vacío si no hubiera selección
'-------------------------------------------------------------------------------------------------------------------
Public Function getFiltroDemandantesFormacion() As String
    Dim str As String
    getFiltroDemandantesFormacion = "[Nivel Formativo]"
    If Forms!FiltroDemandantes!cbx_NivelFormacion <> "" Then
        str = "SELECT nivel AS campoLeido FROM A_NivelFormacion " & _
              "LEFT JOIN A_NivelFormacionSoib ON A_NivelFormacion.fkNivelFormacionSoib = A_NivelFormacionSoib.id " & _
              "WHERE a_nivelformacion.id = " & Forms!FiltroDemandantes!cbx_NivelFormacion
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & getCampoLeido(str)
    End If
    If Forms!FiltroDemandantes!cbx_titulacion1 <> "" Then
        str = "SELECT Titulacion AS campoLeido FROM T_Titulacion WHERE id = " & Forms!FiltroDemandantes!cbx_titulacion1
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & "(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!cbx_titulacion2 <> "" Then
        str = "SELECT Titulacion AS campoLeido FROM T_Titulacion WHERE id = " & Forms!FiltroDemandantes!cbx_titulacion2
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & "(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!formacionComp <> "" Then
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & "Form.Compl.(" & Forms!FiltroDemandantes!formacionComp & ")"
    End If
    If Forms!FiltroDemandantes!txt_fechaInsercion <> "" Then
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & "(Sin inserción posterior a " & Forms!FiltroDemandantes!txt_fechaInsercion & ")"
    End If
    If Forms!FiltroDemandantes!cbx_carneProfesional1 <> "" Then
        str = "SELECT carneProfesional AS campoLeido FROM a_carneprofesional WHERE id = " & Forms!FiltroDemandantes!cbx_carneProfesional1
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & "Carné Profesional(" & getCampoLeido(str) & ")"
        If Forms!FiltroDemandantes!cbx_nivelCp1 <> "" Then
            str = "SELECT nivel AS campoLeido FROM a_carneprofesionalnivel WHERE id = " & Forms!FiltroDemandantes!cbx_nivelCp1
            getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & "Nivel(" & getCampoLeido(str) & ")"
        End If
    End If
    If Forms!FiltroDemandantes!cbx_carneProfesional2 <> "" Then
        str = "SELECT carneProfesional AS campoLeido FROM a_carneprofesional WHERE id = " & Forms!FiltroDemandantes!cbx_carneProfesional2
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & ", (" & getCampoLeido(str) & ")"
        If Forms!FiltroDemandantes!cbx_nivelCp2 <> "" Then
            str = "SELECT nivel AS campoLeido FROM a_carneprofesionalnivel WHERE id = " & Forms!FiltroDemandantes!cbx_nivelCp2
            getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & "Nivel(" & getCampoLeido(str) & ")"
        End If
    End If
    If Forms!FiltroDemandantes!cbx_carneProfesional3 <> "" Then
        str = "SELECT carneProfesional AS campoLeido FROM a_carneprofesional WHERE id = " & Forms!FiltroDemandantes!cbx_carneProfesional3
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & ", (" & getCampoLeido(str) & ")"
        If Forms!FiltroDemandantes!cbx_nivelCp3 <> "" Then
            str = "SELECT nivel AS campoLeido FROM a_carneprofesionalnivel WHERE id = " & Forms!FiltroDemandantes!cbx_nivelCp3
            getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & "Nivel(" & getCampoLeido(str) & ")"
        End If
    End If
    If getFiltroDemandantesFormacion = "[Nivel Formativo]" Then
        getFiltroDemandantesFormacion = ""
    Else
        getFiltroDemandantesFormacion = getFiltroDemandantesFormacion & vbNewLine
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getFiltroDemandantesDispJornada
'   Desc:   Lee la selección, si se hubiera realizado, de:
'               Disponibilidad jornada
'               Disponibilidad distribución horario
'               Disponibilidad semana
'           Se aplica sobre el form FiltroDemandantes
'   Retur:  Devuelve un String con esa información y un salto de línea al final, o un campo vacío si no hubiera selección
'-------------------------------------------------------------------------------------------------------------------
Public Function getFiltroDemandantesDispJornada() As String
    Dim str As String
    getFiltroDemandantesDispJornada = "[Disp.Jornada]"
    If Forms!FiltroDemandantes!cbx_dispJornadaDuracion <> "" Then
        str = "SELECT jornada AS campoLeido FROM A_DisponibilidadJornada WHERE id = " & Forms!FiltroDemandantes!cbx_dispJornadaDuracion
        getFiltroDemandantesDispJornada = getFiltroDemandantesDispJornada & "Duración(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!cbx_dispHorario <> "" Then
        str = "SELECT DistribucionHorario AS campoLeido FROM a_disponibilidaddistribucionhorario WHERE id = " & Forms!FiltroDemandantes!cbx_dispHorario
        getFiltroDemandantesDispJornada = getFiltroDemandantesDispJornada & "Jornada(" & getCampoLeido(str) & ")"
    End If
    If Forms!FiltroDemandantes!cbx_DispSemanal <> "" Then
        str = "SELECT disponibilidadsemana AS campoLeido FROM a_disponibilidadsemana WHERE id = " & Forms!FiltroDemandantes!cbx_DispSemanal
        getFiltroDemandantesDispJornada = getFiltroDemandantesDispJornada & "Semana(" & getCampoLeido(str) & ")"
    End If
    If getFiltroDemandantesDispJornada = "[Disp.Jornada]" Then
        getFiltroDemandantesDispJornada = ""
    Else
        getFiltroDemandantesDispJornada = getFiltroDemandantesDispJornada & vbNewLine
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getFiltroDemandantesDispContrato
'   Desc:   Lee la selección, si se hubiera realizado, de Salario y Población.
'           Se aplica sobre el form FiltroDemandantes
'   Retur:  Devuelve un String con esa información y un salto de línea al final, o un campo vacío si no hubiera selección
'-------------------------------------------------------------------------------------------------------------------
Public Function getFiltroDemandantesDispContrato() As String
    Dim str As String
    getFiltroDemandantesDispContrato = "[Disp.Contrato]"
    If Trim(Forms!FiltroDemandantes!cbx_dispSalario) <> "" Then
        getFiltroDemandantesDispContrato = getFiltroDemandantesDispContrato & "Salario(" & Trim(Forms!FiltroDemandantes!cbx_dispSalario) & ")"
    End If
    If Forms!FiltroDemandantes!cbx_poblacion <> "" Then
        str = "SELECT poblacion AS campoLeido FROM a_poblacion WHERE id = " & Forms!FiltroDemandantes!cbx_poblacion
        getFiltroDemandantesDispContrato = getFiltroDemandantesDispContrato & "Población(" & getCampoLeido(str) & ")"
    End If
    If getFiltroDemandantesDispContrato = "[Disp.Contrato]" Then
        getFiltroDemandantesDispContrato = ""
    Else
        getFiltroDemandantesDispContrato = getFiltroDemandantesDispContrato & vbNewLine
    End If
End Function

'-------------------------------------------------------------------------------------------------------------------
'   Autor:  José Espases Abraham
'   Fecha:  18/11/2010
'   Name:   getFiltroDemandantesVehiculos
'   Desc:   Lee la selección, si se hubiera realizado, de disponibilidad de Coche, Moto, Furgoneta y/o Camión
'           Se aplica sobre el form FiltroDemandantes
'   Retur:  Devuelve un String con esa información y un salto de línea al final, o un campo vacío si no hubiera selección
'-------------------------------------------------------------------------------------------------------------------
Public Function getFiltroDemandantesVehiculos() As String
    getFiltroDemandantesVehiculos = "[Vehículos]"
    If Forms!FiltroDemandantes!chb_coche <> 0 Then
        getFiltroDemandantesVehiculos = getFiltroDemandantesVehiculos & "(Coche)"
    End If
    If Forms!FiltroDemandantes!chb_moto <> 0 Then
        getFiltroDemandantesVehiculos = getFiltroDemandantesVehiculos & "(Moto)"
    End If
    If Forms!FiltroDemandantes!chb_furgoneta <> 0 Then
        getFiltroDemandantesVehiculos = getFiltroDemandantesVehiculos & "(Furgoneta)"
    End If
    If Forms!FiltroDemandantes!chb_camion <> 0 Then
        getFiltroDemandantesVehiculos = getFiltroDemandantesVehiculos & "(Camión)"
    End If
    If Trim(getFiltroDemandantesVehiculos) = "[Vehículos]" Then
        getFiltroDemandantesVehiculos = ""
    Else
        getFiltroDemandantesVehiculos = getFiltroDemandantesVehiculos & vbNewLine
    End If
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  20/02/2013 - Actualización:  20/02/2013
'   Name:   getFiltroCertificadoProf
'   Desc:   obtiene el certificado profesional si hubiera seleccionado alguno
'   Param:
'   Retur:  String con grupo2 del certificado de profesionalidad
'---------------------------------------------------------------------------
Public Function getFiltroCertificadoProf() As String
    If Forms!FiltroDemandantes!cbx_grupoFormacion2cert <> "" Then
        getFiltroCertificadoProf = "[Certificado prof.]" & _
                DLookup("grupo2", "a_grupoformacion2", "[id]=" & Forms!FiltroDemandantes!cbx_grupoFormacion2cert)
    End If
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  20/02/2013 - Actualización:  20/02/2013
'   Name:   getFiltroFormacionComp
'   Desc:   obtiene el certificado profesional si hubiera seleccionado alguno
'   Param:
'   Retur:  String con la formación buscada en filtro
'---------------------------------------------------------------------------
Public Function getFiltroFormacionComp() As String

   If Forms!FiltroDemandantes!txt_formacionComp <> "" Then
        getFiltroFormacionComp = "[Formación comp]=" & Forms!FiltroDemandantes!txt_formacionComp
   End If
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  28/06/2011 - Actualización:  28/06/2011
'   Name:   openFrmGestionOfertaCambioEstadoCandidato
'   Desc:   Abre formulario GestionOfertaCambioEstadoCandidato
'   Param:  idPersona(long), identificador de persona
'
'   Retur:  String con apellido y nombre de la persona activa
'---------------------------------------------------------------------------
Public Function openFrmGestionOfertaCambioEstadoCandidato(idPersona As Long, _
                                                          idOferta As Long)

    'Dim db As database
    Dim rst As dao.Recordset
    Dim strSql As String
        
    Dim frm As New Form_GestionOfertaCambioEstadoCandidato
    
    frm.setIdPersona (idPersona)
    frm.setIdOferta (idOferta)
    frm.actualizaForm
    
    frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  10/04/2013 - Actualización:  10/04/2013
'   Name:   isVacanteFilled
'   Desc:   indica si la oferta tiene todas las vacantes rellenas
'   Param:  idPersona(long), identificador de persona
'
'   Retur:  True    vacantes rellenadas
'           False   vacantes sin rellenar
'---------------------------------------------------------------------------
Public Function isVacanteFilled(idOferta As Long) As Boolean

    Dim strSql As String
    Dim numReg As Integer
    Dim rs As dao.Recordset
    
    strSql = " SELECT t_ofertavacanteresultado.fkOferta, t_ofertavacanteresultado.fkOfertaVacanteResultado" & _
             " FROM t_ofertavacanteresultado" & _
             " WHERE (t_ofertavacanteresultado.fkOfertaVacanteResultado Is Not Null) AND t_ofertavacanteresultado.fkOferta = " & idOferta
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'rs.Open str, CurrentProject.Connection, adOpenKeyset, adLockReadOnly

    If Not rs.EOF Then ' Si hay registros nulos -> no están rellenos los campos
        isVacanteFilled = True
    Else
        isVacanteFilled = False
    End If
    
End Function

Public Static Function getInfoOfertas(idIfocUsuario As Long, fechaInicio As Date, fechaFin As Date) As Variant
    Dim num As Integer
    Dim strSql, sqlCandidatos As String
    Dim rs As dao.Recordset
    Dim fechaI, FECHAF As Date
    Dim estOfertas(3) As Integer
    
    fechaI = Format(fechaInicio, "mm/dd/yyyy")
    FECHAF = Format(fechaFin, "mm/dd/yyyy") & " 23:59:59"
    
    sqlCandidatos = " SELECT fkOferta as idOferta, Count(fkPersona) as personas" & _
                  " FROM r_ofertacandidatos" & _
                  " GROUP BY fkOferta"

    strSql = " SELECT Count(t_oferta.id) as ofertas, Sum(can.personas) as personas" & _
             " FROM t_oferta LEFT JOIN (" & sqlCandidatos & ") as can ON t_oferta.id = can.idOferta" & _
             " WHERE (t_oferta.fechaOferta Between #" & fechaI & "# AND #" & FECHAF & "#)" & _
             " AND ((t_oferta.fkUsuarioIFOCrec =" & idIfocUsuario & ") OR (t_oferta.fkUsuarioIFOCana =" & idIfocUsuario & "))"
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        'Acciones realizadas, Tiempo
        estOfertas(1) = Nz(rs!ofertas, 0)
        estOfertas(2) = Nz(rs!Personas, 0)
    End If
    
    rs.Close
    Set rs = Nothing
    
    getInfoOfertas = estOfertas
    
End Function

Public Static Function getNumOfertasYear() As Integer
    Dim num As Integer
    Dim strSql As String
    Dim rs As dao.Recordset
    Dim fechaI, FECHAF As Date
    
    fechaI = Format(DateSerial(Year(now), 1, 1), "mm/dd/yyyy")
    FECHAF = Format(DateSerial(Year(now), 12, 31), "mm/dd/yyyy") & " 23:59:59"

    strSql = " SELECT t_oferta.id" & _
             " FROM t_oferta" & _
             " WHERE (t_oferta.fechaOferta Between #" & fechaI & "# AND #" & FECHAF & "#)"
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        rs.MoveLast
        num = rs.RecordCount
    End If
    
    rs.Close
    Set rs = Nothing
    
    getNumOfertasYear = num
    
End Function

Public Static Function getNumPuestosDeOfertasYear() As Integer
    Dim num As Integer
    Dim strSql As String
    Dim rs As dao.Recordset
    Dim fechaI, FECHAF As Date
    
    fechaI = Format(DateSerial(Year(now), 1, 1), "mm/dd/yyyy")
    FECHAF = Format(DateSerial(Year(now), 12, 31), "mm/dd/yyyy") & " 23:59:59"

    strSql = " SELECT Sum(t_oferta.numeroPuestos) as puestos" & _
             " FROM t_oferta" & _
             " WHERE (t_oferta.fechaOferta Between #" & fechaI & "# AND #" & FECHAF & "#)"
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        num = rs!puestos
    End If
    
    rs.Close
    Set rs = Nothing
    
    getNumPuestosDeOfertasYear = num
    
End Function

Public Static Function getNumEmpresasDeOfertasYear() As Integer
    Dim num As Integer
    Dim strSql As String
    Dim rs As dao.Recordset
    Dim fechaI, FECHAF As Date
    
    fechaI = Format(DateSerial(Year(now), 1, 1), "mm/dd/yyyy")
    FECHAF = Format(DateSerial(Year(now), 12, 31), "mm/dd/yyyy") & " 23:59:59"

    strSql = " SELECT t_oferta.fkOrganizacion" & _
             " FROM t_oferta" & _
             " WHERE (t_oferta.fechaOferta Between #" & fechaI & "# AND #" & FECHAF & "#)" & _
             " GROUP BY fkOrganizacion"
    
    Set rs = u_db.OpenRecordset(strSql, dbOpenSnapshot)
    'Set rs = New ADODB.Recordset
    'rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not (rs.EOF) Then
        rs.MoveLast
        num = rs.RecordCount
    End If
    
    rs.Close
    Set rs = Nothing
    
    getNumEmpresasDeOfertasYear = num
    
End Function

