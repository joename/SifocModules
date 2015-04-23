Attribute VB_Name = "SIFOC_DuplicadosPersonas"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor: Jose Manuel Sanchez
'   Fecha: 1/7/2007
'   Desc:
'   Param: Ahora no se utiliza esta función
'---------------------------------------------------------------------------
Public Function matchDNIs(id1 As Long, ID2 As Long) As Boolean
    Dim idem As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    
    Dim dni1 As String
    Dim dni2 As String
    
    Set rs = New ADODB.Recordset
    
    str = " SELECT id, nombre, apellido1, apellido2, dni" & _
          " FROM t_persona" & _
          " WHERE id=" & id1 & ";"
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    idem = False
    If Not (rs.EOF) Then
        rs.MoveFirst
        
        dni1 = rs!dni
            
        rs.Close
        
        str = " SELECT id, nombre, apellido1, apellido2, dni" & _
              " FROM t_persona" & _
              " WHERE id=" & ID2 & ";"
        rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
        If Not (rs.EOF) Then
            rs.MoveFirst
            
            dni2 = rs!dni
            
        Else
            MsgBox "Error de algúno de los id", vbOKOnly, "Alert: SIFOC_Duplicados"
            idem = False
        End If
        
        'Comprobamos igualdad de dni
        If (dni1 = dni2) Then
            idem = True
        Else
            idem = False
        End If
    Else
        MsgBox "Error de algúno de los id", vbOKOnly, "Alert: SIFOC_Duplicados"
        idem = False
    End If
    
    'Cerramos recordset
    rs.Close
    Set rs = Nothing
    
    matchDNIs = idem
End Function

'---------------------------------------------------------------------------
'   Autor: Jose Manuel Sanchez
'   Modif: José Espases Abraham
'   Fecha: 1/7/2007
'   Fecha Modif.: 24/11/2010
'   Desc:   Unifica los duplicados de personas moviendo toda la información
'           del id2 al id1
'   Param:  ID1 id de la persona base para unificar
'           ID2 id de la persona a eliminar una vez unificada
'---------------------------------------------------------------------------
Public Function trataDuplicado(id1 As Long, ID2 As Long) As String
    trataDuplicado = ""
    trataDuplicado = trataDuplicado & doUpdate("r_serviciousuario", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_cursoalumno", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_ofertapersona", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_organizacionpersona", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_personacitaresultadooloa", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_ofertacandidatos", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_cursopersona", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_personaifocusuario", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_acreditacion", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_cursoalumnoasistencia", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_cursoalumnopractica", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_cursoalumnocalificacion", "fkPersona = " & id1, "fkPersona = " & ID2)
    'trataDuplicado = trataDuplicado & doUpdate("t_autoempleoconsulta", "fkPersona = " & ID1, "fkPersona = " & ID2)
    'trataDuplicado = trataDuplicado & doUpdate("t_autoempleoproyecto", "fkPersona = " & ID1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_citausuario", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_cnodebusqueda", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_colectivo", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_datospersona", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_datossoib", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_derivacion", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_Diagnostico", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_discapacidad", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_disponibilidad", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_disponibilidadMunicipio", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_EntornoFamiliares", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_EntornoSocioeconomico", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_Empleodificultadesbae", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_EmpleoFuentes", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_EmpleoHerramientasBusqueda", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_EmpleoTecnicasBusqueda", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_Empleotiempobae", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_fechacuestionario", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_formador", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("r_gestionusuario", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_discapacidadgrado", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_ifocusuario", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_insercion", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_interesformacion", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_listadocandidatosbrutos", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_permisotrabajo", "fkPersona = " & id1, "fkPersona = " & ID2)
    'trataDuplicado = trataDuplicado & doUpdate("R_AltaBajaPersonaServicio", "fkPersona = " & ID1, "fkPersona = " & ID2)
    'trataDuplicado = trataDuplicado & doUpdate("R_AutoempleoProyectoPersona", "fkPersona = " & ID1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("R_EntornoProcedenciaIngresos", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("R_OfertaCandidatosHistorico", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("T_ConocimientoIfoc", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("T_RegistroEntrada", "fkPersonaRemitente = " & id1, "fkPersonaRemitente = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("T_RegistroSalida", "fkPersonaDestinataria = " & id1, "fkPersonaDestinataria = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("T_Prestaciones", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("T_SeguridadSocial", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("T_WebAccess", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("R_ProyectoEmprendedorPersona", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("T_DiscapacidadOtras", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_formacionnoreglada", "fkPersona = " & id1, "fkPersona = " & ID2)
    trataDuplicado = trataDuplicado & doUpdate("t_formacionreglada", "fkPersona = " & id1, "fkPersona = " & ID2)
    
    'Comproba que si no hi ha tipus pasa dades si ya hi ha un es perdran
    trataDuplicado = trataDuplicado & checkUpdate("t_carneprofesional", id1, ID2, "fkcarneprofesional")
    trataDuplicado = trataDuplicado & checkUpdate("t_direccion", id1, ID2, "fkTipoDireccion")
    trataDuplicado = trataDuplicado & checkUpdate("t_email", id1, ID2, "fkEmailTipo")
    trataDuplicado = trataDuplicado & checkUpdate("t_idioma", id1, ID2, "fkIdioma")
    trataDuplicado = trataDuplicado & checkUpdate("t_informatica", id1, ID2, "fkInformatica")
    trataDuplicado = trataDuplicado & checkUpdate("t_telefono", id1, ID2, "fkTelefonoTipo")
End Function

'-----------------------------------------------------------------
' Fecha......: 24/11/2010
' Autor......: José Espases Abraham
' Descripción: Actualiza campos de una tabla
' Parámetros.: tabla: tabla a actualizar
'              campo: campo/s a actualizar
'              comparativo: comparativo a aplicar en el WHERE
'-----------------------------------------------------------------
Public Function doUpdate(tabla As String, campo As String, comparativo As String) As String
    Dim str As String
    On Error GoTo ControlErrores
    doUpdate = ""
    str = "UPDATE " & tabla & " SET " & campo & " WHERE " & comparativo & ";"
    CurrentDb.Execute str
    Exit Function
ControlErrores:
    doUpdate = str & vbNewLine
    Resume Next
End Function

'-----------------------------------------------------------------
' Fecha......: 24/11/2010
' Autor......: José Espases Abraham
' Descripción: Actualiza campos de una tabla
' Parámetros.: tabla: tabla a actualizar
'              ID1: fkpersona Base
'              ID1: fkpersona Eliminar
'              comparativo: campo a comparar (además de fkpersona)
'-----------------------------------------------------------------
Private Function checkUpdate(tabla As String, id1 As Long, ID2 As Long, comparativo As String) As String
    Dim str As String, strq As String, comp As String
    Dim rs As ADODB.Recordset
    On Error GoTo ControlErrores
    checkUpdate = ""
    'Construye campo a comparar en el IN
    comp = ""
    Set rs = New ADODB.Recordset
    strq = "SELECT " & comparativo & " FROM " & tabla & " WHERE fkpersona = " & id1 & ";"
    rs.Open strq, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    If Not (rs.EOF) Then
        rs.MoveFirst
    End If
    While Not (rs.EOF)
        comp = comp & IIf(comp = "", "", ",")
        Select Case comparativo
            Case "fkcarneprofesional"
                comp = comp & rs!fkCarneProfesional
            Case "fkTipoDireccion"
                comp = comp & rs!fkTipoDireccion
            Case "fkEmailTipo"
                comp = comp & rs!fkEmailTipo
            Case "fkIdioma"
                comp = comp & rs!fkIdioma
            Case "fkInformatica"
                comp = comp & rs!fkInformatica 'rs.fields(nombreCampo)
            Case "fkTelefonoTipo"
                comp = comp & rs!fkTelefonoTipo
        End Select
        rs.MoveNext
    Wend
    rs.Close
    'Borra los resgistros que no se podrán incluir en ID1
    If comp <> "" Then
        str = "DELETE FROM " & tabla & " WHERE fkpersona = " & ID2 & " AND " & comparativo & " IN (" & comp & ");"
        CurrentDb.Execute str
    End If
    'Actualiza los registros de ID2 que pasan a ser de ID1
    str = "UPDATE " & tabla & " SET fkpersona = " & id1 & " WHERE fkpersona = " & ID2 & ";"
    CurrentDb.Execute str

    Exit Function
    
ControlErrores:
    checkUpdate = str & vbNewLine
    Debug.Print str
    Resume Next
    
End Function

