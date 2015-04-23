Attribute VB_Name = "SIFOC_Diagnostico"
Option Compare Database
Option Explicit

Public U_tipoUsuario As String
Public U_IdVariable As Long
Public U_IdDiagnostico As Long

Public Function NombreUsuarioDeDiagnostico(IdDiagnostico As Long)

    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String

    Set db = CurrentDb()

    strSql = " SELECT fkPersona, fkOrganizacion, fkProyectoEmprendedor, tipoUsuario" & _
             " FROM t_Diagnostico " & _
             " WHERE Id = " & str(IdDiagnostico)
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    If rst.EOF Then Exit Function
    
    Select Case rst!tipoUsuario
        Case "C"
        
            strSql = "SELECT v_apellidosnombre.name AS Valor FROM v_apellidosnombre " & _
                "WHERE Id = " & str(rst!fkPersona)
        
        
        Case "E"
        
            strSql = "SELECT T_organizacion.nombre AS Valor FROM T_organizacion " & _
                " WHERE Id = " & str(rst!fkOrganizacion)
                
        
        Case "P"
        
            strSql = " SELECT T_ProyectoEmprendedor.NombreProy AS Valor FROM T_ProyectoEmprendedor " & _
                     " WHERE Id = " & str(rst!fkProyectoEmprendedor)
    End Select
    
    rst.Close
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    If Not rst.EOF Then
        NombreUsuarioDeDiagnostico = rst!valor
    End If
    
    rst.Close
    db.Close

End Function

Public Function Get_IdDiagnostico() As Long

    Get_IdDiagnostico = U_IdDiagnostico
    
End Function
Public Function Get_tipoUsuario() As String

    If Len(U_tipoUsuario) = 0 Then U_tipoUsuario = "c"
    Get_tipoUsuario = U_tipoUsuario

End Function

'Author: Jose M. Huerta Guillén
'Date: 03/11/09
'Date update: 03/11/09
'Name: DescripcionVariable
'Descr: Obtiene la descripción de una variable
'Param: Id, de la tabla A_diagnosticoVariable

Public Function DescripcionVariable(id As Variant) As String

    Dim db As dao.database
    Dim strSql As String
    Dim rst As dao.Recordset
    
    
    If IsNull(id) Then Exit Function
    
    Set db = CurrentDb
    strSql = "SELECT codigoArea & ""."" & codigoVariable & "": "" & definicionVariable AS descripcionVariable " & _
        "FROM A_diagnosticoArea INNER JOIN A_diagnosticoVariable " & _
        "ON A_diagnosticoArea.Id = A_diagnosticoVariable.fkDiagnosticoArea " & _
        "WHERE A_diagnosticoVariable.Id=" & str(id)
        
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    If Not rst.EOF Then
        DescripcionVariable = rst!DescripcionVariable
    End If

End Function


'Author: Jose M. Huerta Guillén
'Date: 03/11/09
'Date update: 03/11/09
'Name: DescripcionVariable
'Descr: Obtiene la descripción del valor de una variable
'Param: Id, de la tabla A_diagnosticoVariable, valor, un caracter a-m-b para alto medio bajo.


Public Function DescripcionValorVariable(id As Variant, valor As Variant) As String

    Dim db As dao.database
    Dim strTitulo As String
    Dim strCampo As String
    Dim strSql As String
    
    Dim rst As dao.Recordset
    
    If IsNull(id) Then Exit Function
    If IsNull(valor) Then Exit Function
    
    
    Select Case valor
        Case "a"
            strTitulo = "ALTA"
            strCampo = "textoA"
        Case "m"
            strTitulo = "MEDIA"
            strCampo = "textoM"
        Case "b"
            strTitulo = "BAJA"
            strCampo = "textoB"
        Case Else
            Exit Function
    End Select
    
    
    Set db = CurrentDb
    strSql = "SELECT """ & strTitulo & ": "" & " & strCampo & " AS descripcionValorVariable " & _
        "FROM A_diagnosticoArea INNER JOIN A_diagnosticoVariable " & _
        "ON A_diagnosticoArea.Id = A_diagnosticoVariable.fkDiagnosticoArea " & _
        "WHERE A_diagnosticoVariable.Id=" & str(id)
        
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    If Not rst.EOF Then
        DescripcionValorVariable = rst!DescripcionValorVariable
    End If

End Function

'Author: Jose M. Huerta Guillén
'Date: 03/11/09
'Date update: 03/11/09
'Name: DescripcionIntervencion
'Descr: Obtiene la descripción de una intervención
'Param: Id, de la tabla A_diagnosticoIntervenciones

Public Function DescripcionIntervencion(id As Variant) As String

    Dim db As dao.database
    Dim strSql As String
    Dim rst As dao.Recordset
    
    
    If IsNull(id) Then Exit Function
    
    Set db = CurrentDb
    strSql = "SELECT codigoArea & ""."" & codigoVariable & ""."" & codigoIntervencion & "": "" " & _
        "& definicionIntervencion AS DescripcionIntervencion " & _
        "FROM (A_diagnosticoArea INNER JOIN A_diagnosticoVariable " & _
        "ON A_diagnosticoArea.Id = A_diagnosticoVariable.fkDiagnosticoArea) " & _
        "INNER JOIN A_diagnosticoIntervencion " & _
        "ON A_diagnosticoVariable.Id = A_diagnosticoIntervencion.fkDiagnosticoVariable " & _
        "WHERE A_diagnosticoIntervencion.Id=" & str(id)
        
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    If Not rst.EOF Then
        DescripcionIntervencion = rst!DescripcionIntervencion
    End If

End Function


Public Function AbrirFormularioGestionDiagnostico_IdDiagnostico(IdDiagnostico As Long)
    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
    
    Dim frm As New Form_GestionDiagnostico
    
    Set db = CurrentDb
    
    strSql = "SELECT * FROM T_Diagnostico WHERE Id = " & str(IdDiagnostico)
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    frm.cbx_TipoUsuario = rst!tipoUsuario
    frm.cbx_TipoUsuario_AfterUpdate
    Select Case rst!tipoUsuario
        Case "c"
            frm.cbx_usuario = rst!fkPersona
        Case "e"
            frm.cbx_usuario = rst!fkOrganizacion
        Case "p"
            frm.cbx_usuario = rst!fkProyectoEmpresa
    End Select
    
    frm.cbx_usuario_AfterUpdate
    frm.cbx_Diagnostico = rst!id
    frm.cbx_Diagnostico_AfterUpdate
    
    frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm

End Function

Public Function AbrirFormularioGestionDiagnosticoResumen_IdDiagnostico(IdDiagnostico As Long)

    
    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
    
    Dim frm As New Form_GestionDiagnosticoResumen
    
    frm.Filter = "Id = " & str(IdDiagnostico)
    frm.FilterOn = True
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
        
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FUNCIONES DE APERTURA DE DIAGNOSTICO PARA CIUDADANOS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function AbrirFormularioGestionDiagnosticoResumen_IdCiudadano_Ultimo(IdCiudadano As Long)

    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
    
    Set db = CurrentDb
    
    strSql = "SELECT Id FROM T_diagnostico WHERE tipoUsuario = ""c"" AND fkPersona = " & str(IdCiudadano) & _
        " ORDER BY fecha DESC"
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    If rst.EOF Then
        MsgBox "Este ciudadano no tiene ningún diagnóstico creado."
    Else
        AbrirFormularioGestionDiagnosticoResumen_IdDiagnostico rst!id
    End If

End Function

Public Function AbrirFormularioGestionDiagnostico_IdCiudadano_Ultimo(IdCiudadano As Long)

    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
        
    Dim frm As New Form_GestionDiagnostico
    
    Set db = CurrentDb
    
    strSql = "SELECT Id FROM T_diagnostico WHERE tipoUsuario = ""c"" AND fkPersona = " & str(IdCiudadano) & _
        " ORDER BY fecha DESC"
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    frm.cbx_TipoUsuario = "c"
    frm.cbx_TipoUsuario_AfterUpdate
    frm.cbx_usuario = IdCiudadano
    frm.cbx_usuario_AfterUpdate
    
    If Not rst.EOF Then
        frm.cbx_Diagnostico = rst!id
        frm.cbx_Diagnostico_AfterUpdate
    End If
    
    frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function

Public Function AbrirFormularioGestionDiagnostico_IdCiudadano_Todos(IdCiudadano As Long)

    Dim frm As New Form_GestionDiagnostico
    
    
    frm.cbx_TipoUsuario = "c"
    frm.cbx_TipoUsuario_AfterUpdate
    frm.cbx_usuario = IdCiudadano
    frm.cbx_usuario_AfterUpdate
    
    
    frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FUNCIONES DE APERTURA DE DIAGNOSTICO PARA EMPRESAS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function AbrirFormularioGestionDiagnosticoResumen_IdEmpresa_Ultimo(idEmpresa As Long)

    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
    
    Set db = CurrentDb
    
    strSql = "SELECT Id FROM T_diagnostico WHERE tipoUsuario = ""e"" AND fkOrganizacion = " & str(idEmpresa) & _
        " ORDER BY fecha DESC"
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    If rst.EOF Then
        MsgBox "Esta organización no tiene ningún diagnóstico creado."
    Else
        AbrirFormularioGestionDiagnosticoResumen_IdDiagnostico rst!id
    End If

End Function

Public Function AbrirFormularioGestionDiagnostico_IdEmpresa_Ultimo(idEmpresa As Long)

    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
        
    Dim frm As New Form_GestionDiagnostico
    
    Set db = CurrentDb
    
    strSql = "SELECT Id FROM T_diagnostico WHERE tipoUsuario = ""e"" AND fkOrganizacion = " & str(idEmpresa) & _
        " ORDER BY fecha DESC"
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    frm.cbx_TipoUsuario = "e"
    frm.cbx_TipoUsuario_AfterUpdate
    frm.cbx_usuario = idEmpresa
    frm.cbx_usuario_AfterUpdate
    
    If Not rst.EOF Then
        frm.cbx_Diagnostico = rst!id
        frm.cbx_Diagnostico_AfterUpdate
    End If
    
    frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function

Public Function AbrirFormularioGestionDiagnostico_IdEmpresa_Todos(idEmpresa As Long)

    Dim frm As New Form_GestionDiagnostico
    
    
    frm.cbx_TipoUsuario = "e"
    frm.cbx_TipoUsuario_AfterUpdate
    frm.cbx_usuario = idEmpresa
    frm.cbx_usuario_AfterUpdate
    
    
    frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FUNCIONES DE APERTURA DE DIAGNOSTICO PARA PROYECTOS DE EMPRESA
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function AbrirFormularioGestionDiagnosticoResumen_IdProyecto_Ultimo(idProyecto As Long)

    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
    
    Set db = CurrentDb
    
    strSql = "SELECT Id FROM T_diagnostico WHERE tipoUsuario = ""p"" AND fkProyectoEmprendedor = " & str(idProyecto) & _
        " ORDER BY fecha DESC"
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    If rst.EOF Then
        MsgBox "Este Proyecto Emprededor no tiene ningún diagnóstico creado."
    Else
        AbrirFormularioGestionDiagnosticoResumen_IdDiagnostico rst!id
    End If

End Function

Public Function AbrirFormularioGestionDiagnostico_IdProyecto_Ultimo(idProyecto As Long)

    Dim db As dao.database
    Dim rst As dao.Recordset
    Dim strSql As String
        
    Dim frm As New Form_GestionDiagnostico
    
    Set db = CurrentDb
    
    strSql = "SELECT Id FROM T_diagnostico WHERE tipoUsuario = ""p"" AND fkProyectoEmprendedor = " & str(idProyecto) & _
        " ORDER BY fecha DESC"
    
    Set rst = db.OpenRecordset(strSql, dbOpenSnapshot)
    
    frm.cbx_TipoUsuario = "p"
    frm.cbx_TipoUsuario_AfterUpdate
    frm.cbx_usuario = idProyecto
    frm.cbx_usuario_AfterUpdate
    
    If Not rst.EOF Then
        frm.cbx_Diagnostico = rst!id
        frm.cbx_Diagnostico_AfterUpdate
    End If
    
    frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function

Public Function AbrirFormularioGestionDiagnostico_IdProyecto_Todos(idProyecto As Long)

    Dim frm As New Form_GestionDiagnostico
    
    
    frm.cbx_TipoUsuario = "p"
    frm.cbx_TipoUsuario_AfterUpdate
    frm.cbx_usuario = idProyecto
    frm.cbx_usuario_AfterUpdate
    
    
    frm.NavigationButtons = False
    frm.visible = True
    
    LigaFormulario.LigaFormulario frm
    
End Function



