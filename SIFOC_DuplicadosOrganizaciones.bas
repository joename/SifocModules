Attribute VB_Name = "SIFOC_DuplicadosOrganizaciones"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  1/7/2007
'   Descr:
'   Param:
'   Retur:
'---------------------------------------------------------------------------
Public Function matchCIFs(id1 As Long, ID2 As Long) As Boolean
    Dim idem As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    Dim tmp As String
    Dim cif1 As String
    Dim cif2 As String
    
    Set rs = New ADODB.Recordset
    
    str = " SELECT cif" & _
          " FROM T_organizacion" & _
          " WHERE id=" & id1 & ";"
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    idem = False
    If Not (rs.EOF) Then
        rs.MoveFirst
        cif1 = Nz(rs!cif, "")
           ' debugando cif1
        rs.Close
        
        str = " SELECT cif" & _
              " FROM t_organizacion" & _
              " WHERE id=" & ID2 & ";"
        rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        'debugando "Entro aqui"
       
        If Not (rs.EOF) Then
            rs.MoveFirst
            If IsNull(rs!cif) Then
            cif2 = cif1
            'Debug.Print "Entro aqui"
            Else
            cif2 = rs!cif
           ' debugando cif2
            
            End If
        Else
            MsgBox "Error de algúno de los id", vbOKOnly, "Alert: SIFOC_Duplicados"
            idem = False
        End If
        
        'Comprobamos igualdad de los id
        If (cif1 = cif2) Then
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
    
    matchCIFs = idem
End Function

'---------------------------------------------------------------------------
'   Autor: Jose Manuel Sanchez
'   Fecha: 1/7/2007
'   Desc:   Unifica los duplicados de personas moviendo toda la información
'           del id2 al id1
'   Param:  ID1 id de la persona base para unificar
'           ID2 id de la persona a eliminar una vez unificada
'---------------------------------------------------------------------------
Public Function trataDuplicado1(id1 As Long, ID2 As Long) As Integer
    Dim str As String
    Dim rs As ADODB.Recordset
    
    'pasan ID2 -> ID1
    
    If (matchCIFs(id1, ID2)) Then
        joinOrganizacion id1, ID2, "T_Registrosalida", "fkOrganizacionDestinataria"
        joinOrganizacion id1, ID2, "T_RegistroEntrada", "fkOrganizacionRemitente"
        joinOrganizacion id1, ID2, "T_email", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_Telefono", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_Oferta", "fkOrganizacion"
        joinOrganizacion id1, ID2, "r_citausuario", "fkOrganizacion"
        joinOrganizacion id1, ID2, "r_gestionusuario", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_actividadeconomicaorganizacion", "fkOrganizacion"
        joinOrganizacion id1, ID2, "R_organizacionpersona", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_organizacionsolicitudes", "fkOrganizacion"
        'Tablas incluidas el 12/01/2011
        joinOrganizacion id1, ID2, "T_Diagnostico", "fkOrganizacion"
        joinOrganizacion id1, ID2, "R_OrganizacionIFOCUsuario", "fkOrganizacion"
        joinOrganizacion id1, ID2, "R_ServicioUsuario", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_ConocimientoIFOC", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_Insercion", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_ProyectoEmprendedor", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_ProyectoEmprendedorPlanActuacion", "fkOrganizacion"
        joinOrganizacion id1, ID2, "T_webaccess", "fkOrganizacion"
    Else
        MsgBox "Los NIF de las organizaciones a unir no coinciden!!" & vbNewLine & "No se unirán.", vbOKOnly, "Alert: SIFOC_Duplicados"
    End If
End Function

'---------------------------------------------------------------
'                    fusion id por tablas con fkOrganizacion
'-------------------------------------------------------------

Private Function joinOrganizacion(id1 As Long, ID2 As Long, tableName As String, fieldName As String) As Integer
    CurrentDb.Execute " UPDATE " & tableName & _
                      " SET " & fieldName & " = " & id1 & _
                      " WHERE " & fieldName & " = " & ID2 & ";"
End Function

'----------------------------------------------------
'                    Probando
'----------------------------------------------------

Public Function tratar()
    trataDuplicado1 310, 311
End Function

