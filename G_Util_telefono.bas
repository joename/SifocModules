Attribute VB_Name = "G_Util_telefono"
Option Explicit
Option Compare Database

Const separadorMoviles As String = ";"


'devuelve string separado por ; con todos los telefonos moviles para una persona
Public Function telefonosMoviles(idPersona As Long) As String
    Dim strSql As String
    Dim respuesta As String
    Dim rs As ADODB.Recordset
    
    strSql = " SELECT fkPersona, telefono" & _
             " FROM T_Telefono" & _
             " WHERE (fkPersona=" & idPersona & ") AND fkTelefonoTipo=1;"
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    respuesta = ""
    If Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
            If Not IsNull(rs!telefono) Then
                If esMovil(rs!telefono) Then
                    If (respuesta = "") Then
                        respuesta = rs!telefono
                    Else
                        respuesta = respuesta & separadorMoviles & rs!telefono
                    End If
                End If
                rs.MoveNext
            End If
        Wend
    Else
        respuesta = ""
    End If
    
    rs.Close
    Set rs = Nothing
    
    telefonosMoviles = respuesta
End Function

'devuelve string separado por ; con todos los telefonos fijos para una persona
Public Function telefonosFijos(idPersona As Long) As String
    Dim strSql As String
    Dim respuesta As String
    Dim rs As ADODB.Recordset
    
    strSql = " SELECT fkPersona, telefono" & _
             " FROM T_Telefono" & _
             " WHERE (fkPersona=" & idPersona & ") AND fkTelefonoTipo=1;" 'fkTelefonoTipo=1(persona)
    
    Set rs = New ADODB.Recordset
    rs.Open strSql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    respuesta = ""
    If Not rs.EOF Then
        rs.MoveFirst
        While Not rs.EOF
            If Not IsNull(rs!telefono) Then
                If esFijo(rs!telefono) Then
                    If (respuesta = "") Then
                        respuesta = rs!telefono
                    Else
                        respuesta = respuesta & separadorMoviles & rs!telefono
                    End If
                End If
                rs.MoveNext
            End If
        Wend
    Else
        respuesta = ""
    End If
    
    rs.Close
    Set rs = Nothing
    
    telefonosFijos = respuesta
    
End Function
