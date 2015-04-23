Attribute VB_Name = "SIFOC_Control"
Option Explicit
Option Compare Database

Public Function initForm(formName As String, idIfocUsuario As Long)
    hitForm formName, idIfocUsuario
        
    '---Comprobación tareas pendientes por demandante
    checkPendingTasks (formName)
    
End Function

'-------------------------------------------------------------------------------------
'       Actualiza Tabla Formularios
'-------------------------------------------------------------------------------------
Public Function hitForm(formName As String, ifocUsuario As Long) As Integer
    Dim str As String
    Dim Status As Integer
    
    If Not existFormDb(formName, ifocUsuario) Then ' crea nuevo registro con el formulario
        str = " INSERT INTO sysforms ( form, fkIfocUsuario, counter , updDate)" & _
              " VALUES ('" & formName & "', " & ifocUsuario & ", 1, now());"
    Else ' actualiza contador de formulario
        str = " UPDATE sysforms" & _
              " SET counter = counter+1, updDate = Now()" & _
              " WHERE (form='" & formName & "') AND (fkIfocUsuario=" & ifocUsuario & ");"
    End If

    CurrentDb.Execute str
    
    If (Status = -1) Then
        hitForm = -1
    Else
        hitForm = 0
    End If
End Function

'-------------------------------------------------------------------------------------
'       Actualiza Tabla Formularios
'-------------------------------------------------------------------------------------
Public Function hitQuery(queryName As String, ifocUsuario As Long) As Integer
    Dim str As String
    Dim Status As Integer
    
    If Not existQueryDb(queryName, ifocUsuario) Then ' crea nuevo registro con el formulario
        str = " INSERT INTO sysforms ( query, fkIfocUsuario, counter ) VALUES ('" & queryName & "', " & ifocUsuario & ", 1);"
    Else ' actualiza contador de formulario
        str = " UPDATE sysforms" & _
              " SET counter = counter+1" & _
              " WHERE (form='" & queryName & "') AND (fkIfocUsuario=" & ifocUsuario & ");"
    End If

    CurrentDb.Execute str
    
    If (Status = -1) Then
        hitQuery = -1
    Else
        hitQuery = 0
    End If
End Function

Public Function existFormDb(Form As String, ifocUsuario As Long) As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    
    str = " SELECT form, fkIfocusuario" & _
          " FROM sysForms" & _
          " WHERE form = '" & Form & "' AND fkIfocUsuario=" & ifocUsuario & ";"

    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    If (rs.EOF) Then
        existFormDb = False
    Else
        existFormDb = True
    End If

End Function

Public Function existQueryDb(query As String, ifocUsuario As Long) As Boolean
    Dim str As String
    Dim rs As ADODB.Recordset
    
    str = " SELECT query, fkIfocusuario" & _
          " FROM sysForms" & _
          " WHERE query = '" & query & "' AND fkIfocUsuario=" & ifocUsuario & ";"

    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly

    If (rs.EOF) Then
        existQueryDb = False
    Else
        existQueryDb = True
    End If

End Function

