Attribute VB_Name = "SIFOC_Email"
Option Explicit
Option Compare Database


Public Function hasPersonalMail(idPersona As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim withMail As Boolean
    
    withMail = True
    
    sql = " select fkPersona, email from t_email where fkPersona = " & idPersona & " and fkEmailTipo = 1"
    Set rs = New ADODB.Recordset
    
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If (rs.EOF) Then
        withMail = False
    End If
    
    hasPersonalMail = withMail
End Function

'Public Function updTipoMail() As Integer
'    Dim rs As ADODB.Recordset
'    Dim sql As String
'
'On Error GoTo TratarError
'
'    sql = "select id, fkEmailTipo, fkPersona from t_email where not isnull(fkPersona) order by id ASC, timestamp DESC"
'
'    Set rs = New ADODB.Recordset
'    rs.Open sql, CurrentProject.Connection, adOpenDynamic, adLockOptimistic
'
'    If Not rs.EOF Then
'        rs.MoveFirst
'
'        Dim i As Integer
'        Dim idAnterior As Long
'        Dim idActual As Long
'        Dim counter As Long
'
'        idAnterior = 0
'        counter = 0
'        While Not rs.EOF
'            idActual = rs!fkPersona
'            counter = counter + 1
'
'            If (idActual <> idAnterior) Then
'                i = 1
'            ElseIf (idActual = idAnterior) Then
'                If i = 2 Then
'                    rs!fkEmailTipo = 2
'                End If
'                If i = 3 Then
'                    rs!fkEmailTipo = 3
'                End If
'            End If
'
'            idAnterior = idActual
'            rs.MoveNext
'            i = i + 1
'        Wend
'
'    End If
'
'    rs.Close
'    Set rs = Nothing
'
'    updTipoMail = counter
'
'SalirTratarError:
'    Exit Function
'TratarError:
'    MsgBox "Error openform(a):" & Err.description
'    Resume
'End Function


