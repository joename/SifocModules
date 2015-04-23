Attribute VB_Name = "G_Query"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  20/05/2011 - Actualización:  20/05/2011
'   Name:   modifyQuery
'   Desc:   creamos nueva consulta temporal Passthrough
'   Param:  cnStr as string
'           queryName as string
'           sql as string
'           delete as boolean, abre query y la elimina
'   Retur:  -
'---------------------------------------------------------------------------
Public Function createQuery(cnnStr As String, _
                            queryName As String, _
                            sql As String, _
                            Optional delete As Boolean = False)
    Dim dbs As dao.database
    Dim qdfPassThrough As QueryDef
    Dim qdfTemp As QueryDef
    Dim tmpQueryName As String
    
    tmpQueryName = queryName
    Set dbs = CurrentDb()
    
    If existsQuery(tmpQueryName) Then
        dbs.QueryDefs.delete (tmpQueryName)
    End If
    Set qdfPassThrough = dbs.CreateQueryDef(tmpQueryName)
    
    qdfPassThrough.Connect = cnnStr
    
    '"ODBC;DATABASE=sifoclocal;DSN=sifoclocal;OPTION=131082;PORT=0;SERVER=localhost;;"
    
    qdfPassThrough.sql = sql
    
    qdfPassThrough.ReturnsRecords = True
    
'    With dbs
'    Set qdfTemp = .CreateQueryDef("tmpTable", "SELECT telefono from t_telefono where fkPersona = 2000")
'    DoCmd.OpenQuery "tmpTable"
'    .QueryDefs.delete "tmpTable"
'    End With
    
    If delete = True Then
        DoCmd.OpenQuery tmpQueryName, acViewNormal, acReadOnly
        dbs.QueryDefs.delete tmpQueryName
    End If
    dbs.Close

End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  20/05/2011 - Actualización:  20/05/2011
'   Name:   existsQuery
'   Desc:   creamos nueva consulta temporal Passthrough
'   Param:  cnStr as string
'           queryName as string
'           sql as string
'   Retur:  -
'---------------------------------------------------------------------------
Public Function existsQuery(QryName As String) As Boolean
    Dim qd As QueryDef
    For Each qd In CurrentDb.QueryDefs
        If qd.name = QryName Then
            existsQuery = True
            Exit Function
        End If
    Next
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  20/05/2011 - Actualización:  20/05/2011
'   Name:   delTmpQueries
'   Desc:   creamos nueva consulta temporal Passthrough
'   Param:  -
'   Retur:  integer, num of queries deleted
'---------------------------------------------------------------------------
Public Function delTmpQueries() As Integer
    Dim qd As QueryDef
    Dim numDel As Integer
    
    numDel = 0
    
    For Each qd In CurrentDb.QueryDefs
        If Left(qd.name, 4) = "tmp_" Then
            numDel = numDel + 1
            CurrentDb.QueryDefs.delete (qd.name)
        End If
    Next
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  24/05/2011 - Actualización:  24/05/2011
'   Name:   QueryExists
'   Desc:   Checks to see whether the named table exists in the database
'           hlfUtils.TableExists
'   Param:
'   Retur:  True, if table found in current db, False if not found.
'---------------------------------------------------------------------------
Function QueryExists(queryName As String) As Boolean
    Dim strQueryNameCheck
On Error GoTo ErrorCode

    'try to assign tablename value
    strQueryNameCheck = CurrentDb.QueryDefs(queryName)
    
    'If no error and we get to this line, true
    QueryExists = True
    
ExitCode:
    On Error Resume Next
    Exit Function

ErrorCode:
    Select Case Err.Number
        Case 3265  'Item not found in this collection
            QueryExists = False
            Resume ExitCode
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.description, vbCritical, "hlfUtils.queryExists"
            Resume ExitCode
    End Select

End Function

Public Function updQuerySql(queryName As String, sql As String) As String
    Dim db As dao.database
    Dim qdf As dao.QueryDef

    G_Connection.setDatabase

    Set db = u_db
    Set qdf = db.QueryDefs(queryName)

    qdf.sql = sql

End Function

'-----------------------------------------------------------------------------------
'   Get query from access query
'-----------------------------------------------------------------------------------
Public Function getQuerySql(queryName As String) As String
    Dim db As dao.database
    Dim qdf As dao.QueryDef

    G_Connection.setDatabase

    Set db = u_db
    Set qdf = db.QueryDefs(queryName)

    getQuerySql = qdf.sql

End Function

'------------------------------------------------------------------------------------
'   Get Query from table
'------------------------------------------------------------------------------------
Public Function getQuery(id As Long) As String
    Dim str As String
    str = montarSQL(Nz(DLookup("strSelect", "t_query", "[id]=" & id), ""), _
                        Nz(DLookup("strFrom", "t_query", "[id]=" & id), ""), _
                        DLookup("strWhere", "t_query", "[id]=" & id), _
                        Nz(DLookup("strGroupby", "t_query", "[id]=" & id), ""), _
                        Nz(DLookup("strHaving", "t_query", "[id]=" & id), ""), _
                        Nz(DLookup("strOrderby", "t_query", "[id]=" & id), ""))
    getQuery = str
End Function

Public Function getQueryWhere(id As Long, strWhere As String) As String
    Dim str As String
Debug.Print str
    str = montarSQL(Nz(DLookup("strSelect", "t_query", "[id]=" & id), ""), _
                        Nz(DLookup("strFrom", "t_query", "[id]=" & id), ""), _
                        strWhere, _
                        Nz(DLookup("strGroupby", "t_query", "[id]=" & id), ""), _
                        Nz(DLookup("strHaving", "t_query", "[id]=" & id), ""), _
                        Nz(DLookup("strOrderby", "t_query", "[id]=" & id), ""))
    getQueryWhere = str
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  20/05/2011 - Actualización:  20/05/2011
'   Name:   delQuery
'   Desc:   creamos nueva consulta temporal Passthrough
'   Param:  -
'   Retur:  integer, num of queries deleted
'---------------------------------------------------------------------------
Public Function delQuery(queryName As String) As Integer
    Dim qd As QueryDef
    
    If QueryExists(queryName) Then
        CurrentDb.QueryDefs.delete (queryName)
    End If
    
End Function

Public Function tmpgquery()
    Dim strSql As String
    strSql = "select id, fecha from t_gestion order by fecha desc"
    
    CurrentDb.QueryDefs("___comor1").sql = strSql
    
    DoCmd.OpenQuery "___comor1"
End Function
