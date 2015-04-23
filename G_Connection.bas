Attribute VB_Name = "G_Connection"
Option Explicit
Option Compare Database

Global u_db As dao.database

'Cadena de conexion:
'                                            "Type=ODBC Driver;" & _

Public Const DB_CONNECT_LOCAL As String = "Driver={MySQL ODBC 5.1 Driver};" & _
                                            "Server=localhost;" & _
                                            "Port=3306;" & _
                                            "Database=sifoc;" & _
                                            "User=root;" & _
                                            "Password=root;" & _
                                            "Option=3;"

Public Const DB_CONNECT As String = "Driver={MySQL ODBC 5.1 Driver};" & _
                                    "Server=serverifoc;" & _
                                    "Port=3306;" & _
                                    "Database=sifoc;" & _
                                    "User=sifoc;" & _
                                    "Password=userifoc;" & _
                                    "Option=3;"

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/01/2009 Actualización: 21/01/2009
'   Name:   connectionStr
'   Desc:   -
'   Param:  class as integer,
'   Retur:
'---------------------------------------------------------------------------
Public Function getSifocCnnStr(Optional Class As Integer = 0) As String
    Dim cnnStr As String

'"ODBC;DATABASE=sifoclocal;DSN=sifoclocal;OPTION=131082;PORT=0;SERVER=localhost;;"

    If Class = 1 Then
        cnnStr = getConnectionString("ODBC", "sifoc", "sifoc", "serverifoc", , , , 131082)
    Else 'test localhost
        cnnStr = getConnectionString("ODBC", "sifoclocal", "sifoclocal", "localhost", , , , 131082)
    End If
    getSifocCnnStr = cnnStr
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/01/2009 Actualización: 21/01/2009
'   Name:   isLocalConnection
'   Desc:   -
'   Param:
'   Retur:  true, if is a local connection (.mdb)
'           false, if is not local connection (<>mdb)
'---------------------------------------------------------------------------
Private Function isLocalConnection() As Boolean
    Dim conStr As String
    Dim cadenas
    
    conStr = CurrentProject.Connection.ConnectionString
    'cadenas = Split(conStr, ";")
    
    If InStr(conStr, ".accdb") = 0 Then
        isLocalConnection = False
    Else
        isLocalConnection = True
    End If
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/01/2009 Actualización: 21/01/2009
'   Name:   connectionString
'   Desc:   -
'   Param:
'   Retur:  connectionString,

'---------------------------------------------------------------------------
Public Function getMyConnectionString() As String
    Dim conStr As String
    Dim cadenas
        
    If isLocalConnection Then
        getMyConnectionString = DB_CONNECT_LOCAL
    Else
        getMyConnectionString = DB_CONNECT
    End If
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/01/2009 Actualización: 21/01/2009
'   Name:   connectionString
'   Desc:   -
'   Param:
'   Retur:  connectionString,

'---------------------------------------------------------------------------
Public Function getMyDsn() As String
    Dim conStr As String
    Dim cadenas
        
    If isLocalConnection Then
        getMyDsn = "sifoclocal"
    Else
        getMyDsn = "sifoc"
    End If
    
End Function
'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  8/07/2014 Actualización: 8/07/2014
'   Name:   setDatabase
'   Desc:   -
'   Param:
'   Retur:
'---------------------------------------------------------------------------
Public Static Function setDatabase()

    If (u_db Is Nothing) Then
        Set u_db = CurrentDb()
    End If

End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/01/2009 Actualización: 21/01/2009
'   Name:   connectionStr
'   Desc:   -
'   Param:
'   Retur:
'---------------------------------------------------------------------------
Public Function getConnectionString(Optional driver As String = "DRIVER", _
                                    Optional database As String = "DATABASE", _
                                    Optional dsn As String = "DSN", _
                                    Optional server As String = "SERVER", _
                                    Optional port As String = "0", _
                                    Optional user As String = "USER", _
                                    Optional password As String = "PASSWORD", _
                                    Optional opt As String = "0") As String
    Dim cnnStr As String
'"ODBC;DATABASE=sifoclocal;DSN=sifoclocal;OPTION=131082;PORT=0;SERVER=localhost;;"
    cnnStr = IIf(driver = "DRIVER", "", driver & ";") & _
            IIf(database = "DATABASE", "", "Database=" & database & ";") & _
            IIf(dsn = "DSN", "", "Dsn=" & dsn & ";") & _
            IIf(user = "USER", "", "User=" & user & ";") & _
            IIf(password = "PASSWORD", "", "Password=" & password & ";") & _
            IIf(opt = "OPTION", "", "Option=" & opt & ";") & _
            IIf(port = 0, "", "Port=" & port & ";") & _
            IIf(server = "SERVER", "", "Server=" & server & ";") & _
            ";"
    getConnectionString = cnnStr
            
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/01/2009 Actualización: 21/01/2009
'   Name:   createLinkTable
'   Desc:   -
'   Param:  - mirar con TableDef!!!!!!!!!!! parecido a QueryDef
'   Retur:  -
'---------------------------------------------------------------------------
Public Function createLinkTable(tabla As String, _
                                Optional DBType As String = "ODBC Database", _
                                Optional server As String = "SERVER", _
                                Optional dsn As String = "DSN", _
                                Optional port As String = "3306", _
                                Optional database As String = "DATABASE", _
                                Optional uid As String = "USER", _
                                Optional pwd As String = "PASSWORD", _
                                Optional opt As Integer = 0) As String
    Dim cnnStr As String

    cnnStr = IIf(server = "SERVER", "", "Server=" & server & ";") & _
            "Port=" & port & ";" & _
            "Database=" & database & ";" & _
            "UID=" & uid & ";" & _
            "Password=" & pwd & ";" & _
            IIf(opt = 0, "", "Option=" & opt & ";")
   
    'DoCmd.TransferDatabase acLink, DBType, _
    "ODBC;DSN=" & DSN & ";OPTION=131082;PWD=" & PWD & ";SERVER=" & SERVER & ";" _
    & "DATABASE=" & database & ";", acTable, tabla ', "a_sexo"
    
'    Dim tdfLinked As ADODB.TableDef
'    Dim db As ADODB.database
'
'    Set dblocal = CurrentDb
'    Set tdfLinked = dblocal.CreateTableDef("t_weeeeeb")
'    Set tdfLinked.Connect = "ODBC;DSN=sifoclocal;DATABASE=sifoclocal;"
'    Set tdfLinked.SourceTableName = "t_persona"
'    dblocal.TableDefs.Append tdfLinked
'    dblocal.TableDefs.Refresh
    
    

End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/01/2009 Actualización: 21/01/2009
'   Name:   createLinkTableStr
'   Desc:   -
'   Param:  -
'   Retur:  -
'---------------------------------------------------------------------------
Public Function createConStr(Optional DBType As String = "ODBC Database", _
                             Optional server As String = "SERVER", _
                             Optional dsn As String = "DSN", _
                             Optional port As String = "3306", _
                             Optional database As String = "DATABASE", _
                             Optional uid As String = "USER", _
                             Optional pwd As String = "PASSWORD", _
                             Optional opt As Integer = 0) As String
    Dim cnnStr As String

    cnnStr = IIf(server = "SERVER", "", "Server=" & server & ";") & _
            "Port=" & port & ";" & _
            IIf(database = "DATABASE", "", "Database=" & database & ";") & _
            IIf(dsn = "DNS", "", "DSN=" & dsn & ";") & _
            IIf(uid = "USER", "", "UID=" & uid & ";") & _
            IIf(pwd = "PASSWORD", "", "Password=" & pwd & ";") & _
            IIf(opt = 0, "", "Option=" & opt & ";")
    
   
    createConStr = cnnStr

End Function

Public Function LinkTable(strTable As String, _
                          strDefTable As String, _
                          strDb As String)

    Const LT_LINKEDALREADY As Integer = 3012
    Dim db As dao.database
    Dim tdf As TableDef

On Error GoTo Err_LinkTable
    
    Set db = CurrentDb

   '-- Create a new TableDef then link the external
   '-- table to it
   Set tdf = db.CreateTableDef(strDefTable)
   tdf.Connect = strDb
   tdf.SourceTableName = strTable

   '-- Add this TableDef to the current database.
   db.TableDefs.Append tdf

Exit_LinkTable:
   Exit Function

Err_LinkTable:
   If Err.Number = LT_LINKEDALREADY Then
     '-- Do nothing - the table's linked in already
     Resume Exit_LinkTable
   Else
     '-- Put some code here to handle this error
     Debug.Print Err.description
   End If
End Function

Public Function UnlinkTable(strTable As String)
    Dim db As dao.database
    Set db = CurrentDb
On Error GoTo UnlinkTable_Err
    
    If db.TableDefs(strTable).Connect = "" Then
        UnlinkTable = False
        GoTo UnlinkTable_Exit
    End If
    db.TableDefs.delete (strTable)
    UnlinkTable = True
    
UnlinkTable_Exit:
    Set db = Nothing
    Exit Function
UnlinkTable_Err:
    UnlinkTable = False
    GoTo UnlinkTable_Exit
End Function

Public Function passthrough_query_test()
    Dim dbs As dao.database
    Dim qdfPassThrough As QueryDef
    Dim qdfTemp As QueryDef
    
    Dim strSql As String
    Dim strCnn As String
    
    Set dbs = CurrentDb()
    
    Set qdfPassThrough = dbs.CreateQueryDef("ShortCodes")
    
    strCnn = "ODBC;DATABASE=sifoc;DSN=sifoc;OPTION=131082;PORT=3306;SERVER=serverifoc;"
    
    qdfPassThrough.Connect = strCnn
    '"ODBC;DSN=OMPUBLIC;UID=OMPUBLIC;PWD=ompublic;DBQ=PROD ;DBA=W;APA=T;EXC=F;FEN=T;QTO=T;FRC=10;FDL=10;LOB=T;RST=T;BTD=F;BNF=F;BAM=IfAllSuccessful;NUM=NLS;DPM=F;MTS=T;MDI=F;CSR=F ;FWC=F;FBS=64000;TLO=O;"
    
    strSql = " CREATE TEMPORARY TABLE temporal (" & _
             " id INTEGER UNSIGNED NOT NULL AUTO_INCREMENT," & _
             " numero INTEGER UNSIGNED NOT NULL," & _
             " texto VARCHAR(45) NOT NULL," & _
             " PRIMARY KEY (`id`)" & _
             " )" & _
             " ENGINE = InnoDB;"
    
    qdfPassThrough.sql = strSql
    
    qdfPassThrough.ReturnsRecords = True
    
    With dbs
    Set qdfTemp = .CreateQueryDef("tmpTable", strSql)
    DoCmd.OpenQuery "tmpTable"
    .QueryDefs.delete "tmpTable"
    End With
    
    dbs.QueryDefs.delete "ShortCodes"
    dbs.Close

End Function

Public Function temporaltable()
    Dim sSql As String
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim str As String
    Dim strCnn As String
    
    'Set dbs = CurrentDb()
    'OPTION=131082;PORT=3306;SERVER=serverifoc;
    strCnn = "Driver={MySQL ODBC 5.1 Driver};ODBC;DSN=sifoclocal5;DATABASE=sifoclocal;"
    
    str = " CREATE TEMPORARY TABLE IF NOT EXISTS temporal (" & _
          " id INTEGER UNSIGNED NOT NULL AUTO_INCREMENT," & _
          " numero INTEGER UNSIGNED NOT NULL," & _
          " texto VARCHAR(45) NOT NULL," & _
          " PRIMARY KEY (`id`)" & _
          " )" & _
          " ENGINE = InnoDB;"
    
    Set cn = New ADODB.Connection
    cn.Open strCnn
    '"ODBC;DSN=sifoclocal5;Port=3306;UID=sifoclocal;" & _
            "PWD=usersifoclocal;SERVER=localhost;" 'from your ODBC setup
    
    Set rs = New ADODB.Recordset
    
    rs.Open str, strCnn, adOpenDynamic, adLockOptimistic 'this line does the sql statement
 

    
    'createLinkTable "temporal", "ODBC Database", "localhost", "sifoclocal5", , "sifoclocal", "root", "root"
    
    DoCmd.TransferDatabase acLink, "ODBC", strCnn, acTable, "temporal", "temporal"
 
 'DoCmd.TransferDatabase acLink, "ODBC", _
    "ODBC;DSN=sifoclocal5;UID=root;PWD=root;LANGUAGE=us_english;" _
    & "DATABASE=sifoclocal", acTable, "temporal", "temporal"
   
    
    str = "insert into temporal (numero, texto) values (1,'hola');"
    CurrentDb.Execute str
    
    DoCmd.DeleteObject acTable, "temporal"
    
    'cerramos despues de too
    cn.Close
End Function

Public Function gcon()
    Set u_db = CurrentDb()
    SIFOC_IfocTarea.insertIfocTarea "hola", "alñskdfj", 14
    
End Function
