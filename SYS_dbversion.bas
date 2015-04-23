Attribute VB_Name = "SYS_dbversion"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/8/2009
'   Name:   dbVersion
'   Desc:   Versión de la base de datos
'   Param:  -
'   Retur:  string, con la versión de la base de datos
'---------------------------------------------------------------------------
Public Function dbVersion() As String
    'Actualizacion de la base de datos
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim fechaV As Date
    
    fechaV = Format(U_applicationDate, "mm/dd/yyyy hh:mm:ss")
    
    str = " SELECT sysdbversion.version as version" & _
          " FROM sysdbversion" & _
          " WHERE fecha < #" & fechaV & "#" & _
          " ORDER BY fecha DESC, id DESC;"

    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
   
    If Not rs.EOF Then
        rs.MoveFirst
        dbVersion = "v " & rs!Version
    Else
        dbVersion = ""
    End If
    
    rs.Close
    Set rs = Nothing
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/8/2009
'   Name:   esUltimaVersion
'   Desc:   Nos indica si es la última versión de la base de datos
'   Param:  -
'   Retur:  string, con la versión de la base de datos
'---------------------------------------------------------------------------
Public Function esDBUltimaVersion() As Boolean
    'Actualizacion de la base de datos
    Dim rs As ADODB.Recordset
    Dim str As String
    
    str = " SELECT Max(sysdbversion.fecha) AS fecha" & _
          " FROM sysdbversion;"
          
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        rs.MoveFirst
        If DateDiff("n", rs!fecha, U_applicationDate) > 0 Then
            esDBUltimaVersion = True
        Else
            esDBUltimaVersion = False
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Function

