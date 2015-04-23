Attribute VB_Name = "WIFOC_DB"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Update: Jose Manuel Sanchez
'   Fecha:  06/04/2011 - Actualización: 06/04/2011
'   Name:   cambiaEstadoPersonaWeb
'   Desc:   Cambia el estado de la personaWeb
'   Param:  idWebPersona
'           estado, -1 transpasado, 0 web, 1 bloqueado (ya estaba en sifoc)
'   Retur:   0, Ok
'           -1, Ko
'---------------------------------------------------------------------------
Public Function cambiaEstadoWebPersona(idWebPersona As Long, _
                                       estado As Integer)
    Dim strSql As String
    
    strSql = " UPDATE t_webpersona" & _
             " SET estadoPersona=" & estado & _
             " WHERE id =" & idWebPersona & ";"
    CurrentDb.Execute strSql
    
End Function

Public Function linkWIFOC_personaTable()
    Dim tblConStr As String
    Dim tblDefStr As String
    Dim tblStr As String
    
    tblDefStr = "t_webpersona"
    tblStr = "t_persona"
    
    tblConStr = createConStr(, "serverifoc", "wifoc", , "wifoc")
Debug.Print tblConStr
    LinkTable tblStr, tblDefStr, "Driver={MySQL ODBC 5 Driver};Server=serverifoc;Port=3306;Database=wifoc;User=sifoc;Password=userifoc;Option=3;" 'DB_CONNECT 'tblConStr

End Function

Public Function unlinkWIFOC_personaTable()
    UnlinkTable "t_persona"
End Function


Public Function abcd()
    'DoCmd.TransferDatabase acLink, "ODBC Database", _
    "Server=serverifoc;Port=3306;Database=wifoc;Uid=sifoc;Pwd=userifoc;", _
    , acTable, "t_persona", "t_webpersona"

    DoCmd.TransferDatabase acLink, "Bases de datos ODBC", _
    DB_CONNECT, acTable, "t_persona", "t_webpersona"

    'DoCmd.TransferDatabase acLink, "ODBC Database", _
    "ODBC;DSN=wifoc;DATABASE=wifoc;", _
    acTable, "t_persona", "t_webpersona"  ', "a_sexo"
    
    'DoCmd.TransferDatabase acLink, DBType, _
    "ODBC;DSN=" & DSN & ";OPTION=131082;PWD=" & PWD & ";SERVER=" & SERVER & ";" _
    & "DATABASE=" & database & ";", acTable, tabla ', "a_sexo"

End Function

