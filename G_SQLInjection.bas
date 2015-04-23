Attribute VB_Name = "G_SQLInjection"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/07/2009
'   Name:   filterSQL
'   Desc:   Filtra palabra palabras clave sql injection
'   Param:  str, string a filtrar
'   Retur:  devuelve str filtrado
'---------------------------------------------------------------------------
Public Function filterSQL(str As String) As String
    Dim palabrasClave()
    Dim intPosition As Integer
    Dim filteredStr As String
    
    palabrasClave = Array("SELECT", "FROM", "UNION", "INSERT INTO", "DROP", _
                          "DATABASE", "CUBE", "FUNCTION", "VIEW", "TRANSACTION", _
                          "INDEX", "PROCEDURE", "TABLE", "TRIGGER", "PERMISSIONS", _
                          "ALTER", "CREATE", "BACKUP", "DUMP", "DENY", _
                          "KILL", "TRIGGER", "CALL", "CONNECT", "CURRENT", _
                          "SET", "USER", "begin", "end", "declare", _
                          "#13", "#10", "#20", "#13", _
                          """", "'", "'", "--", ";")
    
    filteredStr = str
    For intPosition = LBound(palabrasClave) To UBound(palabrasClave) Step 1
        filteredStr = Replace(filteredStr, palabrasClave(intPosition), "@")
    Next intPosition
    
    filterSQL = filteredStr
    
End Function
