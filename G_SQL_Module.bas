Attribute VB_Name = "G_SQL_Module"
Option Explicit
Option Compare Database

Const espacio As String = " "

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2008
'   Name:   addConditionWhere
'   Desc:   Esta funcion permite añadir condiciones al where
'   Param:  str(String), es el string de condicines anteriores
'           condicion(String), es la condicion del where a añadir
'   Retur:  devuelve String con la condición añadida con AND
'---------------------------------------------------------------------------
Public Function addConditionWhere(str As String, condition As String) As String
    Dim strWhere As String
    
    If Len(str) > 0 Then
        strWhere = str & " AND (" & condition & ")"
    Else
        strWhere = " (" & condition & ")"
    End If
    
    addConditionWhere = strWhere
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2008
'   Name:   addConditionHaving
'   Desc:   Esta funcion permite añadir condiciones al having
'   Param:  str(String), es el string de condicines anteriores
'           condicion(String), es la condicion del having a añadir
'   Retur:  devuelve String con la condición añadida con AND
'----------------------------------------------------------------------------
Public Function addConditionHaving(str As String, condition As String) As String
    Dim strHaving As String
    
    If Len(str) > 0 Then
        strHaving = str & " AND (" & condition & ")"
    Else
        strHaving = " (" & condition & ")"
    End If
    
    addConditionHaving = strHaving
End Function


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2008
'   Name:   addConditionGroupBy
'   Desc:   Esta funcion permite añadir condiciones al groupby
'   Param:  str(String), es el string de condicines anteriores
'           condicion(String), es la condicion del groupby a añadir
'   Retur:  devuelve String con la condición añadida con AND
'----------------------------------------------------------------------------
Public Function addConditionGroupBy(str As String, field As String) As String
    Dim strGroup As String
    
    If Len(str) > 0 Then
        strGroup = str & ", " & field
    Else
        strGroup = field
    End If
    
    addConditionGroupBy = strGroup
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2008 - actualización: 21/1/2008
'   Name:   montarSQL
'   Desc:   Monta SQL con las sentencias pasadas por parámetro
'   Param:  strSelect, string del select sin palabra clave SELECT
'           strFrom, string del from sin la palabra clave FROM
'           strWhere, string del where sin la palabra clave WHERE
'           strGroupBy, string del groupBy sin la palabra clave GROUP BY
'           strHaving, string del having sin la palabra clave HAVING
'           strOrderBy, string del orderBy sin la palabra clave ORDER BY
'   Retur:  devuelve String con el sql formado con los parámetros pasados
'----------------------------------------------------------------------------
Public Function montarSQL(strselect As String, _
                          strFrom As String, _
                          Optional strWhere As String = "", _
                          Optional strGroupBy As String = "", _
                          Optional strHaving As String = "", _
                          Optional strOrder As String = "") As String
    Dim strSql As String
    
    'Montamos el select
    strSql = " SELECT" & espacio & strselect & _
             " FROM" & espacio & strFrom
             
    If Len(strWhere) > 0 Then
         strSql = strSql & " WHERE" & espacio & strWhere
    End If
    If Len(strGroupBy) > 0 Then
         strSql = strSql & " GROUP BY" & espacio & strGroupBy
    End If
    If Len(strHaving) > 0 Then
         strSql = strSql & " HAVING" & espacio & strHaving
    End If
    If Len(strOrder) > 0 Then
         strSql = strSql & " ORDER BY" & espacio & strOrder
    End If
    'strSQL = strSQL '& ";"
        
    montarSQL = strSql
        
End Function
