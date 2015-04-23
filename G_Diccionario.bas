Attribute VB_Name = "G_Diccionario"
Option Explicit
Option Compare Database

'-----------------------------------------------------------------------------
Public Function tmpcuentaTags(str As String) As Integer
    'el caracter separador entre telefonos es el ";"
    Dim cadenas As Variant
    Dim i As Integer
    Dim counter As Integer
    Dim subcadena As Variant
    
    cadenas = Split(str, " ")
    i = 0
    counter = 0
    For Each subcadena In cadenas
        counter = counter + 1
    Next
    
    tmpcuentaTags = counter
    
End Function

Public Function tmpnewTag(fk As Integer, formacion As String)
    Dim str As String
    Dim i As Integer
    Dim n As Integer
    Dim tags As Variant
    Dim tag As String
    
    n = tmpcuentaTags(formacion)
    tags = Split(formacion, " ")
    
    For i = 0 To n - 1
        tag = tags(i)
        str = " INSERT INTO T_Diccionario ( fkGrupoformacion2, tag )" & _
              " SELECT '" & fk & "' AS fkGrupo2, '" & tag & "' AS newtag;"
             
'debugando str
        CurrentDb.Execute str
    Next i
    
End Function

Public Function tmpGetTags()
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim fk As Integer
    
    str = " SELECT T_InteresFormacion.fkGrupoFormacion2, T_InteresFormacion.formacion" & _
          " FROM T_InteresFormacion;"
    
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    rs.MoveFirst
    While Not rs.EOF
        tmpnewTag Nz(rs!fkGrupoFormacion2, 0), Replace(Nz(rs!formacion, ""), "'", ",")
        rs.MoveNext
    Wend
    
End Function

