Attribute VB_Name = "WIFOC_Desplegables"
Option Explicit
Option Compare Database


Public Function paises()
    Dim fs, a, reg

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("c:\paises.txt", True)
    
    Dim rs As ADODB.Recordset
    Dim str As String
    
    str = " SELECT id, pais FROM a_pais ORDER BY id"
    
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rs.EOF Then
        rs.MoveFirst
    End If
    Do While Not rs.EOF
        a.WriteLine (rs!id & "|" & rs!pais)
        rs.MoveNext
    Loop
    
    a.Close

End Function

Public Function tipoAyuda()
    Dim fs, a, reg

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("c:\tipoayuda.txt", True)
    
    Dim rs As ADODB.Recordset
    Dim str As String
    
    str = " SELECT a_prestacionestipo.id, [entidad] & ' - ' & [tipoprestacion] AS Prestacion" & _
          " FROM a_prestacionesentidades INNER JOIN a_prestacionestipo ON a_prestacionesentidades.id = a_prestacionestipo.fkprestacionentidad;"
    
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If rs.EOF Then
        rs.MoveFirst
    End If
    Do While Not rs.EOF
        a.WriteLine (rs!id & "|" & rs!PRESTACION)
        rs.MoveNext
    Loop
    
    a.Close

End Function
