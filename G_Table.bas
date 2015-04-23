Attribute VB_Name = "G_Table"
Option Explicit
Option Compare Database

Public Function makeTable(tableName As String, _
                          rst As ADODB.Recordset) As Boolean
   'References: Microsoft Access 11.0 Object Library, Microsoft DAO 3.6 Object Library
   'Set references by Clicking Tools and Then References in the Code View window
   '' Creates an empty table (TableName) with a Field (FieldName).
   ' Accepts ' TableName: Name of table to create
   ' FieldName: Name of the field to create in the table
   ' Returns True on success, false otherwise.
   'USAGE: MakeTable "TABLENAME", "FIELDNAME"
    Dim db As dao.database ' DAO Vars
    Dim strSql As String, fieldsName As String
    Dim x As Integer
    
On Error GoTo ErrHandler 'If there is an error capture the error.

    'Construye el String a utilizar para crear la Tabla
    rst.MoveFirst
    'Por si no devuelve ningún registro
    If rst.EOF Or rst.BOF Then
        makeTable = True
        Set db = Nothing
        Exit Function
    End If
    
    'Eliminamos tabla antes de crear por si ya existe
    delTable tableName
    
    strSql = "SELECT "
    fieldsName = " ("
    For x = 0 To rst.fields.count - 1
        strSql = strSql & IIf(x = 0, "", ", ")
        fieldsName = fieldsName & IIf(x = 0, "", ", ")
'Debug.Print rst.fields(X).Type & " - " & rst.fields(X).name
        Select Case rst.fields(x).Type
            'Tipos numéricos: (3 Integer)(20 BigInt)(128 Binary)(11 Boolean)(136 Chapter)(6 Currency)(14 Decimal)
            '                 (5 Double)(64 FileTime)(72 GUID)(205 LongVarBinary)(131 Numeric)(4 Single)
            '                 (2 SmallInt)(16 TinyInt)(21 UnsignedBigInt)(19 UnsignedInt)(18 UnsignedSmallInt)
            '                 (17 UnsignedTinyInt)
            Case 3, 20, 128, 11, 136, 6, 14, 5, 64, 72, 205, 131, 4, 2, 16, 21, 19, 18, 17, 139
                strSql = strSql & rst.fields(x) & " AS " & rst.fields(x).name
                fieldsName = fieldsName & rst.fields(x).name
            'Tipos alfanuméricos: (200 VarChar)(135 DBTimeStamp)(8 BSTR)(129 Char)(7 Date)(133 DBDate)(134 DBTime)
            '                     (201 LongVarChar)(202 VarWChar, 50)(203 LongVarWChar)(130 WChar)
            Case 200, 135, 8, 129, 7, 133, 134, 201, 202, 203, 130
                strSql = strSql & "'" & rst.fields(x) & "' AS " & rst.fields(x).name
                fieldsName = fieldsName & rst.fields(x).name
            'Tipos que no deben utilizarse, y los anula directamente
            '    (0 Empty)(10 Error)(9 IDispatch)(138 PropVariant)(132 UserDefined)
            '    (204 VarBinary)(12 Variant)(139 VarNumeric)
            Case Else
                strSql = strSql & "''" & " AS " & rst.fields(x).name
                fieldsName = fieldsName & rst.fields(x).name
        End Select
    Next
    fieldsName = fieldsName & ") "
    strSql = strSql & " INTO " & tableName & ";"
    
    'Aquí crea la tabla en la BBDD CurrentDB, y además inserta el primer registro
    Set db = CurrentDb()
    db.Execute strSql
    
    'Aquí inserta el resto de registros en la tabla ya creada
    rst.MoveNext
    Do While Not rst.EOF
        strSql = "INSERT INTO " & tableName & fieldsName & " VALUES ("
        For x = 0 To rst.fields.count - 1
            strSql = strSql & IIf(x = 0, "", ", ")
            Select Case rst.fields(x).Type
                'Tipos numéricos: (3 Integer)(20 BigInt)(128 Binary)(11 Boolean)(136 Chapter)(6 Currency)(14 Decimal)
                '                 (5 Double)(64 FileTime)(72 GUID)(205 LongVarBinary)(131 Numeric)(4 Single)
                '                 (2 SmallInt)(16 TinyInt)(21 UnsignedBigInt)(19 UnsignedInt)(18 UnsignedSmallInt)
                '                 (17 UnsignedTinyInt)
                Case 3, 20, 128, 11, 136, 6, 14, 5, 64, 72, 205, 131, 4, 2, 16, 21, 19, 18, 17, 139
                    strSql = strSql & rst.fields(x)
                'Tipos alfanuméricos: (200 VarChar)(135 DBTimeStamp)(8 BSTR)(129 Char)(7 Date)(133 DBDate)(134 DBTime)
                '                     (201 LongVarChar)(202 VarWChar, 50)(203 LongVarWChar)(130 WChar)
                Case 200, 135, 8, 129, 7, 133, 134, 201, 202, 203, 130
                    strSql = strSql & Chr(34) & rst.fields(x) & Chr(34)
                'Tipos que no deben utilizarse, y los anula directamente
                '    (0 Empty)(10 Error)(9 IDispatch)(138 PropVariant)(132 UserDefined)
                '    (204 VarBinary)(12 Variant)(139 VarNumeric)
                Case Else
                    strSql = strSql & Chr(34) & Chr(34)
            End Select
        Next
        strSql = strSql & ");"
        db.Execute strSql
        rst.MoveNext
    Loop
    makeTable = True

ExitHere:
    Set db = Nothing
    Exit Function
   
ErrHandler:
    If rst.BOF Or rst.EOF Then
        MsgBox "No hay registros a devolver"
        Set db = Nothing
        Exit Function
    End If
    
    MsgBox "Error " & Err.Number & vbCrLf & Err.description, _
                vbOKOnly Or vbCritical, "Make Table Incomplete"
Debug.Print strSql
    makeTable = False
    Resume ExitHere
End Function

Public Function TableDefX(nameTbl As String, _
                          fields As String, _
                          dataTypes As String) As Integer

   Dim dbs As dao.database
   Dim tdfNew As TableDef
   Dim tdfLoop As TableDef
   Dim prpLoop As Property

   Set dbs = CurrentDb()

   ' Create new TableDef object, append Field objects
   ' to its Fields collection, and append TableDef
   ' object to the TableDefs collection of the
   ' Database object.
   Set tdfNew = dbs.CreateTableDef(nameTbl)
   tdfNew.fields.Append tdfNew.CreateField("Date", dbDate)
   dbs.TableDefs.Append tdfNew

   With dbs
      Debug.Print .TableDefs.count & _
         " TableDefs in " & .name

      ' Enumerate TableDefs collection.
      For Each tdfLoop In .TableDefs
         Debug.Print "  " & tdfLoop.name
      Next tdfLoop

      With tdfNew
         Debug.Print "Properties of " & .name

         ' Enumerate Properties collection of new
         ' TableDef object, only printing properties
         ' with non-empty values.
         For Each prpLoop In .Properties
            Debug.Print "  " & prpLoop.name & " - " & _
               IIf(prpLoop = "", "[empty]", prpLoop)
         Next prpLoop

      End With

      ' Delete new TableDef since this is a
      ' demonstration.
      .TableDefs.delete tdfNew.name
      .Close
   End With

End Function

Public Function createTableDefX(nameTbl As String, _
                                fields As String, _
                                dataTypes As String) As Integer
    Dim dbs As dao.database
    Dim tdfNew As TableDef
    Dim prpLoop As Property
    
    Dim argsFields As Variant
    Dim argsTypes As Variant
    Dim num As Integer
    Dim i As Integer
    
    Set dbs = CurrentDb()
    
    ' Delete table if exists
    If TableExists(nameTbl) Then
        delTable nameTbl
    End If
    
    ' Create a new TableDef object.
    Set tdfNew = dbs.CreateTableDef(nameTbl)

    argsFields = Split(fields, "##")
    argsTypes = Split(dataTypes, "##")

    num = UBound(argsFields)

    For i = 0 To num Step 1
        With tdfNew
           ' Create fields and append them to the new TableDef
           ' object. This must be done before appending the
           ' TableDef object to the TableDefs collection of the
           ' Northwind database.
           .fields.Append .CreateField(argsFields(i), myVarType(CStr(argsTypes(i)))) 'argsTypes(i))
        
'           Debug.Print "Properties of new TableDef object " & _
'              "before appending to collection:"
'
'           ' Enumerate Properties collection of new TableDef
'           ' object.
'           For Each prpLoop In .Properties
'              On Error Resume Next
'              If prpLoop <> "" Then Debug.Print "  " & _
'                prpLoop.name & " = " & prpLoop
'              On Error GoTo 0
'           Next prpLoop
        
           ' Append the new TableDef object to the Northwind
           ' database.
        
        End With
    Next i
    
    dbs.TableDefs.Append tdfNew

   'dbs.Close
   
End Function

'Public Function delTable(nameTbl As String)
'    Dim dbs As database
'
'    Set dbs = CurrentDb()
'
'    ' Delete table if exists
'    If TableExists(nameTbl) Then
'        dbs.TableDefs.delete nameTbl
'    End If
'
'    Set dbs = Nothing
'    'dbs.Close
'End Function

Function delTable(ByVal tableName As String) As Boolean
   'References: Microsoft Access 11.0 Object Library, Microsoft DAO 3.6 Object Library
   'Set references by Clicking Tools and Then References in the Code View window
   ' Deletes a Table with a given TableName that exists in a database
   ' Accepts
   ' tableName: Name of table
   ' Returns True on success, false otherwise
   'USAGE: DeleteTable "TABLENAME"

On Error GoTo ErrHandler

    CurrentDb.TableDefs.delete (tableName)
    
   'If no errors return true
   delTable = True

ExitHere:
   Exit Function

ErrHandler:
    '3265= Tabla no existe en la colección
    If Err.Number <> 3265 Then
        'There is an error return false
        delTable = False
        With Err
            MsgBox "Error " & .Number & vbCrLf & .description, _
                vbOKOnly Or vbCritical, "DeleteTable"
        End With
        Resume ExitHere
    Else
        delTable = True
    End If

End Function

'---------------------------------------------------------------------------
'   Autor:  Copyright by Heather L. Floyd - Floyd Innovations - www.floydinnovations.com
'   Update: Jose Manuel Sanchez
'   Fecha:  08-01-2005 - Actualización:  21/05/2011
'   Name:   TableExists
'   Desc:   Checks to see whether the named table exists in the database
'           hlfUtils.TableExists
'   Param:
'   Retur:  True, if table found in current db, False if not found.
'---------------------------------------------------------------------------
Public Function TableExists(tableName As String) As Boolean
    Dim strTableNameCheck
On Error GoTo ErrorCode

    'try to assign tablename value
    strTableNameCheck = CurrentDb.TableDefs(tableName)
    
    'If no error and we get to this line, true
    TableExists = True
    
ExitCode:
    On Error Resume Next
    Exit Function

ErrorCode:
    
    Select Case Err.Number
        Case 3265  'Item not found in this collection
            TableExists = False
            Resume ExitCode
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.description, vbCritical, "hlfUtils.TableExists"
            'Debug.Print "Error " & Err.number & ": " & Err.Description & "hlfUtils.TableExists"
            Resume ExitCode
    End Select

End Function

Public Function createLocalTable(tableName As String, _
                                 sql As String) As Integer
    Const queryName = "tmp_query"
    'v_usuariosTR_ultimas (2)
    createQuery getSifocCnnStr(1), queryName, sql, False
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "select * from " & queryName
    rs.Open sql, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    makeTable tableName, rs

    rs.Close
    Set rs = Nothing
    
    delQuery queryName
End Function

Public Function myVarType(name As String) As Integer
    Select Case name
        Case "vbEmpty": myVarType = 0
        Case "vbNull": myVarType = 1
        Case "vbInteger": myVarType = 2
        Case "vbLong": myVarType = 3
        Case "vbSingle": myVarType = 4
        Case "vbDouble": myVarType = 5
        Case "vbCurrency": myVarType = 6
        Case "vbDate": myVarType = 7
        Case "vbString": myVarType = 8
        Case "vbObject": myVarType = 9
        Case "vbError": myVarType = 10
        Case "vbBoolean": myVarType = 11
        Case "vbVariant": myVarType = 12
        Case "vbDataObject": myVarType = 13
        Case "vbByte": myVarType = 17
        Case "vbUserDefinedType": myVarType = 36
        Case "vbArray": myVarType = 8192
        Case Else: myVarType = 0
    End Select
End Function

