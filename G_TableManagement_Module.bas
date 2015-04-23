Attribute VB_Name = "G_TableManagement_Module"
Option Explicit
Option Compare Database

'Formato de la cadena strPares, caracter separacion "##"
' Primer argumento, numero de campos de la cadena
' Siguientes args,  "##nombreCampo@@valorCampo"

Public Const separador As String = "##"
Public Const separador1 As String = "@@"
Public Const primaryKey As String = "id"

'----------------------------------------------------------------------------------------------
'       Funciones de alta, baja, modificacion y consulta de registro en tabla
'----------------------------------------------------------------------------------------------
Public Function altaRegistro(tabla As String, strPares As String, Optional idstr As String = "") As Long
'On Error GoTo TratarError
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim num As Integer
    Dim id As Long
    Dim i As Integer

    num = numeroCamposString(strPares)

    If (num > 0) Then
        'Abrimos recordset
        str = tabla
        Set rs = New ADODB.Recordset
        rs.Open str, CurrentProject.Connection, adOpenDynamic, adLockOptimistic, adCmdTable
        
        id = 0
        rs.AddNew
        For i = 1 To num
            'debugando nombreCampoString(strPares, i) & " , " & valorCampoString(strPares, i)
            rs.fields(nombreCampoString(strPares, i)) = valorCampoString(strPares, i)
        Next i
'debugando strPares
        rs.update
        
        'devolvemos la clave primaria id, si tiene, sino 0
        If (idstr <> "") Then
            id = rs.fields(idstr).value
        End If
    
        'Cerramos recordset
        rs.Close
        Set rs = Nothing
        
        altaRegistro = id
    Else
        altaRegistro = 0
    End If
    
'SalirTratarError:
'    Exit Function
'TratarError:
'    MsgBox "Error (Alta Registro): " & vbNewLine & _
'            Err.Description, , "Alert: TableManagement"
'    Resume SalirTratarError
End Function

Public Function bajaRegistro(tabla As String, condition As String)
    Dim rs As ADODB.Recordset
    Dim str As String
    
    'Abrimos recordset
    str = " SELECT *" & _
          " FROM " & tabla & _
          " WHERE " & condition & ";"
    
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockOptimistic
        
    If Not (rs.EOF) Then
        rs.MoveFirst
        rs.delete
        rs.update
    Else
        MsgBox "Error(DEL) no existe ese registro( operacion tabla " & tabla & " )", vbOKOnly, "SIFOC_TableManagement"
    End If

    'Cerramos recordset
    rs.Close
    Set rs = Nothing
End Function

Public Sub actualizaRegistro(tabla As String, condition As String, strPares As String)
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim num As Integer
    Dim i As Integer
    
    'Abrimos recordset
    str = " SELECT *" & _
          " FROM " & tabla & _
          " WHERE " & condition & ";"
    
    If (Len(strPares) > 0) Then
        Set rs = New ADODB.Recordset
        rs.Open str, CurrentProject.Connection, adOpenStatic, adLockOptimistic
            
        If Not rs.EOF Then
        
            'comprobamos que los numeros de argumentos sean los mismos
            num = numeroCamposString(strPares)
            rs.MoveFirst
            If (num <= rs.fields.count) Then
                For i = 1 To num
                    'debugando fields(i) & " , " & values(i)
                    rs.fields(nombreCampoString(strPares, i)) = valorCampoString(strPares, i)
                Next i
            Else
                MsgBox "Error(UPD) numero campos( operacion tabla " & tabla & " )", vbOKOnly, "SIFOC_TableManagement"
            End If
        Else
            MsgBox "Error(UPD) no existe ese registro( operacion tabla " & tabla & " )", vbOKOnly, "SIFOC_TableManagement"
        End If
        
        rs.update

        'Cerramos recordset
        rs.Close
        Set rs = Nothing
    End If
End Sub

Public Function consultaRegistro(tabla As String, id As Long) As String
'Devuelve string con parejas de valores "campo@@valor##" comenzando por en numero de campos
'Ejemplo "2;id,1;descripcion,esto es una prueba;"

    Dim rs As ADODB.Recordset
    Dim str As String
    Dim respuesta As String
    Dim num As Integer
    Dim i As Integer
    
    'Abrimos recordset
    str = " SELECT *" & _
          " FROM " & tabla & _
          " WHERE " & primaryKey & "=" & id & ";"
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
        
    respuesta = ""
    If Not rs.EOF Then
        rs.MoveFirst
        num = rs.fields.count
        'añadimos el numero de campos como primer valor
        respuesta = num
        For i = 0 To num - 1
            'añadimos par, nombre campo, valor "nombreCampo@@valor##"
            'añadimos "nombreCampo,"
            respuesta = respuesta & separador & rs.fields(i).name
            'añadimos "valorCampo;"
            respuesta = respuesta & separador1 & rs.fields(i).value
        Next i
    Else
        MsgBox "Error(QRY) no existe ese registro( operacion tabla " & tabla & " )", vbOKOnly, "SIFOC_TableManagement"
    End If
    
    'Cerramos recordset
    rs.Close
    Set rs = Nothing
    
    consultaRegistro = respuesta
    
End Function

'----------------------------------------------------------------------------------------------
'       Funciones SQL
'----------------------------------------------------------------------------------------------
Public Function altaRegistroSQL(tabla As String, strPares As String) As Long
    Dim str As String
    Dim strInsert As String
    Dim strFields As String
    Dim strValues As String
    
    Dim value As String
    Dim num As Integer
    Dim id As Long
    Dim i As Integer

    num = numeroCamposString(strPares)

    If (num > 0) Then

        id = 0
        strFields = ""
        strValues = ""

        'Montamos los campos para el insert into
        For i = 1 To num
            If (strFields = "") Then
                strFields = "" & nombreCampoString(strPares, i) & ""
            Else
                strFields = strFields & "," & nombreCampoString(strPares, i) & ""
            End If
        Next i
        
        'Montamos los valores a insertar
        For i = 1 To num
            value = valorCampoString(strPares, i)
            If Not IsNumeric(value) Then
                value = "'" & value & "'"
            End If
            
            If (strValues = "") Then
                strValues = "" & value & ""
            Else
                strValues = strValues & "," & value & ""
            End If
        Next i
        
        str = "INSERT INTO " & tabla & " (" & strFields & ") " & _
              "VALUES (" & strValues & ");"

        'RunSQL
        CurrentDb.Execute str
                
        altaRegistroSQL = 0
    Else
        altaRegistroSQL = -1
    End If
End Function
'Public Function insertSQL()

'End Function

'Public Function updateSQL(strSelect as String,)

'End Function

'----------------------------------------------------------------------------------------------
'       funciones para tratamiento de cadena de pares
'       "numeroCampos;nombreCampo,valorCampo;..."
'----------------------------------------------------------------------------------------------

Private Function numeroCamposString(str As String) As Integer
    Dim pares As Variant    'conjunto de pares, un par en cada posicion del array
    Dim par As Variant      '(0)campo, (1)valor
    
    If (Len(str) > 0) Then
        pares = Split(str, separador)
    
        numeroCamposString = pares(0)
    Else
        numeroCamposString = 0
    End If
    
End Function

Private Function nombreCampoString(str As String, i As Integer) As String
    Dim pares As Variant    'conjunto de pares, un par en cada posicion del array
    Dim par As Variant      '(0)campo, (1)valor
    
    pares = Split(str, separador)
    
    par = Split(pares(i), separador1)
    
    nombreCampoString = par(0)
    
End Function

Public Function valorCampoString(str As String, i As Integer) As String
    Dim pares As Variant    'conjunto de pares, un par en cada posicion del array
    Dim par As Variant      '(0)campo, (1)valor
    
    pares = Split(str, separador)
    
    If (pares(0) >= i) Then
'debugando str
        par = Split(pares(i), separador1)
        valorCampoString = par(1)
    Else
        MsgBox "Indice fuera de intervalo" & vbNewLine & "i(" & i & ")>max(" & pares(0) & ")", _
               vbOKOnly, "Alert: TableManagement"
    End If
End Function

Public Function valorNombreCampoString(str As String, nombre As String) As String
    Dim numCampos As Integer
    Dim i As Integer
    Dim pares As Variant      '(0)campo, (1)valor
    
    pares = Split(str, separador)
    
    numCampos = numeroCamposString(str)
    
    'empezamos por 1 pk el primero es el numero de campos
    For i = 1 To numCampos Step 1
        If (nombreCampoString(str, i) = nombre) Then
            Exit For
        End If
    Next i

    valorNombreCampoString = valorCampoString(str, i)
    
End Function

'----------------------------------------------------------------------------------------------
'       Funciones de borrar todos los registros de la tabla
'----------------------------------------------------------------------------------------------
Public Function borrarTabla(tabla As String) As Boolean
    Dim str As String
    Dim respuesta As String
    
    str = " DELETE * FROM " & tabla & ";"
    
    CurrentDb.Execute str
    
    borrarTabla = respuesta
    
End Function

'----------------------------------------------------------------------------------------------
'       Funciones de prueba
'----------------------------------------------------------------------------------------------
Public Function pruebatm() As Integer
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim num As Integer
    Dim id As Long
    Dim i As Integer
    Dim strPares As String
    
    strPares = "2##uno@@'asf'123456789789456123##dos@@123456789"
    num = numeroCamposString(strPares)

    If (num > 0) Then
        'Abrimos recordset
        
    'debugando altaregistro("
        
    
        'Cerramos recordset
        rs.Close
        Set rs = Nothing
        
        pruebatm = id
    Else
        pruebatm = 0
    End If
End Function
