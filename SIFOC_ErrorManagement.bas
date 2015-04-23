Attribute VB_Name = "SIFOC_ErrorManagement"
Option Explicit
Option Compare Database

Private Const separador As String = "##"
Private Const separador1 As String = "@@"

'--------------------------------------------------------------------------
'           Error Management
'--------------------------------------------------------------------------
Public Function saveError(place As String, _
                          code As String, _
                          description As String, _
                          usuarioIFOC As Long) As Integer
    Dim strPares As String
    Dim Status As Integer
    
    strPares = "4"
    strPares = strPares & separador & "place" & separador1 & place
    strPares = strPares & separador & "code" & separador1 & code
    strPares = strPares & separador & "description" & separador1 & description
    strPares = strPares & separador & "fkUsuarioIFOC" & separador1 & usuarioIFOC

    Status = altaRegistroSQL("z_dberror", strPares)
    
    If (Status = -1) Then
        saveError = -1
    Else
        saveError = 0
    End If
    
End Function

Private Function altaRegistroSQL(tabla As String, strPares As String) As Long
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

debugando str
        'RunSQL
        CurrentDb.Execute str
                
        altaRegistroSQL = 0
    Else
        altaRegistroSQL = -1
    End If
End Function

'----------------------------------------------------------------------------------------------
'       Funciones para tratamiento de cadena de pares
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

Private Function valorCampoString(str As String, i As Integer) As String
    Dim pares As Variant    'conjunto de pares, un par en cada posicion del array
    Dim par As Variant      '(0)campo, (1)valor
    
    pares = Split(str, separador)
    
    If (pares(0) >= i) Then
'debugandostr
        par = Split(pares(i), separador1)
        valorCampoString = par(1)
    Else
        MsgBox "Indice fuera de intervalo" & vbNewLine & "i(" & i & ")>max(" & pares(0) & ")", _
               vbOKOnly, "Alert: TableManagement"
    End If
End Function

Private Function valorNombreCampoString(str As String, nombre As String) As String
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

