Attribute VB_Name = "G_String"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  22/03/2011
'   Name:   concatString
'   Desc:   Genera un string aleatorio para un password
'   Param:  longitud(int)opt, longitud del string a general
'   Retur:  string, cadena de caracteres generado aleatoriamente
'---------------------------------------------------------------------------
Public Function concatString(strBase As String, strNew As String, Optional separador As String = ", ") As String
    Dim strConcat As String
    
    strConcat = strBase
    If (strBase = "") Then
        strConcat = strNew
    ElseIf Len(strNew) > 0 Then
        strConcat = strBase & separador & strNew
    End If
    
    concatString = strConcat
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  19/10/2009
'   Name:   generatePassword
'   Desc:   Genera un string aleatorio para un password
'   Param:  longitud(int)opt, longitud del string a general
'   Retur:  string, cadena de caracteres generado aleatoriamente
'---------------------------------------------------------------------------
Public Function generatePassword(Optional longitud As Integer) As String
    Randomize
    Dim i As Byte
    Dim tipoCaracter As Integer
    Dim j As Integer
    Dim h As Integer
    Dim resultado As String
    Dim Numeros(10) As String
    Dim Mayusculas(25) As String
    Dim Minusculas(25) As String
    Dim Simbolos(15) As String
    Dim generacionPSS() As String
    ReDim generacionPSS(longitud)
    
    j = 0
    'Cargamos array numeros con numeros chr(48-57)
    For i = 48 To 57
        Numeros(j) = Chr(i)
        j = j + 1
    Next i
    
    j = 0
    'Cargamos array mayúsculas con mayúsculas chr(65-90)
    For i = 65 To 90
        Mayusculas(j) = Chr(i)
        j = j + 1
    Next i
    
    j = 0
    'Cargamos array minusculas con minusculas chr(97-122)
    For i = 97 To 122
        Minusculas(j) = Chr(i)
        j = j + 1
    Next i
    
    j = 0
    'Cargamos array simbolos con simbolos chr(34-47,95)
    'For i = 34 To 47
    '    Simbolos(j) = Chr(i)
    '    j = j + 1
    'Next i
    'j = j + 1
    'Simbolos(j) = Chr(95) ' underscore(_)
    
    For h = 0 To longitud - 1
        'aleatorio para saber si concateno numero, mayuscula o minuscula
        tipoCaracter = Int((3 * Rnd))
        Select Case tipoCaracter
            'numero
            Case 0:
                generacionPSS(h) = Numeros(Int((UBound(Numeros) * Rnd)))
            'minuscula
            Case 1:
                generacionPSS(h) = Minusculas(Int((UBound(Minusculas) * Rnd)))
            'mayuscula
            Case 2:
                generacionPSS(h) = Mayusculas(Int((UBound(Mayusculas) * Rnd)))
            'símbolos
            'Case 3:
            '    generacionPSS(h) = Simbolos(Int((UBound(Simbolos) * Rnd)))
        End Select
    Next h
    
    resultado = ""
    'concateno resultado para devolver solo un string y asi olvidarse de los vectores
    For j = 0 To longitud - 1
        resultado = resultado & generacionPSS(j)
    Next j
    
    generatePassword = resultado
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/9/2009
'   Name:   countSubstrings
'   Desc:   Cuenta el númeoro de substrings separados por un delimitador
'   Param:  cadena, string de caracteres
'           delimiter, caracter separador de strings
'   Retur:  número de strings en la cadena
'---------------------------------------------------------------------------
Public Function countSubStrings(cadena As String, _
                                delimiter As String) As Integer
    Dim i As Integer
    Dim lenght As Integer
    Dim counter As Integer
    Dim cadenas As Variant
    Dim subcadena As Variant
    
    cadenas = Split(cadena, delimiter)
    i = 0
    counter = 0
    For Each subcadena In cadenas
        counter = counter + 1
    Next
    countSubStrings = counter
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/9/2009
'   Name:   mixString
'   Desc:   mezcla el string
'   Param:  str, string de caracteres
'   Retur:  string, mezclado
'---------------------------------------------------------------------------
Public Function mixString(Optional texto As String = "") As String
    Dim i As Integer
    Dim i1 As Integer
    Dim i2 As Integer
    Dim c As String
    Dim resultado As String
    Dim miValor As Integer
    
    Dim longitud As Integer

    longitud = Len(texto)
    i1 = 1
    If (longitud Mod 2 = 0) Then
        i2 = longitud
    Else
        i2 = longitud - 1
    End If
    
    Randomize    ' Inicializa el generador de números aleatorios.
    resultado = ""
    For i = i1 To i2
        miValor = Int((i2 * Rnd) + 1)
        resultado = resultado & Mid(texto, miValor, 1)
    Next i
    
    If Len(resultado) = 0 Then
        resultado = " "
    End If
    mixString = resultado

End Function
