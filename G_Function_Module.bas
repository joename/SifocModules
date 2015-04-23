Attribute VB_Name = "G_Function_Module"
Option Explicit
Option Compare Database

'----------------------------------------------------------------------------------------------------
'                   Funciones eliminar acentos
'----------------------------------------------------------------------------------------------------
Function CambiaVocales(texto As String, caracter As String) As String
    Dim i As Integer
    Dim s1 As String
    Dim s2 As String
    
    s1 = "¡¿…»Õœ”“⁄‹·‡ËÈÌÔÛÚ˙¸AEIOUaeiou"
    If Len(texto) <> 0 Then
        For i = 1 To Len(s1)
            texto = Replace(texto, Mid(s1, i, 1), caracter, , , vbBinaryCompare) 'vbusecompareoption, vbtextcompare, vbdatabasecompare
        Next
    End If
    
    If (Mid(s1, 7, 1) = Mid(s1, 17, 1)) Then
        'debugando "mierda2" & Mid(s1, 7, 1) & Mid(s1, 17, 1)
    End If
    
    CambiaVocales = texto
End Function

'----------------------------------------------------------------------------------------------------
'                   Funciones eliminar acentos
'----------------------------------------------------------------------------------------------------
Function eliminaAcentos(texto As String) As String
    Dim i As Integer
    Dim s1 As String
    Dim s2 As String
    
    s1 = "¡¿…»Õœ”“⁄‹·‡ËÈÌÔÛÚ˙¸"
    s2 = "AAEEIIOOUUaaeeiioouu"
    
    If Len(texto) <> 0 Then
        For i = 1 To Len(s1)
            texto = Replace(texto, Mid(s1, i, 1), Mid(s2, i, 1), , , vbBinaryCompare)
        Next
    End If
    
    If (Mid(s1, 7, 1) = Mid(s1, 17, 1)) Then
        'debugando "mierda2" & Mid(s1, 7, 1) & Mid(s1, 17, 1)
    End If
    
    eliminaAcentos = texto
End Function

'----------------------------------------------------------------------------------------------------
'   Name: PrimeraLetraPalabraMayuscula
'   Desc: Pone la primera letra de cada palabra en may˙sculas y todas las demas en min˙sculas
'   Parm: texto as string
'   Retr: devuelve texto modificado as string
'----------------------------------------------------------------------------------------------------
Public Function PrimeraLetraPalabraMayuscula(texto As String) As String
    Dim str As String
    Dim palabras
    Dim palabra
    Dim letra As String
    Dim longitud As Integer
    
    texto = LTrim(texto)
    palabras = Split(texto, " ")
    
    str = ""
    For Each palabra In palabras
        palabra = LCase(palabra)
        longitud = Len(palabra)
        'hacemos mayuscula la primera letra
        letra = Mid(palabra, 1, 1)
        letra = UCase(letra)
        'cogemos la palabra menos la primera letra
        palabra = Mid(palabra, 2)
        'unimos las palabras ya transformadas
        If Len(str) > 0 Then
            str = str & " " & letra & palabra
        Else 'caso 1
            str = letra & palabra
        End If
    Next
    
    PrimeraLetraPalabraMayuscula = str
    
End Function

'----------------------------------------------------------------------------------------------------
'   Name: PrimeraLetraOracionMayuscula
'   Desc: Pone la primera letra de cada palabra en may˙sculas y todas las demas en min˙sculas
'   Parm: texto as string
'   Retr: devuelve texto modificado as string
'----------------------------------------------------------------------------------------------------
Public Function PrimeraLetraOracionMayuscula(texto As String) As String
    Dim str As String
    Dim letra As String
    Dim oracion As String
    
    oracion = LTrim(texto)
    
    str = ""
    'hacemos mayuscula la primera letra
    letra = Mid(oracion, 1, 1)
    letra = UCase(letra)
    'cogemos la palabra menos la primera letra y pasamos a min˙sculas
    oracion = Mid(oracion, 2)
    oracion = LCase(oracion)
    'unimos las palabras ya transformadas
    str = letra & oracion
        
    PrimeraLetraOracionMayuscula = str
    
End Function

'----------------------------------------------------------------------------------------------------
'                   Funciones de DNI
'----------------------------------------------------------------------------------------------------

Public Function letraDni(ByVal longDni As Long) As String
    letraDni = Mid("TRWAGMYFPDXBNJZSQVHLCKE", (longDni Mod 23) + 1, 1)
End Function

Public Function formaString8(dni As String) As String
    Dim longDni As Integer
    Dim newDni As String
    
    longDni = Len(dni)
    
    If (longDni < 8) Then
        newDni = Mid("00000000", 1, 8 - longDni) & dni
    Else
        newDni = Mid(dni, 1, 8)
    End If
    
    formaString8 = newDni
    
End Function

Public Function devuelveNIF(dni As String) As String
    Dim strDni As String
    Dim longDni As Integer
    Dim EXT As String
    Dim i As Integer
    
    'si no es extrangero
    EXT = ""
    longDni = 9
    
    strDni = dni
    
    If StrComp(Left(strDni, 1), "X", vbTextCompare) = 0 Or StrComp(Left(strDni, 1), "Y", vbTextCompare) = 0 Then
        EXT = Left(strDni, 1)
        longDni = 10
    End If
    
    'Quitar cualquier caracter que no sea numero
    strDni = eliminaNoNumeros(strDni)
    
    If (Len(dni) = longDni) Or (Len(strDni) = 8) Then
        'devolvemos cadena de 8 numeros
        'strDni = formaString8(strDni)
        
        'aÒadimos la letra del dni al string strdni
        strDni = EXT & strDni & letraDni(strDni)
        
        devuelveNIF = strDni
    Else
        devuelveNIF = "Error"
    End If
End Function

Public Function nifValido(dni As String) As Boolean
    If (devuelveNIF(dni) = dni) Then
        nifValido = True
    Else
        nifValido = False
    End If
End Function

'----------------------------------------------------------------------------------------------------
'                   Funciones de Telefono
'----------------------------------------------------------------------------------------------------
Public Function esTelefono(Optional telefono As String = "") As Boolean
    If IsNumeric(telefono) And Len(telefono) = 9 And (esFijo(telefono) Or esMovil(telefono)) Then
        esTelefono = True
    Else
        esTelefono = False
    End If
End Function

Public Function esFijo(Optional telefono As String = "") As Boolean
    If IsNumeric(telefono) And Len(telefono) = 9 Then
        If (Left(telefono, 1) = 9) Or (Left(telefono, 1) = 8) Then
            esFijo = True
        Else
            esFijo = False
        End If
    Else
        esFijo = False
    End If
End Function

Public Function esMovil(Optional telefono As String = "") As Boolean
    If IsNumeric(telefono) And Len(telefono) = 9 Then
        If (Left(telefono, 1) = 6) Or (Left(telefono, 1) = 7) Then
            esMovil = True
        Else
            esMovil = False
        End If
    Else
        esMovil = False
    End If
End Function

'----------------------------------------------------------------------------------------------------
'                   Funciones de Fecha
'----------------------------------------------------------------------------------------------------
Public Function miTipoFecha(fecha As Date) As Date
    miTipoFecha = Format(fecha, "dd/mm/yyyy hh:mm:ss")
End Function

Public Function miTipoHora(hora As Date) As Date
    miTipoHora = Format(hora, "dd/mm/yyyy hh:mm:ss")
End Function

Public Function sumaFechaDias(fecha As Date, nDias As Integer) As Date
    sumaFechaDias = DateAdd("d", nDias, fecha)
End Function

Public Function sumaHoraMinutos(hora As Date, nMinutos As Integer) As Date
    sumaHoraMinutos = DateAdd("n", nMinutos, hora)
End Function

'----------------------------------------------------------------------------------------------------
'                   Funciones de VARIAS
'----------------------------------------------------------------------------------------------------
Public Function IsFormLoaded(frmName As String) As Boolean
    Dim f As Form
    IsFormLoaded = False
    For Each f In Forms
       If f.name = frmName Then
          IsFormLoaded = True
          Exit Function
       End If
    Next
End Function

Public Function IsReportLoaded(repName As String) As Boolean
    Dim r As Report
    IsReportLoaded = False
    For Each r In Reports
       If r.name = repName Then
          IsReportLoaded = True
          Exit Function
       End If
    Next
End Function

Public Function argumentos(a As String) As Variant
    Dim args As Variant
    
    args = Split(a, "##")
    
    argumentos = args
End Function

Public Function eliminaNoNumeros(str As String) As String
    Dim caracter As String
    Dim newStr As String
    Dim tamanyo As Integer
    Dim i As Integer
    
    tamanyo = Len(str)
    newStr = ""

    'Quitar los espacios en blanco y los caracteres - y /
    For i = 1 To tamanyo
    
        caracter = Mid(str, i, 1)
        
        'hacemos el nuevo string con sÛlo numeros
        If IsNumeric(caracter) Then
            newStr = newStr & caracter
        End If
    Next i
    
    eliminaNoNumeros = newStr
    
End Function

'---------------------------------------------------------------------------
'
'---------------------------------------------------------------------------

Public Function devuelveString(strItems As String, _
                               nItem As Integer, _
                               separador As String) As String
    Dim items As Variant
    Dim i As Integer
    Dim nInicio As Integer
    Dim nFin As Integer
    Dim numItems As Integer
    
    items = Split(strItems, separador)
    numItems = cuentaSubstrings(strItems)
    
    If (Len(nItem) > 0) And (nItem > 0) And (nItem <= numItems) Then
        devuelveString = items(nItem - 1)
    Else
        devuelveString = ""
    End If
    
End Function

'---------------------------------------------------------------------------
'   CalculaNIF
'---------------------------------------------------------------------------
Public Function Calc_NIF(valor As String) As String
    Const cCADENA = "TRWAGMYFPDXBNJZSQVHLCKE"
    Dim resto As Integer
    Dim letra_NIF As String
   
    letra_NIF = ""

    If valor = "" Then
        MsgBox "No se ha introducido datos", , "Aviso"
        Calc_NIF = ""
        Exit Function
   ElseIf Len(valor) < 7 Then
        MsgBox "No se puede calcular el NIF, faltan dÌgitos"
        Calc_NIF = ""
       
        Exit Function
    ElseIf Not IsNumeric(valor) Then
        MsgBox "El dato introducido no es numÈrico", , "Aviso"
        Calc_NIF = ""
       
        Exit Function
    Else
        resto = val(valor) Mod 23
        letra_NIF = Mid(cCADENA, resto + 1, 1)
        Calc_NIF = valor & letra_NIF
        Exit Function
    End If
End Function

'---------------------------------------------------------------------------
'   Prueba
'---------------------------------------------------------------------------
Public Function pruebaFunctionModule()
    Dim str As String
    
    debugando "Res: "
    
End Function
