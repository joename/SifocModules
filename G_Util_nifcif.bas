Attribute VB_Name = "G_Util_nifcif"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  17/08/2010 - Actualización:  17/08/2010
'   Name:   existeNif
'   Desc:   Devuelve si el nif existe en la bd, excepto el idPersona pasado
'   Param:  nif(string)
'           idPersona(long), identificador de persona
'   Retur:  Verdadero, si el nif se encuentra en la bd y no es idPersona
'           Falso, si el nif no se encuentra en la bd y no es idPersona
'---------------------------------------------------------------------------------
Public Function existeNif(nif As String, _
                          Optional idPersona = 0) As Boolean
    Dim numNif As String
    Dim rs As ADODB.Recordset
    Dim str As String
    Dim isNif As Boolean
    
    numNif = Left(nif, 8) 'eliminaNoNumeros(nif)
    
    str = " SELECT id, nombre, apellido1, dni" & _
          " FROM t_persona" & _
          " WHERE dni Like '%" & numNif & "%'" & IIf(idPersona = 0, ";", "AND id <> " & idPersona & ";")
    
    Set rs = New ADODB.Recordset
    rs.Open str, CurrentProject.Connection, adOpenStatic, adLockReadOnly
    
    If Not rs.EOF Then
        isNif = True
    Else
        isNif = False
    End If
    
    rs.Close
    Set rs = Nothing
    
    existeNif = isNif
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernandez
'   Fecha:  09/02/1/2009
'   Desc:   Verifica CIF y nos dice si es correcto o no.
'   Param:  cif, string con el cid correcto!!
'                 formato dni: NNNNNNNNL
'                   nie: CNNNNNNNL
'                  C = :A,B,C,D,E,F,G,H,K,L,M,P,Q,S,X"
    
'                 N : numero entre 0 y 9
'                 L : letra A-Z mayúscula, calculada de los primeros 8 caracteres.
'   Retur:  Boolean
'             true si CIF es correcto
'             false si CIF es incorrecto
'---------------------------------------------------------------------------
 Public Function VerificaCIF(cif As String) As Boolean
 
    Dim letraCIF As String, strNumero As String, quitaPrimeraLetra As String
    Dim quitaPrimeraLetraAux As String
    Dim auxNum As Integer
    Dim i As Integer
    Dim suma As Integer
    Dim letras As String
    Dim esCif As Boolean
    
    'inicializamos variables
    esCif = False
    letras = "ABCDEFGHKLMPQSX" 'primera letras del CIF
    cif = UCase(cif)
    If Len(cif) < 9 Or Not IsNumeric(Mid(cif, 2, 7)) Then
        
        esCif = False
    End If
    letraCIF = Mid(cif, 1, 1)       'letra del CIF
    strNumero = Mid(cif, 2, 7)      'Codigo de Control
    quitaPrimeraLetra = Mid(cif, 9) 'CIF menos primera y ultima posiciones
    
    If InStr(letras, letraCIF) = 0 Then 'comprobamos la letra del CIF (1ª posicion)
        esCif = False
    End If
   
    i = 0
    For i = 1 To 7
        If i Mod 2 = 0 Then
            suma = suma + CInt(Mid(strNumero, i, 1))
        Else
            auxNum = CInt(Mid(strNumero, i, 1)) * 2
            suma = suma + (auxNum \ 10) + (auxNum Mod 10)
            'debugando suma
        End If
    Next
    suma = (10 - (suma Mod 10)) Mod 10
    
    Select Case letraCIF
    
    Case "K", "P", "Q", "S"
        suma = suma + 64
        quitaPrimeraLetraAux = Chr(suma)
    Case "X"
        quitaPrimeraLetraAux = Mid(VerificaDNI(strNumero), 8, 1)
    Case Else
        quitaPrimeraLetraAux = CStr(suma)
    End Select
    'debugando quitaPrimeraLetraAux
    
    If quitaPrimeraLetra = quitaPrimeraLetraAux Then
        esCif = True
        debugando DameFormaJuridica(cif)
        'Me.txt_FormaJuridica = DameFormaJuridica(Me.txt_CIF)
    Else
        esCif = False
    End If
    VerificaCIF = esCif
End Function


'---------------------------------------------------------------------------
'   Autor:  Nelson A. Hernandez Payano
'   Fecha:  10/2/2009
'   Desc:   Verifica DNI o CIF y nos dice si es correcto o no.
'   Param:  dniCif, string con el dniCif !!
'                 formato dni: NNNNNNNNL
'                   CIF: CNNNNNNNN
'                 C : ABCDEFGHKLMPQSX
'                 N : numero entre 0 y 9
'                 L : letra A-Z mayúscula, calculada de los primeros 8 caracteres.
'   Retur:  String
'             si DNI/CIF es correcto
'             no DNI/CIF es incorrecto
'---------------------------------------------------------------------------
Public Function VerificaDniCif(dnicif As String) As Boolean
    Dim esCorrectoDniCif As Boolean
    Dim primerCaracter As String
    Dim ultimoCaracter As String
    Dim caracterDniNie As String
        
    If (Len(dnicif) = 9) Then 'Si logitud = 9 verificar dni/cif
        primerCaracter = Mid(dnicif, 1, 1) 'extrae el primer caracter para saber si es numero o letra
        ultimoCaracter = Mid(dnicif, 9, 1) 'extrae el primer caracter para saber si es numero o letra
       
        'Miramos si es el primer digito es numero o caracter del dni o nie para tratarlos por separado
        If IsNumeric(primerCaracter) _
            Or primerCaracter = "X" _
            Or primerCaracter = "Y" _
            Or primerCaracter = "Z" Then  'es NIF o NIE
            esCorrectoDniCif = VerificaDniNie(dnicif)
        Else ' es CIF porque ultimo caracter es numérico
            esCorrectoDniCif = VerificaCIF(dnicif)
        End If
    Else 'no es dniCif pk no tiene longutid = 9
        esCorrectoDniCif = False
    End If
    
    VerificaDniCif = esCorrectoDniCif
    
End Function

'funcion que recibe un CIF y devuelve la forma jurídica
Public Function DameFormaJuridica(cif As String) As String

    Dim rs As New ADODB.Recordset
    Dim letra As String
    Dim FormaJuridica As String
    Dim strSql As String
    FormaJuridica = ""

    If IsNull(cif) Or cif = "" Then
        DameFormaJuridica = ""
    Else
        letra = Left(cif, 1)
        strSql = "SELECT FormaJuridica " & _
                 "FROM a_organizacionformajuridica " & _
                 "WHERE Letra = '" & letra & "';"
        rs.Open strSql, CurrentProject.Connection, adOpenDynamic, adLockOptimistic
        If Not rs.EOF Then
            DameFormaJuridica = rs!FormaJuridica
        Else
            DameFormaJuridica = ""
        End If
        rs.Close
        Set rs = Nothing
    End If

End Function

