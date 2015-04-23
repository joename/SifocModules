Attribute VB_Name = "G_DNINIE_Module"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Desc:   Calcula el ultimo digito(letra) del DNI
'   Param:  Dni, string ( 8 numeros)
'                 formato dni: NNNNNNNN
'                         nie: NNNNNNNN
'   Retur:  devuelve dniNie correcto NNNNNNNL
'---------------------------------------------------------------------------
Public Function CalculoLetraDni(dni As String) As String
    Const cCADENA = "TRWAGMYFPDXBNJZSQVHLCKE" 'string letra dni/nie
    Dim posletra As Integer
    Dim letra_Dni As String
    
    'Inicializamos variables
    letra_Dni = ""
    
    posletra = val(dni) Mod 23 'Nos da la posicion de la letra, dentro del string Dni/nie
        
    letra_Dni = Mid(cCADENA, posletra + 1, 1)
    
    CalculoLetraDni = letra_Dni 'devuelvo la letra obtenida

End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Desc:   Verifica DNI nos dice si es correcto o no.
'   Param:  dni, string con el NIE correcto!!
'                 formato dni: NNNNNNNNL
'                 N : numero entre 0 y 9
'                 L : letra A-Z mayúscula, calculada de los primeros 8 caracteres.
'   Retur:  boolean
'           true,  si DNI/NIE es correcto
'           false, si DNI/NIE es incorrecto
'---------------------------------------------------------------------------
Public Function VerificaDNI(dni As String) As Boolean
    Dim nif As String
    Dim numeroDni As String
    Dim concaUltimaLetra As String
    Dim compara As String
    Dim esDniCorrecto As Boolean
    
    'Inicializamos variables
    esDniCorrecto = False
    
    dni = UCase(dni) 'ponemos la letra en mayúscula
    numeroDni = Mid(dni, 1, Len(dni) - 1) 'quitamos la letra del DNI
    
    If Len(numeroDni) = 8 And IsNumeric(numeroDni) Then 'verifico que el dni sea igual a 8 y que sea numerico
        nif = numeroDni & CalculoLetraDni(numeroDni) 'calculamos la letra del DNI para comparar con el que tenemos
    'Else 'Verificar_DNI = True
    '    debugando "El dato introducido no corresponde a un DNI"
    End If
   
    If nif <> dni Then 'comparamos las DNI
        'debugando "El DNI " & dni & " es INCORRECTO" & "D.N.I. Correcto: " & letraObtenida
        'debugando "El DNI " & DNI & " es INCORRECTO"
        esDniCorrecto = False
    Else
        'debugando "El NIE " & DNI & " es CORRECTO"
        esDniCorrecto = True
    End If
    
    VerificaDNI = esDniCorrecto
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Desc:   Quita la primera letra del NIE y le asigna (si es X=0,si es Y=1,si es Z=2)
'   Param:  dninie, string con el nie correcto!!
'                 formato nie: CNNNNNNNL
'                 C : X o Y o Z
'                 N : numero entre 0 y 9
'                 L : letra A-Z mayúscula, calculada de los primeros 8 caracteres.
'   Retur:  boolean
'           true,  si DNI/NIE es correcto
'           false, si DNI/NIE es incorrecto
'---------------------------------------------------------------------------
Public Function VerificaNie(nie As String) As Boolean
    Dim primeraLetra As String
    Dim letraNIE As Integer
    Dim nieSinPrimeraLetra As String
    Dim concatena As String     'concatena = nie modificado( como dni/nif)
    Dim esNieCorrecto As Boolean
    
    'Inicializamos varibles
    esNieCorrecto = False
    
    primeraLetra = Mid(nie, 1, 1) 'saco la primera letra para ver si es X,Y o Z
    nieSinPrimeraLetra = Mid(nie, 2, 8) 'EL nie sin la primera letra
    
    'Seleciona una opcion con el dato de la variable primeraLetra
    If (primeraLetra = "X") Then
        letraNIE = 0
        concatena = letraNIE & nieSinPrimeraLetra
        esNieCorrecto = VerificaNieModificado(concatena)
    ElseIf (primeraLetra = "Y") Then
        letraNIE = 1
        concatena = letraNIE & nieSinPrimeraLetra
        esNieCorrecto = VerificaNieModificado(concatena)
    ElseIf (primeraLetra = "Z") Then
        letraNIE = 2
        concatena = letraNIE & nieSinPrimeraLetra
        esNieCorrecto = VerificaNieModificado(concatena)
    End If
    
    VerificaNie = esNieCorrecto
    
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Desc:   Verifica  NIE modificado y nos dice si es correcto o no.
'           Nie modificado(nie con X, Y o Z cambiada por 0, 1 o 2 resp.)
'   Param:  niemodificado, string con el nie correcto!!
'                 formato NIE: LNNNNNNNL
'                 N : numero entre 0 y 9
'                 L : letra A-Z mayúscula, calculada de los primeros 8 caracteres.
'   Retur:  boolean
'           true,  si NIE es correcto
'           false, si NIE es incorrecto
'---------------------------------------------------------------------------
Public Function VerificaNieModificado(nieModificado As String) As Boolean
        VerificaNieModificado = VerificaDNI(nieModificado)
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Desc:   Verifica DNI o NIE y nos dice si es correcto o no.
'   Param:  dninie, string con el dni/nie correcto!!
'                 formato dni: NNNNNNNNL
'                   nie: CNNNNNNNL
'                 C : X o Y o Z
'                 N : numero entre 0 y 9
'                 L : letra A-Z mayúscula, calculada de los primeros 8 caracteres.
'   Retur:  String
'             si DNI/NIE es correcto
'             no DNI/NIE es incorrecto
'---------------------------------------------------------------------------
Public Function VerificaDniNie(DniNie As String) As String
    Dim esCorrectoDniNie As Boolean
    Dim primerCaracter As String
    Dim caracterDniNie As String
        
    If (Len(DniNie) = 9) Then 'Si logitud = 9 verificar dni/nie
        primerCaracter = Mid(DniNie, 1, 1) 'extrae el primer caracter para saber si es numero o letra
        
        'Miramos si es el primer digito es numero o caracter del dni o nie para tratarlos por separado
        If IsNumeric(primerCaracter) Then 'es DNI
            esCorrectoDniNie = VerificaDNI(DniNie)
                'debugando "Hola"
                
        'Si el primer caracter es una letra llamo ala function verifica letra y le paso como parametro el nie
        Else 'no puede ser dni, miramos si es nie
            If primerCaracter = "X" Or primerCaracter = "Y" Or primerCaracter = "Z" Then
                esCorrectoDniNie = VerificaNie(DniNie)
            Else 'no empieza por X o Y o Z -> no es nie
                esCorrectoDniNie = False
            End If
        End If
    Else 'no es dninie pk no tiene longutid = 9
        esCorrectoDniNie = False
    End If
    
    VerificaDniNie = esCorrectoDniNie
    
End Function


