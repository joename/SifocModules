Attribute VB_Name = "G_SMSDescom"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  15/07/2013 - Upd:  22/07/2013
'   Name:   getSaldoSMSDescom
'   Desc:   Obtiene el saldo en SMS que dispone la cuenta pasada por parametro
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  long, número de créditos disponibles
'---------------------------------------------------------------------------
Public Function getSaldoSMSDescom() As Long
    Dim usr As String   ' Usuario
    Dim pass As String  ' Password
    Dim lic As String ' Licencia
    Dim account As Long
    
    Dim prx As String   ' ProxyServer
    Dim port As String  ' Puerto
    Dim prxusr As String ' Proxy user
    Dim prxpass As String ' Proxy password
    
    Dim saldo As Long
    
    account = 3 ' cuenta Descom id=3
    usr = DLookup("[cuenta]", "t_smscuenta", "[id]=" & account)
    pass = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & account)
    'lic = DLookup("[licencia]", "t_smscuenta", "[id]=" & account)
    
    'Crear Clase
    Dim DMSMS As Object 'New dcXMLSend.XMLSendClass
    Set DMSMS = CreateObject("dcXMLSend.XMLSendClass")
    
    'Asignar acceso a plataforma Descom Mensajes
    DMSMS.Usuario = usr
    DMSMS.Clave = pass

    'Obtener Saldo
    DMSMS.getSaldo

    'Verificar acceso
    If DMSMS.RXAutentificacion.resultado = "1" Then
        saldo = DMSMS.RXAutentificacion.saldo
        'MsgBox "Saldo de la cuenta: " & DMSMS.RXAutentificacion.saldo, vbInformation, "Saldo"
    Else
        saldo = -1
        'MsgBox DMSMS.RXAutentificacion.comentario, vbCritical, "Error de Acceso"
    End If
    
    Set DMSMS = Nothing
    getSaldoSMSDescom = saldo
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  15/07/2013 - Upd:  22/07/2013
'   Name:   sendSMSDescom
'   Desc:   Envio SMS a los móviles pasados por parámetro y telefóno responsable
'           en el momento(ahora)
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  msg As String
'           sender As String
'           recipients As String
'           idPersonas As String
'---------------------------------------------------------------------------
Public Function sendSMSDescom(msg As String, _
                              sender As String, _
                              recipients As String, _
                              idPersonas As String) As Integer
    Dim account As Long
    Dim user As String
    Dim pass As String
    
    Dim respuestaSMS As String
    Dim id As Long
    Dim numTels As Integer
    Dim credito As Variant
    Dim envioOK As Integer
    
    Dim cadenas As Variant
    Dim i As Integer
    Dim subcadena As Variant
    
    Dim idmsg As Long
    Dim idrcp As Long
    Dim resultado As String
    
On Error GoTo TratarError

    'consultamos cuenta y contrasenya
    account = 3 ' id cuenta Descom
    user = DLookup("[cuenta]", "t_smscuenta", "[id]=" & account) ' idSMSCuenta Mensario = 2
    pass = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & account)
    
    envioOK = -1
    
    'Crear Clase
    Dim DMSMS As Object
    Set DMSMS = Nothing
    Set DMSMS = CreateObject("dcXMLSend.XMLSendClass")
    
    'Asignar acceso a plataforma Descom Mensajes
    DMSMS.Usuario = user
    DMSMS.Clave = pass

    DMSMS.remitente = Left(sender, 11)
    
    numTels = cuentaTelefonosMoviles(recipients)
    
Debug.Print recipients & " :rec = cuent: " & cuentaSubstrings(recipients)

    If (numTels = cuentaSubstrings(recipients)) Then
        credito = getSaldoSMS(account)
        
        If (numTels <= credito) Then
            
            'Envío SMS a uno o varios telefonos móviles
            'Añadimos teléfonos
            cadenas = Split(recipients, separadorTelefonos)
            Dim n As Integer
            n = 0
            For Each subcadena In cadenas
                'Crear Mensaje
                n = n + 1 'identificador de mensaje en envío 1..n
                DMSMS.XMLMensajes.MensajesSMS.Add "" & n, "0034" & subcadena, msg, sender
            Next
            
            'Id Envío
            DMSMS.IdEnvioext = "EnvioSMS" & Format(now, "YYYYMMDDhhmmss")
            'Enviar
            DMSMS.SendXML "", "" 'Sin programación se envía directamente en el momento

            envioOK = 0
            
            If envioOK = 0 Then
                respuestaSMS = "correcto"
            Else
                respuestaSMS = "incorrecto"
            End If
            
        Else
            MsgBox "Error al enviar sms. Posibles causas:" & vbNewLine & _
                   " - Error de conexión con el servidor de sms." & vbNewLine & _
                   " - Saldo sms insuficiente" & "(" & numTels & " > " & credito & ")" & vbNewLine & _
                   "*** No se enviaron SMS's. ***", vbOKOnly, "Alert: SIFOC_SMS3"
            envioOK = -1
        End If
    Else
        MsgBox "Error, en telefono/s del envio.", vbOKOnly, "Alert: SIFOC_SMS3"
    End If
    
    'Verificar acceso
    'If MsgBox("¿Desea un resumen del resultado del envío?", vbYesNo, "Alert: Envío SMS") Then
        'Creamos fichero resumen de envío
        Dim fs As Object
        Dim a As Object
        Dim Ruta As String
        
        'Creamos fichero XML en la carpeta donde se encuentra SIFOC
        Ruta = CurrentProject.path
        Ruta = Ruta & "\Resumen envío sms" & Format(now, "YYYYMMDDhhmmss") & ".txt"
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(Ruta, True, True)
        
        If DMSMS.RXAutentificacion.resultado = "1" Then  'si resultado es correcto
            resultado = "Saldo de la cuenta: " & DMSMS.RXAutentificacion.saldo _
                & vbCrLf & "ID Envio Descom: " & DMSMS.IdEnviodm _
                & vbCrLf & "ID Envio DLL: " & DMSMS.IdEnvioext _
                & vbCrLf & "Total mensajes: " & DMSMS.TotalMensajesEnviados _
                & vbCrLf & "Total mensajes enviaros: " & DMSMS.TotalMensajesEnviadosOK _
                & vbCrLf & "Total errores: " & DMSMS.TotalMensajesEnviadosError _
                & vbCrLf & "Créditos gastados: " & DMSMS.TotalCreditosGastados _
                & vbCrLf & "---------------------------------------------------" & vbCrLf & vbNewLine
            a.write (resultado)
            Dim msge  'As DMSMS.MensajeSMS
            resultado = "Descripción de envío por mensaje:" & vbCrLf
            a.write (resultado)
            For Each msge In DMSMS.XMLMensajes.MensajesSMS
                If msge.resultado Then
                    resultado = "Mensajes " & msge.Key & ": " & msge.comentario & " - Id envio: " & msge.Iddm & vbCrLf ', vbInformation
                Else
                    resultado = "Mensajes " & msge.Key & ": " & msge.comentario & " - Id envio: " & msge.Iddm & vbCrLf ', vbExclamation
                End If
                a.write (resultado)
            Next
        Else 'si resultado es incorrecto
            MsgBox DMSMS.RXAutentificacion.comentario, vbCritical, "Error de Acceso"
        End If
        
        'Cerramos fichero
        a.Close
        Set a = Nothing
    'End If
    
    'Vaciamos la colección de mnesajes para no replicar envíos
    DMSMS.XMLMensajes.MensajesSMS.clean
        
    sendSMSDescom = envioOK
    
SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error(EnviaMensaje):" & vbNewLine & _
            Err.description, vbOKOnly, "Alert: SIFOC_SMS_Descom"
            
    saveError "SIFOC_SMS3(EnviaMensaje)", Err.Number, Err.description, usuarioIFOC()
    sendSMSDescom = -1
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  15/07/2013 - Upd:  22/07/2013
'   Name:   sendScheduledSMSDescom
'   Desc:   Envio SMS a los móviles pasados por parámetro y telefóno responsable
'           en el momento(ahora)
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  dateTime As Date - día y hora en que se realiza la entrega
'           msg As String
'           sender As String
'           recipients As String
'           idPersonas As String
'           Optional tlfResponsible
'---------------------------------------------------------------------------
Public Function sendScheduledSMSDescom(dateTime As Date, _
                                       msg As String, _
                                       sender As String, _
                                       recipients As String, _
                                       idPersonas As String) As Integer
    Dim account As Long
    Dim user As String
    Dim pass As String
    
    Dim respuestaSMS As String
    Dim id As Long
    Dim numTels As Integer
    Dim credito As Variant
    Dim envioOK As Integer
    
    Dim cadenas As Variant
    Dim i As Integer
    Dim subcadena As Variant
    
    Dim idmsg As Long
    Dim idrcp As Long
    Dim resultado As String
    
On Error GoTo TratarError

    'consultamos cuenta y contrasenya
    account = 3 ' id cuenta Descom
    user = DLookup("[cuenta]", "t_smscuenta", "[id]=" & account) ' idSMSCuenta Mensario = 2
    pass = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & account)
    
    envioOK = -1
    
    'Crear Clase
    Dim DMSMS As Object
    Set DMSMS = Nothing
    Set DMSMS = CreateObject("dcXMLSend.XMLSendClass")
    
    'Asignar acceso a plataforma Descom Mensajes
    DMSMS.Usuario = user
    DMSMS.Clave = pass

    DMSMS.remitente = Left(sender, 11)
    
    numTels = cuentaTelefonosMoviles(recipients)
    If (numTels = cuentaSubstrings(recipients)) Then
        credito = getSaldoSMS(account)
        
        If (numTels <= credito) Then
            'Definición de programación
            Dim prog As Object
            
            Set prog = CreateObject("dcXMLSend.programacion")
            Set prog = DMSMS.programadoUnaSolaVez("ProgamacionEnvioSMS" & Format(dateTime, "YYYYMMDDhhmmss"), dateTime)
            
            'Envío SMS a uno o varios telefonos móviles
            'Añadimos teléfonos
            cadenas = Split(recipients, separadorTelefonos)
            Dim n As Integer
            n = 0
            For Each subcadena In cadenas
                'Crear Mensaje
                n = n + 1 'identificador de mensaje en envío 1..n
                DMSMS.XMLMensajes.MensajesSMS.Add "" & n, "0034" & subcadena, msg, sender
            Next
            
            'Id Envío
            DMSMS.IdEnvioext = "EnvioSMSProg" & Format(dateTime, "YYYYMMDDhhmmss")
            'Enviar
            DMSMS.SendXML "", "", , prog 'Sin programación se envía directamente en el momento

            envioOK = 0
            
            If envioOK = 0 Then
                respuestaSMS = "correcto"
            Else
                respuestaSMS = "incorrecto"
            End If
            
        Else
            MsgBox "Error al enviar sms. Posibles causas:" & vbNewLine & _
                   " - Error de conexión con el servidor de sms." & vbNewLine & _
                   " - Saldo sms insuficiente" & "(" & numTels & " > " & credito & ")" & vbNewLine & _
                   "*** No se enviaron SMS's. ***", vbOKOnly, "Alert: SIFOC_SMS3"
            envioOK = -1
        End If
    Else
        MsgBox "Error, en telefono/s del envio.", vbOKOnly, "Alert: SIFOC_SMS3"
    End If

    
    'Verificar acceso
    'If MsgBox("¿Desea un resumen del resultado del envío?", vbYesNo, "Alert: Envío SMS") Then
        'Creamos fichero resumen de envío
        Dim fs As Object
        Dim a As Object
        Dim Ruta As String
        
        'Creamos fichero XML en la carpeta donde se encuentra SIFOC
        Ruta = CurrentProject.path
        Ruta = Ruta & "\Resumen envío sms programado" & Format(now, "YYYYMMDDhhmmss") & ".txt"
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(Ruta, True, True)
        
        If DMSMS.RXAutentificacion.resultado = "1" Then  'si resultado es correcto
            resultado = "Saldo de la cuenta: " & DMSMS.RXAutentificacion.saldo _
            & vbCrLf & "ID Envio Descom: " & DMSMS.IdEnviodm _
            & vbCrLf & "ID Envio DLL: " & DMSMS.IdEnvioext _
            & vbCrLf & "Mensajes Programado: " & IIf(DMSMS.ResultProgramacion, "Correctamente", "Error: revise que la programación es correcta.") _
            & vbCrLf & "Identificativo de la programación: " & DMSMS.ResultProgramacionId _
            & vbCrLf & "Mensajes programados correctamente: " & DMSMS.ResultProgramacionMsgOk _
            & vbCrLf & "Mensajes programados con errores: " & DMSMS.ResultProgramacionMsgErr _
                & vbCrLf & "---------------------------------------------------" & vbCrLf & vbNewLine
            
            a.write (resultado)
            
        Else 'si resultado es incorrecto
            MsgBox DMSMS.RXAutentificacion.comentario, vbCritical, "Error de Acceso"
        End If
        
        'Cerramos fichero
        a.Close
        Set a = Nothing
    'End If
    
    'Vaciamos la colección de mnesajes para no replicar envíos
    DMSMS.XMLMensajes.MensajesSMS.clean
        
    sendScheduledSMSDescom = envioOK

Debug.Print "BEFORE EXIT 1: " & msg

SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error(EnviaMensaje):" & vbNewLine & _
            Err.description, vbOKOnly, "Alert: SIFOC_SMS_Descom"
            
    saveError "SIFOC_SMS3(EnviaMensaje)", Err.Number, Err.description, usuarioIFOC()
    sendScheduledSMSDescom = -1
End Function
