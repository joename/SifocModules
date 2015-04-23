Attribute VB_Name = "G_SMSMensario"
Option Explicit
Option Compare Database

Dim G_msg As String

'General'
Private Declare Function ConvCStringToVBString Lib "kernel32" Alias "lstrcpyA" (ByVal lpsz As String, ByVal pt As Long) As Long

'Funciones generales'
Private Declare Function SetLoginParamsAPI Lib "semapi.dll" (ByVal AUserName As Long, ByVal APassword As Long, ByVal ALicense As Long, ByVal AServer As Long, ByVal APort As Long) As Boolean
Private Declare Function SetProxyParamsAPI Lib "semapi.dll" (ByVal AProxyServer As Long, ByVal AProxyPort As Long, ByVal AProxyUsername As Long, ByVal AproxyPassword As Long) As Boolean

'Sincronización'
Private Declare Function SetTimeZoneSynchronization Lib "semapi.dll" (ByVal ATimeZone As String) As Boolean
Private Declare Function ExecuteSynchronization Lib "semapi.dll" () As Long
Private Declare Function GetDateTime Lib "semapi.dll" () As Long

'Saldo'
Private Declare Function ExecuteBalanceEnquiry Lib "semapi.dll" () As Long
Private Declare Function GetBalance Lib "semapi.dll" () As Long

'Envio'
Private Declare Function SetTimeZoneSending Lib "semapi.dll" (ByVal ATimeZone As String) As Boolean
Private Declare Function ExecuteSending Lib "semapi.dll" () As Long
Private Declare Function AddRequestMessage Lib "semapi.dll" (ByVal ASender As Long, ByVal lpText As Long, ByVal ADate As String) As Long
Private Declare Function AddMessageRecipient Lib "semapi.dll" (ByVal AMsgIndex As Long, ByVal ACod As String, ByVal APhn As String) As Long
Private Declare Function GetSendRequestId Lib "semapi.dll" () As Long
Private Declare Function GetSendMessageId Lib "semapi.dll" (ByVal AMsgIndex As Long) As Long

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  15/07/2013 - Upd:  22/17/2013
'   Name:   getSaldoSMSMensario
'   Desc:   Obtiene el saldo en SMS que dispone la cuenta pasada por parametro
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  long, número de créditos disponibles
'---------------------------------------------------------------------------
Public Function getSaldoSMSMensario() As Long
    Dim usr As String   ' Usuario
    Dim pass As String  ' Password
    Dim lic As String ' Licencia
    Dim account As Long
    
    Dim prx As String   ' ProxyServer
    Dim port As String  ' Puerto
    Dim prxusr As String ' Proxy user
    Dim prxpass As String ' Proxy password
    
    account = 2 ' cuenta mensario id=2
    usr = DLookup("[cuenta]", "t_smscuenta", "[id]=" & account)
    pass = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & account)
    lic = DLookup("[licencia]", "t_smscuenta", "[id]=" & account)
    
    prx = ""
    port = ""
    prxpass = ""
    
    Dim msg As String
    msg = ""
    If SetLoginParamsAPI(StrPtr(usr & vbNullChar), _
                         StrPtr(pass), _
                         StrPtr(lic), _
                         StrPtr("es.servicios.mensario.com"), 0) _
                         And SetProxyParamsAPI(StrPtr(prx), _
                                                      val(port), _
                                                      StrPtr(port), _
                                                      StrPtr(prxpass)) Then
        If PCharToString(ExecuteBalanceEnquiry()) = "OK" Then
            msg = PCharToString(GetBalance())
        End If
    End If
    'r = MsgBox("Saldo: " + msg, 0, "Respuesta")

    getSaldoSMSMensario = CLng(msg)
End Function

Function PCharToString(ByVal lpString As Long) As String
    Dim zpos As Long
    Dim s As String
    s = String(255, 0)
    ConvCStringToVBString s, lpString
    zpos = InStr(s, vbNullChar)
    s = Left(s, zpos - 1)
    PCharToString = s
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  15/07/2013 - Upd: 22/17/2013
'   Name:   sendSMSMensario
'   Desc:   Envio SMS a los móviles pasados por parámetro y telefóno responsable
'           en el momento(ahora)
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  msg As String
'           sender As String
'           recipients As String
'           idPersonas As String
'---------------------------------------------------------------------------
Public Function sendSMSMensario(msg As String, _
                                sender As String, _
                                recipients As String, _
                                idPersonas As String) As Integer
    Dim account As Long
    Dim user As String
    Dim pass As String
    Dim lic As String
        
    Dim prx As String     ' ProxyServer
    Dim port As String    ' Puerto
    Dim prxusr As String  ' Proxy user
    Dim prxpass As String ' Proxy password
    
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
    
On Error GoTo TratarError

    'consultamos cuenta y contrasenya
    account = 2 ' id cuenta Mensario
    user = DLookup("[cuenta]", "t_smscuenta", "[id]=" & account) ' idSMSCuenta Mensario = 2
    pass = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & account)
    lic = DLookup("[licencia]", "t_smscuenta", "[id]=" & account)
        
    prx = ""
    port = ""
    prxusr = ""
    prxpass = ""
    
    envioOK = -1
    
    numTels = cuentaTelefonosMoviles(recipients)
    If (numTels = cuentaSubstrings(recipients)) Then
        credito = getSaldoSMS(account)
        
        If (numTels <= credito) Then
            
            'Envío SMS a uno o varios telefonos móviles
            If SetLoginParamsAPI(StrPtr(user & vbNullChar), _
                                 StrPtr(pass), _
                                 StrPtr(lic), _
                                 StrPtr("es.servicios.mensario.com"), 0) _
                And SetProxyParamsAPI(StrPtr(prx), _
                                      val(port), _
                                      StrPtr(prxusr), _
                                      StrPtr(prxpass)) Then
                idmsg = AddRequestMessage(StrPtr(sender), StrPtr(msg), "00000000000000")
                'Añadimos teléfonos
                cadenas = Split(recipients, separadorTelefonos)
                For Each subcadena In cadenas
                    idrcp = AddMessageRecipient(idmsg, "34", subcadena)
                Next
                
                SetTimeZoneSending ("Europe/Madrid")
                If PCharToString(ExecuteSending()) = "OK" Then
                    envioOK = 0
Dim id1 As Long
Dim balance As Long
id1 = GetSendMessageId(0)
balance = GetBalance()
Debug.Print id1 & " -id/balance- " & balance
                    msg = "Petición: " + str(GetSendRequestId()) + Chr(13) + Chr(10) _
                          + "Mensaje: " + str(GetSendMessageId(0)) + Chr(13) + Chr(10) _
                          + "Crédito utilizado: " + (credito - GetBalance())
                End If
            End If
            
            'envioOK = enviaSMS(cuenta, contrasenya, FechaCreacion, mensaje, remitente, destinatarios)
            
            If envioOK = 0 Then
                respuestaSMS = "correcto"
            Else
                respuestaSMS = "incorrecto"
            End If
            
            'guardamos
            'id = altaSMS(fkSMSCuenta, remitente, fkPersonas, mensaje, FechaCreacion, respuestaSMS)
        Else
            MsgBox "Error al enviar sms. Posibles causas:" & vbNewLine & _
                   " - Error de conexión con el servidor de sms." & vbNewLine & _
                   " - Saldo sms insuficiente" & "(" & numTels & " > " & credito & ")" & vbNewLine & _
                   "*** No se enviaron SMS's. ***", vbOKOnly, "Alert: SIFOC_SMS2"
            envioOK = -1
        End If
    Else
        MsgBox "Error, en telefono/s del envio.", vbOKOnly, "Alert: SIFOC_SMS2"
    End If
    
    sendSMSMensario = envioOK
    
SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error(EnviaMensaje):" & vbNewLine & _
            Err.description, vbOKOnly, "Alert: SIFOC_SMS_Mensario"
            
    saveError "SIFOC_SMS2(EnviaMensaje)", Err.Number, Err.description, usuarioIFOC()
    sendSMSMensario = -1
    Resume SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  15/07/2012 - Upd: 22/17/2013
'   Name:   sendScheduledSMSMensario
'   Desc:   Envio SMS a los móviles pasados por parámetro y telefóno responsable
'           en el momento(ahora)
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  dateTime As Date
'           msg As String
'           sender As String
'           recipients As String
'           idPersonas As String
'           Optional tlfResponsible
'---------------------------------------------------------------------------
Public Function sendScheduledSMSMensario(dateTime As Date, _
                                        msg As String, _
                                        sender As String, _
                                        recipients As String, _
                                        idPersonas As String) As Integer
    Dim account As Long
    Dim user As String
    Dim pass As String
    Dim lic As String
        
    Dim prx As String     ' ProxyServer
    Dim port As String    ' Puerto
    Dim prxusr As String  ' Proxy user
    Dim prxpass As String ' Proxy password
    
    Dim respuestaSMS As String
    Dim id As Long
    Dim numTels As Integer
    Dim credito As Variant
    Dim envioOK As Integer
    
    Dim fechaHora As String
    Dim cadenas As Variant
    Dim i As Integer
    Dim subcadena As Variant
    
    Dim idmsg As Long
    Dim idrcp As Long
    
On Error GoTo TratarError

    'consultamos cuenta y contrasenya
    account = 2 ' id cuenta Mensario
    user = DLookup("[cuenta]", "t_smscuenta", "[id]=" & account) ' idSMSCuenta Mensario = 2
    pass = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & account)
    lic = DLookup("[licencia]", "t_smscuenta", "[id]=" & account)
        
    prx = ""
    port = ""
    prxusr = ""
    prxpass = ""
    
    envioOK = -1
    
    numTels = cuentaTelefonosMoviles(recipients)
    If (numTels = cuentaSubstrings(recipients)) Then
        credito = getSaldoSMS(account)
        
        If (numTels <= credito) Then
            
            'Envío SMS a uno o varios telefonos móviles
            If SetLoginParamsAPI(StrPtr(user & vbNullChar), _
                                 StrPtr(pass), _
                                 StrPtr(lic), _
                                 StrPtr("es.servicios.mensario.com"), 0) _
                And SetProxyParamsAPI(StrPtr(prx), _
                                      val(port), _
                                      StrPtr(prxusr), _
                                      StrPtr(prxpass)) Then
                fechaHora = Format(dateTime, "YYYYMMDDhhmmss")
                idmsg = AddRequestMessage(StrPtr(sender), StrPtr(msg), fechaHora)
                'Añadimos teléfonos
                cadenas = Split(recipients, separadorTelefonos)
                For Each subcadena In cadenas
                    idrcp = AddMessageRecipient(idmsg, "34", subcadena)
                Next
                
                SetTimeZoneSending ("Europe/Madrid")
                If PCharToString(ExecuteSending()) = "OK" Then
                    envioOK = 0
                    msg = "Petición: " + str(GetSendRequestId()) + Chr(13) + Chr(10) _
                          + "Mensaje: " + str(GetSendMessageId(0)) + Chr(13) + Chr(10) _
                          + "Crédito utilizado: " + (credito - GetBalance())
                End If
            End If
            
            'envioOK = enviaSMS(cuenta, contrasenya, FechaCreacion, mensaje, remitente, destinatarios)
            
            If envioOK = 0 Then
                respuestaSMS = "correcto"
            Else
                respuestaSMS = "incorrecto"
            End If
            
            'guardamos
            'id = altaSMS(fkSMSCuenta, remitente, fkPersonas, mensaje, FechaCreacion, respuestaSMS)
        Else
            MsgBox "Error al enviar sms. Posibles causas:" & vbNewLine & _
                   " - Error de conexión con el servidor de sms." & vbNewLine & _
                   " - Saldo sms insuficiente" & "(" & numTels & " > " & credito & ")" & vbNewLine & _
                   "*** No se enviaron SMS's. ***", vbOKOnly, "Alert: SIFOC_SMS2"
            envioOK = -1
        End If
    Else
        MsgBox "Error, en telefono/s del envio.", vbOKOnly, "Alert: SIFOC_SMS2"
    End If
    
    sendScheduledSMSMensario = envioOK
    
SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error(EnviaMensaje):" & vbNewLine & _
            Err.description, vbOKOnly, "Alert: SIFOC_SMS_Mensario"
            
    saveError "SIFOC_SMS2(EnviaMensaje)", Err.Number, Err.description, usuarioIFOC()
    sendScheduledSMSMensario = -1
    Resume SalirTratarError
End Function

'Private Sub BotonEnvio_Click()
'    'msg = "" '
'    If SetLoginParamsAPI(StrPtr(Usuario.Text & vbNullChar), StrPtr(Clave.Text), StrPtr(Licencia.Text), StrPtr("es.servicios.mensario.com"), 0) And SetProxyParamsAPI(StrPtr(proxyServer.Text), Val(proxyPort.Text), StrPtr(proxyUsuario.Text), StrPtr(proxyClave.Text)) Then
'        idmsg = AddRequestMessage(StrPtr(rte.Text), StrPtr(texto.Text), "00000000000000")
'        idrcp = AddMessageRecipient(id, cod.Text, tlf.Text)
'        SetTimeZoneSending ("Europe/Brussels")
'        If PCharToString(ExecuteSending()) = "OK" Then
'            msg = "Petición: " + str(GetSendRequestId()) + Chr(13) + Chr(10) + "Mensaje: " + str(GetSendMessageId(0))
'        End If
'    End If
'    r = MsgBox(msg, 0, "Respuesta")
'End Sub
'
'Private Sub BotonSaldo_Click()
'    msg = ""
'    If SetLoginParamsAPI(StrPtr(Usuario.Text & vbNullChar), StrPtr(Clave.Text), StrPtr(Licencia.Text), StrPtr("es.servicios.mensario.com"), 0) And SetProxyParamsAPI(StrPtr(proxyServer.Text), Val(proxyPort.Text), StrPtr(proxyUsuario.Text), StrPtr(proxyClave.Text)) Then
'        If PCharToString(ExecuteBalanceEnquiry()) = "OK" Then
'            msg = PCharToString(GetBalance())
'        End If
'    End If
'    r = MsgBox("Saldo: " + msg, 0, "Respuesta")
'End Sub
'
'Private Sub BotonSincronizacion_Click()
'    msg = ""
'    If SetLoginParamsAPI(StrPtr(Usuario.Text & vbNullChar), StrPtr(Clave.Text), StrPtr(Licencia.Text), StrPtr("es.servicios.mensario.com"), 0) And SetProxyParamsAPI(StrPtr(proxyServer.Text), Val(proxyPort.Text), StrPtr(proxyUsuario.Text), StrPtr(proxyClave.Text)) Then
'        SetTimeZoneSynchronization ("Europe/Brussels")
'        If PCharToString(ExecuteSynchronization()) = "OK" Then
'            msg = PCharToString(GetDateTime())
'         End If
'    End If
'    r = MsgBox("Fecha y hora del servidor(AAAAMMDDhhmmss): " + msg, 0, "Respuesta")
'End Sub
