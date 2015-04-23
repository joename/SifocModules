Attribute VB_Name = "G_SMSArsys"
Option Explicit
Option Compare Database

Public Function getSaldoSMSArsys() As Integer
    'Dim strPares As String
    Dim account As Long
    Dim cuenta As String
    Dim contrasenya As String
    
    'consultamos cuenta y contraseña de T_SMSCuenta
    'strPares = consultaRegistro("T_SMSCuenta", fkSMSCuenta)
    account = 1 'Cuenta con Arsys
    cuenta = DLookup("[cuenta]", "t_smscuenta", "[id]=" & account) 'valorNombreCampoString(strPares, "cuenta")
    contrasenya = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & account) 'valorNombreCampoString(strPares, "contrasenya")
    
'debugando cuenta & "|" & contrasenya
    getSaldoSMSArsys = consultaCreditoSMS(cuenta, contrasenya)
End Function

Public Function sendSMSArsys(mensaje As String, _
                             remitente As String, _
                             destinatarios As String, _
                             idPersonas As String, _
                             Optional idSMSCuenta As Long = 1) As Integer

    Dim cuenta As String
    Dim contrasenya As String
    Dim respuestaSMS As String
    'Dim strPares As String
    Dim id As Long
    Dim numTels As Integer
    Dim credito As Variant
    Dim envioOK As Integer
    
On Error GoTo TratarError

    'consultamos cuenta y contrasenya
    cuenta = DLookup("[cuenta]", "t_smscuenta", "[id]=" & idSMSCuenta)
    contrasenya = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & idSMSCuenta)
    'strPares = consultaRegistro("T_SMSCuenta", fkSMSCuenta)
    'cuenta = valorNombreCampoString(strPares, "cuenta")
    'contrasenya = valorNombreCampoString(strPares, "contrasenya")
    
    envioOK = -1
    
    'Añadimos telefono responsable
'    If esMovil(telefonoResponsable) Then
'        destinatarios = destinatarios & ";" & telefonoResponsable
'    End If
    
    numTels = cuentaTelefonosMoviles(destinatarios)
    If (numTels = cuentaSubstrings(destinatarios)) Then
        credito = consultaCreditoSMS(cuenta, contrasenya)
        
'debugando credito & "!" & VarType(credito)
'(VarType(credito) = vbInteger) And

        If (numTels <= credito) Then
            envioOK = enviaSMS(cuenta, contrasenya, mensaje, remitente, destinatarios)
            'respuestaSMS = "8##fkSMSMensaje@@18##destinatarios@@12683##descripcionEP@@no programado##respuestaXML@@<SendSMS>" & _
                           "<result>OK</result><description>Envío OK</description><credit>4810</credit></SendSMS>##respuestaResultado@@OK##respuestaDescripcion@@Envío OK##respuestaCredito@@4810##SMSenviados@@1"
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
                   "*** No se enviaron SMS's. ***", vbOKOnly, "Alert: SIFOC_SMS"
            envioOK = -1
        End If
    Else
        MsgBox "Error, en telefono/s del envio.", vbOKOnly, "Alert: SIFOC_SMS"
    End If
    
    sendSMSArsys = envioOK
    
SalirTratarError:
    Exit Function
TratarError:
    MsgBox "Error(EnviaMensaje):" & vbNewLine & _
            Err.description, vbOKOnly, "Alert: SIFOC_SMS_Arsys"
            
    saveError "SIFOC_SMS(EnviaMensaje)", Err.Number, Err.description, usuarioIFOC()
    sendSMSArsys = -1
    Resume SalirTratarError
End Function


'esta funcion recive todos los campos correctos, la validacion debe ser anterior!!
Public Function sendScheduledSMSArsys(dateTime As Date, _
                                      mensaje As String, _
                                      remitente As String, _
                                      destinatarios As String, _
                                      Optional ByRef respuestaSMS As String) As Integer

    Dim periodo As String
    Dim respuestaXML As String
    Dim respuestaResultado As String
    Dim respuestaDescripcion As String
    Dim respuestaCredito As String
    Dim cantidadSMS As String
    Dim strDevuelto As String
    
    Dim strPares As String
    Dim account As String
    Dim cuenta As String
    Dim contrasenya As String
    
    Dim descripcion As String
'    Dim fecha As Date
'    Dim hora As Date

On Error GoTo Error_EnvioProgramado

    'Fecha desdoblada en dia y hora
'    fecha = CDate(Format(dateTime, "YYYYMMDD"))
'    hora = CDate(Format(dateTime, "hh:mm:ss"))

    Dim objSendSMS As Object
    
    descripcion = "Ifoc SMS programado"
    
    'consultamos cuenta y contrasenya
    account = 1 ' id cuenta Arsys
    cuenta = DLookup("[cuenta]", "t_smscuenta", "[id]=" & account) ' idSMSCuenta Mensario = 2
    contrasenya = DLookup("[contrasenya]", "t_smscuenta", "[id]=" & account)
        
    Set objSendSMS = CreateObject("SMSCOM.SMSSend")

    'Defino las propiedades(envio)
    objSendSMS.setAccount (cuenta)
    objSendSMS.setPwd (contrasenya)
    objSendSMS.SetText (mensaje)
    objSendSMS.setTo (destinatarios)
    objSendSMS.setFrom (remitente)      ' Parámetro opcional

    'Defino las propiedades(envio programado)
    objSendSMS.setDescriptionEP (descripcion)
    objSendSMS.setDateEP (Format(dateTime, "dd/mm/yyyy"))               'formato "dd/mm/yyyy"
    objSendSMS.setTimeEP (Format(dateTime, "hh:mm:ss"))                 'formato "hh:mm"
    periodo = "periodUnica"
    objSendSMS.setPeriodEP (periodo)

    'Programacion del envio de SMS
    respuestaXML = objSendSMS.Program   'envio programado
    
    'Resultado de la operación
    respuestaResultado = objSendSMS.getResult
    respuestaDescripcion = objSendSMS.getDescription
    respuestaCredito = objSendSMS.getCredit
    cantidadSMS = cuentaTelefonosMoviles(destinatarios)
    If (cantidadSMS <> cuentaTelefonos(destinatarios)) Then
        MsgBox "Error grave: avise al administrador", vbOKOnly, "SMS Module"
    End If
    
    Set objSendSMS = Nothing
    
    'Montamos str respuesta
    'destinatarios, descripcionEP, fechaEP, horaEP, periodo, respuestaXML, respuestaResultado, respuestaDescripcion, respuestaCredito, SMSenviados
    strDevuelto = "10"
    strDevuelto = strDevuelto & separador & "destinatarios" & separador1 & destinatarios
    strDevuelto = strDevuelto & separador & "descripcionEP" & separador1 & descripcion
    strDevuelto = strDevuelto & separador & "fechaEP" & separador1 & dateTime
    strDevuelto = strDevuelto & separador & "horaEP" & separador1 & dateTime
    strDevuelto = strDevuelto & separador & "periodo" & separador1 & periodo
    strDevuelto = strDevuelto & separador & "respuestaXml" & separador1 & respuestaXML
    strDevuelto = strDevuelto & separador & "respuestaResultado" & separador1 & respuestaResultado
    strDevuelto = strDevuelto & separador & "respuestaDescripcion" & separador1 & respuestaDescripcion
    strDevuelto = strDevuelto & separador & "respuestaCredito" & separador1 & respuestaCredito
    strDevuelto = strDevuelto & separador & "SMSenviados" & separador1 & cantidadSMS
    
    respuestaSMS = strDevuelto
    
    If (respuestaResultado = "OK") Then
        sendScheduledSMSArsys = 0
    Else
        sendScheduledSMSArsys = -1
    End If

Exit_EnvioProgramado:
    Exit Function

Error_EnvioProgramado:
    MsgBox "Error envio programado Arsys (" & Err.description & ")", vbOKOnly, "SMS Module"
    Debug.Print strDevuelto
    Resume Exit_EnvioProgramado

End Function

'Esta funcion devuelve el numero de sms que nos quedan por enviar
Private Function consultaCreditoSMS( _
                                    cuenta As String, _
                                    contrasenya As String) As String
    Dim strPares As String

    Dim respuestaCredito As String

    Dim objSendSMS As Object

On Error GoTo Error_ConsultaCredito

    Set objSendSMS = CreateObject("SMSCOM.SMSSend")

    'Defino las propiedades(envio)
    objSendSMS.setAccount (cuenta)
    objSendSMS.setPwd (contrasenya)

    'Dara un error al enviar pero obtendremos el credito que nos queda.
    'dura un par de segundos en la consulta!!
    objSendSMS.Send

    'Consulta de crédito
    respuestaCredito = objSendSMS.getCredit

    'desvinculamos objeto de sms
    Set objSendSMS = Nothing

    'devolvemos credito restante
    consultaCreditoSMS = respuestaCredito

Exit_ConsultaCredito:
    Exit Function

Error_ConsultaCredito:
    MsgBox "Error consulta credito (" & Err.description & ")", vbOKOnly, "SMS Module"
    Resume Exit_ConsultaCredito

End Function

'------------------------------------------------------------------------------------------------
'                   pruebas
'------------------------------------------------------------------------------------------------
Public Function pruebaSMSIfoc()
    Dim TEL As String
    
    Dim fecha As Date
    Dim hora As Date
    Dim nextDateTime As Date
    Dim nextDate As Date
    Dim nextHour As Date
    Dim intervalo As Integer
    Dim horarioDesde As Date
    Dim horarioHasta As Date
    Dim respuestaSMS As String
    Dim i As Integer
    
    intervalo = 28
    horarioDesde = "08:00"
    horarioHasta = "14:00"
    
    respuestaSMS = "10" & separador & "destinatarios" & separador1 & "destinatarios" & _
        separador & "descripcionEP" & separador1 & "descripcion" & _
        separador & "fechaEP" & separador1 & "fecha" & _
        separador & "horaEP" & separador1 & "hora" & _
        separador & "periodo" & separador1 & "periodo" & _
        separador & "respuestaXml" & separador1 & "respuestaXML" & _
        separador & "respuestaResultado" & separador1 & "respuestaResultado" & _
        separador & "respuestaDescripcion" & separador1 & "respuestaDescripcion" & _
        separador & "respuestaCredito" & separador1 & "respuestaCredito" & _
        separador & "SMSenviados" & separador1 & "1"
    
    nextDateTime = "01/01/2008 22:00"
'    nextHour = "22:00"
    primeraFechaHora nextDateTime, intervalo, horarioDesde, horarioHasta '  nextHour,
    For i = 1 To 10 Step 1
        'Guardamos SMS en base de datos(guardamos id personas, no teléfonos)
        'altaSMS "1", "test", "1,2,3,4", "mensajito", Now(), respuestaSMS
'debugando "'" & nextDate & "' | '" & nextHour & "'"
        siguienteFechaHora nextDateTime, intervalo, horarioDesde, horarioHasta ', nextHour,
    Next i
    
    'TEL = "600000000;600000000;600000000;600000000;600000000;600000000;600000000;600000000;600000000;600000000;600000000"
    'enviaMensajeProgramado 2, Date, "hola cara bola", "ifoc", "1;2;3", TEL, "descripcion", Date, miTipoHora(Now), 20, 3, "09:00", "15:00"
End Function
