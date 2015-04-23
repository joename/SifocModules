Attribute VB_Name = "SIFOC_SMS"
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  15/07/2013 - Upd:  21/1/2009
'   Name:   getSaldoSMS
'   Desc:   Obtiene el saldo en SMS que dispone la cuenta pasada por parametro
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  long, número de créditos disponibles
'---------------------------------------------------------------------------
Public Function getSaldoSMS(idSMSCuenta As Long) As Long
    Dim saldo As Long
    
    If (idSMSCuenta = 1) Then ' Arsys
        saldo = getSaldoSMSArsys()
    ElseIf (idSMSCuenta = 2) Then ' Mensario
        saldo = getSaldoSMSMensario()
    ElseIf (idSMSCuenta = 3) Then ' Mensario
        saldo = getSaldoSMSDescom()
    Else
        saldo = -1
    End If
    
    getSaldoSMS = saldo
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  15/07/2013 - Upd:  17/07/2013
'   Name:   sendSMS
'   Desc:   Envía sms a los destinatarios pasados por parámetro de forma
'           inmediata
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  long, número de créditos disponibles
'---------------------------------------------------------------------------
Public Function sendSMS(idSMSCuenta As Long, _
                        msg As String, _
                        sender As String, _
                        recipients As String, _
                        idPersonas As String) As Integer
    Dim resultado As Long
    If (idSMSCuenta = 1) Then ' Arsys
        resultado = sendSMSArsys(msg, sender, recipients, idPersonas, idSMSCuenta)
    ElseIf (idSMSCuenta = 2) Then ' Mensario
        resultado = sendSMSMensario(msg, sender, recipients, idPersonas)
    ElseIf (idSMSCuenta = 3) Then ' Descom
        resultado = sendSMSDescom(msg, sender, recipients, idPersonas)
    End If
    
    sendSMS = resultado
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  17/07/2013 - Upd:  17/07/2013
'   Name:   sendSMS
'   Desc:   Envía sms a los destinatarios pasados por parámetro de forma
'           inmediata
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  idSmsCuenta aslong, número de créditos disponibles
'           dateTime As Date
'           msg As String
'           sender As String
'           recipients As String
'           idPersonas As String
'           Optional tlfResponsible
'---------------------------------------------------------------------------
Public Function sendScheduledSMS(idSMSCuenta As Long, _
                                 dateTime As Date, _
                                 msg As String, _
                                 sender As String, _
                                 recipients As String, _
                                 idPersonas As String) As Integer
    
    Dim resultado As Integer
    
    If (idSMSCuenta = 1) Then ' Arsys
        resultado = sendScheduledSMSArsys(dateTime, msg, sender, recipients)
    ElseIf (idSMSCuenta = 2) Then ' Mensario
        resultado = sendScheduledSMSMensario(dateTime, msg, sender, recipients, idPersonas)
    ElseIf (idSMSCuenta = 3) Then ' Mensario
        resultado = sendScheduledSMSDescom(dateTime, msg, sender, recipients, idPersonas)
    Else
        MsgBox "Error cuenta sms no definida.", vbCritical, "Alert: SIFOC_SMS"
        Debug.Print "Error cuenta sms no definida. vbCritical Alert: SIFOC_SMS"
    End If

    sendScheduledSMS = resultado
End Function


'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  17/07/2007 - Upd:  17/07/2013
'   Name:   scheduledSMS
'   Desc:   Envía sms a los destinatarios pasados por parámetro de forma
'           programada segun parametros de la función
'   Param:  idSMSCuenta, Long (id cuenta sms)
'   Retur:  idSmsCuenta aslong, número de créditos disponibles
'           mensaje As String
'           remitente As String
'           destinatarios As String
'           idPersonas As String
'           fechaHoraEP As Date
'           intervalo as Integer - en minutos
'           numSmsMensaje as integer
'           horarioDesde as date
'           horarioHasta as date
'WARNING:   los parametros de esta funcion tienen que validarse antes!!
'---------------------------------------------------------------------------
Public Function scheduleSMS(idSMSCuenta As Long, _
                            mensaje As String, _
                            remitente As String, _
                            destinatarios As String, _
                            idPersonas As String, _
                            fechaHoraInicio As Date, _
                            intervalo As Integer, _
                            numSmsMensaje As Integer, _
                            horarioDesde As Date, _
                            horarioHasta As Date, _
                            Optional ByRef progFechasTimes As Variant, _
                            Optional ByRef progPaqueteTelefonos As Variant, _
                            Optional ByRef progPaquetePersonas As Variant) As Integer
        
On Error GoTo TratarError

    Dim i As Integer
'    Dim tmp1 As Integer
'    Dim tmp2 As Integer
'    Dim nextHour As Date
'    Dim nextDate As Date
    Dim nextDateTime As Date
    Dim numTelefonos As Integer
    Dim numEnvios As Integer
    Dim numUltimoPaquete As Integer
'    Dim paqueteTelefonos As String
    Dim strDateEnvios As String         'guardamos fecha y hora de los envios
'    Dim respuestaSMS As String
'    Dim altaSMS As String
    Dim resultadoEnvioSMS As Integer
    Dim credito As Integer
    Dim msglocal As String
    
    Dim progFechaHoras As Variant           'Date() - fechaHora de envíos programados
    Dim progPaqueteTels As Variant          'String() - paquetes de telefonos
    Dim progPaqueteIdPersonas As Variant     'String() - paquetes de idPersonas
    
'    Dim id As Long
'    Dim strPares As String
'
'    Dim cuenta As String
'    Dim contrasenya As String
'
    Dim respuesta
  
    'Numero de telefonos del string destinatarios
    numTelefonos = cuentaTelefonosMoviles(destinatarios)
    credito = getSaldoSMS(idSMSCuenta) 'consultaCreditoSMS(cuenta, contrasenya)
    
    If (numTelefonos <= credito) Then
    '+++ Calculamos el numero de envios(paquetes) que habrá y preguntamos si son correctos
        calculaEnvios numEnvios, numUltimoPaquete, numTelefonos, numSmsMensaje
        
        progFechaHoras = getScheduledDateTimes(numEnvios, fechaHoraInicio, intervalo, numSmsMensaje, horarioDesde, horarioHasta)
        progPaqueteTels = devuelvePaqueteTelefonos(destinatarios, numSmsMensaje, numEnvios)
        progPaqueteIdPersonas = devuelvePaqueteIdPersonas(idPersonas, numSmsMensaje, numEnvios)
Debug.Print progPaqueteIdPersonas(1) & " - " & progPaqueteIdPersonas(2)
Debug.Print progPaqueteTels(1) & " - " & progPaqueteTels(2)
Debug.Print progFechaHoras(1) & " - " & progFechaHoras(2)
        'copiamos de vuelta paquetes
        progFechasTimes = progFechaHoras
        progPaqueteTelefonos = progPaqueteTels
        progPaquetePersonas = progPaqueteIdPersonas
        
'Calcular envios > mensaje previo envio
        nextDateTime = fechaHoraInicio
'        nextDate = CDate(Format(nextDateTime, "YYYY/MM/DD hh:mm:ss"))
'        nextHour = CDate(Format(nextDateTime, "YYYY/MM/DD hh:mm:ss"))
        strDateEnvios = ""
'        primeraFechaHora nextDateTime, intervalo, horarioDesde, horarioHasta ' nextHour,
'        For i = 1 To numEnvios Step 1
'            paqueteTelefonos = devuelvePaqueteTelefonos(destinatarios, numSmsMensaje, numEnvios)
'            If (i = numEnvios) Then
'                numTelefonos = numUltimoPaquete
'            Else
'                numTelefonos = numSmsMensaje
'            End If
'
'            'guardamos la fecha hora de cada envio para mostrar por pantalla
'            strDateEnvios = strDateEnvios & Format(nextDate, "dd/mm/yyyy") & " " & Format(nextHour, "hh:mm") & " (" & numTelefonos & ") " & vbNewLine
'
'            siguienteFechaHora nextDateTime, intervalo, horarioDesde, horarioHasta ' nextHour,
'        Next i
        strDateEnvios = printArray(progFechaHoras)
Debug.Print printArray(progFechaHoras)
        respuesta = MsgBox("Programación de los envios programados" & vbNewLine & _
                           "Envios programados:" & vbNewLine & strDateEnvios & vbNewLine & vbNewLine & _
                           "¿Enviamos sms's con esta programación?", _
                           vbYesNo, "Modulo de SMS")
    
'Enviamos sms con programación aceptada por el responsable(pero calculamos envios de nuevo).
        If (respuesta = vbYes) Then
            Dim idMensaje As Long
            
'            nextDateTime = fechaHoraInicio
'            nextDate = CDate(Format(fechaHoraInicio, "YYYY/MM/DD hh:mm:ss"))
'            nextHour = CDate(Format(fechaHoraInicio, "YYYY/MM/DD hh:mm:ss"))
'            primeraFechaHora nextDateTime, intervalo, horarioDesde, horarioHasta ' nextHour,
            
            'Guardamos mensaje pasado por parametro
            msglocal = mensaje
            
            'Programammos los diferentes envios
            For i = 1 To numEnvios Step 1
'                paqueteTelefonos = devuelvePaqueteTelefonos(destinatarios, i, numSmsMensaje)
'
'                'añadimos telefono del responsable
'                paqueteTelefonos = paqueteTelefonos
'
'                nextDateTime = CDate(Format(nextDate, "YYYY/MM/DD") & " " & Format(nextHour, "hh:mm:ss"))
                
                'Guardamos mensaje en msglocal porque tras ejecución sendscheduledsms elimina mensaje
                msglocal = mensaje
                
'Debug.Print "ANTES> nextdatetime: " & nextDateTime & vbNewLine & "mensaje:" & msglocal & vbNewLine & "paqueteTelefonos: " & paqueteTelefonos & vbNewLine
                resultadoEnvioSMS = sendScheduledSMS(idSMSCuenta, _
                                                     CDate(progFechaHoras(i)), _
                                                     msglocal, _
                                                     remitente, _
                                                     CStr(progPaqueteTels(i)), _
                                                     "")
'                resultadoEnvioSMS = sendScheduledSMS(idSMSCuenta, _
                                                     nextDateTime, _
                                                     msglocal, _
                                                     remitente, _
                                                     paqueteTelefonos, _
                                                     "")
                                                     
'Debug.Print "DESPUES> nextdatetime: " & nextDateTime & vbNewLine & "mensaje:" & msglocal & vbNewLine & "paqueteTelefonos: " & paqueteTelefonos & vbNewLine

'                siguienteFechaHora nextDateTime, intervalo, horarioDesde, horarioHasta ' nextHour,
            Next i
            
            scheduleSMS = 0
        Else
            scheduleSMS = -1
        End If
    Else
        MsgBox "Error al enviar sms. Posibles causas:" & vbNewLine & _
                   " - Saldo sms insuficiente" & "(" & numTelefonos & " > " & credito & ")" & vbNewLine & _
                   " - Error de conexión con el servidor de sms." & vbNewLine & _
                   "*** No se enviaron SMS's. ***", vbOKOnly, "Alert: SIFOC_SMS"
        scheduleSMS = -1
    End If

SalirTratarError:
    Exit Function
TratarError:
    saveError "SIFOC_SMS(ScheduleSMS)", Err.Number, Err.description, usuarioIFOC()
    Debug.Print "SIFOC_SMS(ScheduleSMS)Error: " & Err.Number & " " & Err.description
    scheduleSMS = -1
    Resume Next 'SalirTratarError
End Function

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez  Update: Jose Manuel Sanchez
'   Fecha:  29/01/2014 - Upd:  29/01/2014
'   Name:   getScheduledDateTime
'   Desc:   Calcula las fechasHora para los envíos de paquetes de sms
'   Param:  fechaHoraInicio As Date
'           intervalo as Integer - en minutos
'           numSmsMensaje as integer
'           horarioDesde as date
'           horarioHasta as date
'   Return: array de DateTimes con las programaciones
'  WARNING: los parametros de esta funcion tienen que validarse antes!!
'---------------------------------------------------------------------------
Public Function getScheduledDateTimes(nEnvios As Integer, _
                                      fechaHoraInicio As Date, _
                                      intervalo As Integer, _
                                      numSmsMensaje As Integer, _
                                      horarioDesde As Date, _
                                      horarioHasta As Date) As Variant

    Dim nextHour As Date
    Dim nextDate As Date
    Dim nextDateTime As Date
    Dim strDateEnvios As String
    Dim numTelefonos As String
    Dim i As Integer
    
    'Declare a one-dimensional Date array.
    Dim datetimes() As Date
    ReDim datetimes(1 To nEnvios)
    
    nextDateTime = fechaHoraInicio
    'nextDate = CDate(Format(nextDateTime, "YYYY/MM/DD hh:mm:ss"))
    'nextHour = CDate(Format(nextDateTime, "YYYY/MM/DD hh:mm:ss"))
    strDateEnvios = ""
    
    primeraFechaHora nextDateTime, intervalo, horarioDesde, horarioHasta
    
    
    For i = 1 To nEnvios Step 1
        'paqueteTelefonos = devuelvePaqueteTelefonos(destinatarios, i, numSmsMensaje)
'        If (i = nEnvios) Then
'            numTelefonos = numUltimoPaquete
'        Else
'            numTelefonos = numSmsMensaje
'        End If
        
        'guardamos la fecha hora de cada envio para mostrar por pantalla
'        strDateEnvios = strDateEnvios & Format(nextDateTime, "dd/mm/yyyy") & " " & Format(nextDateTime, "hh:mm") & " (" & numTelefonos & ") " & vbNewLine
        datetimes(i) = nextDateTime
        ' date(year(nextdate),month(nextdate),day(nextdate)) + TIME(hour(nexhour),minute(nexthour),second(nexthour))
        siguienteFechaHora nextDateTime, intervalo, horarioDesde, horarioHasta
    Next i
printArray (datetimes)
    getScheduledDateTimes = datetimes
End Function

Public Static Function printArray(arrayy As Variant) As String
    Dim i As Integer
    Dim str As String
    
    str = ""
    For i = LBound(arrayy) To UBound(arrayy)
        str = str & CStr(arrayy(i)) & vbNewLine
    Next
Debug.Print str
    printArray = str
End Function

