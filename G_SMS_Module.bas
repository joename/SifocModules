Attribute VB_Name = "G_SMS_Module"
Option Explicit
Option Compare Database

Public Const maxMessageLong As Integer = 160
Public Const maxTitleLong As Integer = 11

Global Const separadorTelefonos As String = ";" 'Public (former scope)

Public Const separadorPersonas As String = ";"

Public Function addfkPersona(strfkPersonas As String, idPersona As String)
    Dim Personas As String

    If (strfkPersonas = "") Then
        Personas = idPersona
    Else
        Personas = strfkPersonas & separadorPersonas & idPersona
    End If

    addfkPersona = Personas

End Function

Public Function addTelefono(strTelefonos As String, _
                            telefono As String, _
                            Optional mode As Integer = 0) As String
    'Mode = 1 -> añade telefono al principio
    'Mode = 0 -> añade telefono al final
    
    Dim tels As String
    tels = strTelefonos
    
    If Len(telefono) > 0 Then
        If esMovil(telefono) Then
            If (mode = 1) Then
                If (strTelefonos = "") Then
                    tels = telefono
                Else
                    tels = telefono & separadorTelefonos & strTelefonos
                End If
            ElseIf (mode = 0) Then
                If (strTelefonos = "") Then
                    tels = telefono
                Else
                    tels = strTelefonos & separadorTelefonos & telefono
                End If
            End If
        Else
            tels = strTelefonos
        End If
    End If
    addTelefono = tels
    
End Function

Public Function cuentaSubstrings(str As String) As Integer
    'el caracter separador entre telefonos es el ";"
    Dim cadenas As Variant
    Dim i As Integer
    Dim counter As Integer
    Dim subcadena As Variant
    
    cadenas = Split(str, separadorTelefonos)
    i = 0
    counter = 0
    For Each subcadena In cadenas
        counter = counter + 1
    Next
    
    cuentaSubstrings = counter
    
End Function

Public Function cuentaTelefonos(str As String) As Integer
    'el caracter separador entre telefonos es el ";"
    Dim telefonos As Variant
    Dim i As Integer
    Dim counter As Integer
    Dim telefono As Variant
    Dim tlf As String
    
    telefonos = Split(str, separadorTelefonos)
    i = 0
    counter = 0
    For Each telefono In telefonos
        tlf = telefono
        If esTelefono(tlf) Then
            counter = counter + 1
        End If
    Next
    
    cuentaTelefonos = counter
    
End Function

Public Function cuentaTelefonosMoviles(str As String) As Integer
    'el caracter separador entre telefonos es el ";"
    Dim telefonos As Variant
    Dim i As Integer
    Dim counter As Integer
    Dim telefono As Variant
    Dim tlf As String
    
    telefonos = Split(str, separadorTelefonos)
    i = 0
    counter = 0
    For Each telefono In telefonos
        tlf = telefono
        If esMovil(tlf) Then
            counter = counter + 1
        End If
    Next
    
    cuentaTelefonosMoviles = counter
    
End Function

Public Function devuelvePaqueteTelefonos(strTelefonos As String, _
                                         numSmsPaquete As Integer, _
                                         numPaquetes As Integer) As Variant
'Parametro numPaquete As Integer, _

    Dim paqTelefonos() As String
    ReDim paqTelefonos(1 To numPaquetes)
    
    Dim telefonos As Variant
    Dim i As Integer
    Dim nInicio As Integer
    Dim nFin As Integer
    Dim numTelefonos As Integer
    Dim strPaquete As String
    Dim numPaquete As Integer
    
    telefonos = Split(strTelefonos, separadorTelefonos)
    numTelefonos = cuentaTelefonosMoviles(strTelefonos)
    
    For numPaquete = 1 To numPaquetes
        strPaquete = ""
        If (numPaquete > 0) Then
            'calculamos valor inicial de bucle(array empieza en 0)
            nInicio = (numPaquete - 1) * numSmsPaquete
            If (nInicio > numTelefonos) Then
                nInicio = numTelefonos
            End If
            
            nFin = (nInicio + numSmsPaquete) - 1
            If (nFin >= numTelefonos) Then
                nFin = numTelefonos - 1
            End If
        Else 'controlamos numPaquete por debajo de 0, para nInicio y nFin
            nInicio = 1
            nFin = 0
        End If
    
        strPaquete = ""
        For i = nInicio To nFin Step 1
            If (strPaquete = "") Then
                strPaquete = telefonos(i)
            Else
                strPaquete = strPaquete & ";" & telefonos(i)
            End If
        Next i
        paqTelefonos(numPaquete) = strPaquete
    Next
'printArray (paqTelefonos)
    devuelvePaqueteTelefonos = paqTelefonos
    
End Function

Public Function devuelvePaqueteIdPersonas(strIdPersonas As String, _
                                          numSmsPaquete As Integer, _
                                          numPaquetes As Integer) As Variant
'Parametro numPaquete As Integer, _

    Dim paqIdPersonas() As String
    ReDim paqIdPersonas(1 To numPaquetes)
    
    Dim idPersonas As Variant
    Dim i As Integer
    Dim nInicio As Integer
    Dim nFin As Integer
    Dim numPersonas As Integer
    Dim strPaquete As String
    Dim numPaquete As Integer
    
    idPersonas = Split(strIdPersonas, separadorTelefonos)
    numPersonas = UBound(idPersonas) - LBound(idPersonas) + 1        'cuentaTelefonosMoviles(strIdPersonas)
    
    For numPaquete = 1 To numPaquetes
        strPaquete = ""
        If (numPaquete > 0) Then
            'calculamos valor inicial de bucle(array empieza en 0)
            nInicio = (numPaquete - 1) * numSmsPaquete
            If (nInicio > numPersonas) Then
                nInicio = numPersonas
            End If
            
            nFin = (nInicio + numSmsPaquete) - 1
            If (nFin >= numPersonas) Then
                nFin = numPersonas - 1
            End If
        Else 'controlamos numPaquete por debajo de 0, para nInicio y nFin
            nInicio = 1
            nFin = 0
        End If
        
        strPaquete = ""
        For i = nInicio To nFin Step 1
            If (strPaquete = "") Then
                strPaquete = idPersonas(i)
            Else
                strPaquete = strPaquete & ";" & idPersonas(i)
            End If
        Next i
        paqIdPersonas(numPaquete) = strPaquete
    Next
printArray (paqIdPersonas)
    devuelvePaqueteIdPersonas = paqIdPersonas
    
End Function

'esta funcion recive todos los campos correctos, la validacion debe ser anterior!!
'mirar antes si hay credito suficiente para evitar el error!!
Public Function enviaSMS(cuenta As String, _
                        contrasenya As String, _
                        mensaje As String, _
                        remitente As String, _
                        destinatarios As String) As Integer

    Dim destinatariosResponsable As String
    Dim periodo As String
    Dim respuestaXML As String
    Dim respuestaResultado As String
    Dim respuestaDescripcion As String
    Dim respuestaCredito As String
    Dim cantidadSMS As Integer
    
    Dim id As Long
    Dim respuestaSMS As String

    Dim objSendSMS As Object
    
On Error GoTo Error_Envio

    Set objSendSMS = CreateObject("SMSCOM.SMSSend")
    
    destinatariosResponsable = destinatarios

    'Defino las propiedades(envio)
    objSendSMS.setAccount (cuenta)
    objSendSMS.setPwd (contrasenya)
    objSendSMS.SetText (mensaje)
    objSendSMS.setTo (destinatariosResponsable)
    objSendSMS.setFrom (remitente)      ' Parámetro opcional

    'Resultado de la operación
    respuestaXML = objSendSMS.Send
    respuestaResultado = objSendSMS.getResult
    respuestaDescripcion = objSendSMS.getDescription
    respuestaCredito = objSendSMS.getCredit
    cantidadSMS = cuentaTelefonosMoviles(destinatariosResponsable)

    If (cantidadSMS <> cuentaTelefonos(destinatariosResponsable)) Then
        MsgBox "Error grave: avise al administrador", vbOKOnly, "SMS Module"
    End If
    
    Set objSendSMS = Nothing
    
    respuestaSMS = "8"
    respuestaSMS = respuestaSMS & separador & "fkSMSMensaje" & separador1 & id
    respuestaSMS = respuestaSMS & separador & "destinatarios" & separador1 & destinatarios
    respuestaSMS = respuestaSMS & separador & "descripcionEP" & separador1 & "no programado"
    respuestaSMS = respuestaSMS & separador & "respuestaXml" & separador1 & respuestaXML
    respuestaSMS = respuestaSMS & separador & "respuestaResultado" & separador1 & respuestaResultado
    respuestaSMS = respuestaSMS & separador & "respuestaDescripcion" & separador1 & respuestaDescripcion
    respuestaSMS = respuestaSMS & separador & "respuestaCredito" & separador1 & respuestaCredito
    respuestaSMS = respuestaSMS & separador & "SMSenviados" & separador1 & cantidadSMS
    
    'enviaSMS = respuestaSMS
    enviaSMS = 0
Exit_Envio:
    Exit Function

Error_Envio:
    enviaSMS = -1
    MsgBox "Error envio (" & Err.description & ")", vbOKOnly, "SMS Module"
    Resume Exit_Envio

End Function


Public Sub primeraFechaHora(ByRef fechaHora As Date, _
                            intervalo As Integer, _
                            horarioDesde As Date, _
                            horarioHasta As Date)
'Eliminamos: ByRef hora As Date, _

    'intervalo en minutos
    
    Dim nextHour As Date
    Dim nextDate As Date
    Dim myDate As Date  'hoy
    
    If (intervalo < (23 * 60)) Then
        'fecha y hora que nos pasan del formulario
        myDate = fechaHora 'CDate(Format(fecha, "dd/mm/yyyy") & " " & Format(hora, "hh:mm:ss")) 'miTipoFecha(fecha)
        
        'INICIALIZAMOS nextHour y nextDate
        'Cogemos hora y fecha minima de envio como la actual
        'nextHour minima damos 5min para errores de hora servidor y hora pc.
'        nextHour = myDate
'        nextDate = myDate 'miTipoFecha(fecha)
        
        'Actualizamos horarioDesde y horarioHasta para que la fecha sea mínimo la de inicio(No modificamos hora)
        horarioDesde = CDate(Format(fechaHora, "dd/mm/yyyy") & " " & Format(horarioDesde, "hh:mm:ss"))
        horarioHasta = CDate(Format(fechaHora, "dd/mm/yyyy") & " " & Format(horarioHasta, "hh:mm:ss"))
                
        'Nos aseguramos que las fecha minima sea la actual y no una anterior
'        If (nextDate < myDate) Then
'            nextDate = myDate
'        End If
        'Si myDate(hora) < horarioDesde(hora)
        If (myDate < horarioDesde) Then
            nextHour = CDate(Format(myDate, "dd/mm/yyyy") & " " & Format(horarioDesde, "hh:mm:ss"))
        End If
        'Si myDate(hora) > horarioDesde(hora)
        If (myDate > horarioHasta) Then
            myDate = CDate(Format(myDate, "dd/mm/yyyy") & " " & Format(horarioDesde, "hh:mm:ss"))
            myDate = sumaFechaDias(myDate, 1)
            'nextHour = CDate(Format(nextDate, "dd/mm/yyyy") & " " & Format(horarioDesde, "hh:mm:ss"))
        End If
        
        'Actualizamos variables a retornar
        fechaHora = myDate
'        hora = nextHour
    Else
        MsgBox "Error: intervalo demasiado grande(max.23horas)", vbOKOnly, "SMS Module"
    End If
End Sub

Public Sub siguienteFechaHora(ByRef fechaHora As Date, _
                              intervalo As Integer, _
                              horarioDesde As Date, _
                              horarioHasta As Date)
'ByRef hora As Date, _

    'intervalo en minutos
    Dim nextDateTime As Date
'    Dim nextHour As Date
'    Dim nextDate As Date
    Dim myHour As Date
    Dim myDate As Date
    Dim horarioFrom As Date
    Dim horarioTo As Date
    
    If (intervalo < (23 * 60)) Then
        'INICIALIZAMOS nextHour y nextDate
        'Cogemos hora y fecha minima de envio como la actual
        'nextHour minima damos 5min para errores de hora servidor y hora pc.
        nextDateTime = fechaHora
'        nextHour = CDate(Format(fecha, "dd/mm/yyyy") & " " & Format(hora, "hh:mm:ss")) 'miTipoHora(hora)
'        nextDate = CDate(Format(fecha, "dd/mm/yyyy") & " " & Format(hora, "hh:mm:ss")) 'miTipoFecha(Date)
        
        'fecha y hora que nos pasan del formulario
'        myHour = CDate(Format(fecha, "dd/mm/yyyy") & " " & Format(hora, "hh:mm:ss")) 'miTipoHora(hora)
'        myDate = CDate(Format(fecha, "dd/mm/yyyy") & " " & Format(hora, "hh:mm:ss")) 'miTipoFecha(fecha)
        
        'horario from y to (modificamos fecha, no hora, a la fecha para que comparen sólo hora)
        horarioFrom = CDate(Format(fechaHora, "dd/mm/yyyy") & " " & Format(horarioDesde, "hh:mm:ss"))
        horarioTo = CDate(Format(fechaHora, "dd/mm/yyyy") & " " & Format(horarioHasta, "hh:mm:ss"))
        
        'Nos aseguramos que las fecha minima sea la actual y no una anterior
'        If (nextDate < myDate) Then
'            nextDate = myDate
'        End If
        
        'Incrementamos intervalo de siguiente hora
        nextDateTime = sumaHoraMinutos(nextDateTime, intervalo)
'        nextHour = sumaHoraMinutos(nextHour, intervalo)
        
        'calculamos siguiente hora
        'si esta en el intervalo, incrementamos el dia y ponemos la hora a horarioDesde
        'sino significa que es el mismo dia, no importa actualizar nada
        If Not ((nextDateTime >= horarioFrom) And (nextDateTime <= horarioTo)) Then
            'Si la siguiente hora es mayor que la fechaFin incrementamos dia
            If (nextDateTime > horarioTo) Then
                nextDateTime = sumaFechaDias(nextDateTime, 1)
            End If
            nextDateTime = CDate(Format(nextDateTime, "dd/mm/yyyy") & " " & Format(horarioFrom, "hh:mm:ss"))
            'nextHour = horarioFrom
        End If
        
        'Actualizamos variables a retornar
        fechaHora = nextDateTime
'       fecha = nextDate
'       hora = nextHour
    Else
        MsgBox "Error: intervalo demasiado grande(max.23horas)", vbOKOnly, "SMS Module"
    End If
End Sub

Public Function calculaEnvios(ByRef numEnvios As Integer, _
                              ByRef tamanyoUltimoEnvio As Integer, _
                              numeroDeMoviles As Integer, _
                              movilesPorPaquete As Integer)
    
    Dim envios As Integer
    Dim numUltimoEnvio As Integer
    Dim resto As Integer
    
    envios = numeroDeMoviles \ movilesPorPaquete
    resto = numeroDeMoviles Mod movilesPorPaquete
    
    If (resto > 0) Then
        numUltimoEnvio = resto
        envios = envios + 1
    Else 'si resto es 0, ultimo = movilesPorPaquete
        numUltimoEnvio = movilesPorPaquete
    End If
    
'debugando "numenvios: " & envios & " |tamanyo ultimo: " & numUltimoEnvio

    numEnvios = envios
    tamanyoUltimoEnvio = numUltimoEnvio
    
End Function

Public Function pruebaSMS()
    Dim itmp As Integer
    Dim itmp1 As Integer
    Dim stmp As String
    
    Dim a() As String
    ReDim a(1 To 5)
    
    Debug.Print UBound(a)
End Function

