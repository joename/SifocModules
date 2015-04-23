Attribute VB_Name = "G_Mail_Module"
Option Explicit
Option Compare Database

Public Const maxMessageLong As Integer = 160
Public Const maxTitleLong As Integer = 11

'Public Const separadorTelefonos As String = ";"

'---------------------------------------------------------------------------
'   Autor:  ¿?
'   Fecha:  24/02/2010 - Actualización:  24/02/2010
'   Name:   esEmail
'   Desc:   Función que comprueba si una diección de email es válida
'   Param:  email(string)
'   Retur:  Verdadero, si es correcto
'           Falso, si es incorrecto
'---------------------------------------------------------------------------
Public Function esEmail(ByVal email As String) As Boolean
    
    Dim i As Integer, iLen As Integer, caracter As String
    Dim pos As Integer, bp As Boolean, iPos As Integer, iPos2 As Integer

On Local Error GoTo Err_Sub

    email = Trim$(email)

    If email = vbNullString Then
        Exit Function
    End If

    email = LCase$(email)
    iLen = Len(email)

    
    For i = 1 To iLen
        caracter = Mid(email, i, 1)
        If (Not (caracter Like "[a-z]")) _
            And (Not (caracter Like "[0-9]")) _
            And (Not caracter Like "-") _
            And (Not caracter Like "_") Then
            
            If InStr(1, "_-" & "." & "@", caracter) > 0 Then
                If bp = True Then
                   Exit Function
                Else
                    bp = True
                   
                    If i = 1 Or i = iLen Then
                        Exit Function
                    End If
                    If caracter = "@" Then
                        If iPos = 0 Then
                            iPos = i
                        Else
                            Exit Function
                        End If
                    End If
                    If caracter = "." Then
                        iPos2 = i
                    End If
                End If
            Else
                Exit Function
            End If
        Else
            bp = False
        End If
    Next i
    
    If iPos = 0 Or iPos2 = 0 Then
        Exit Function
    End If
    
    If iPos2 < iPos Then
        Exit Function
    End If
    
    esEmail = True

    Exit Function

Err_Sub:
    On Local Error Resume Next
    esEmail = False
End Function

Function sendMail(texto As String, from As String)
 
    Dim iMsg 'Create the message object
    Set iMsg = CreateObject("CDO.Message")
    Dim iConf 'Create the configuration object
    Set iConf = CreateObject("CDO.Configuration")
    Dim Flds 'Set the fields of the configuration object to send using exchange server
    Set Flds = iConf.fields
    Flds("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    'SMTP SERVER
    Flds("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.ifoc.es"
    Flds("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    
    
    ' Setup server login information if your server require it
    ' 0- anonimos !!!! EL MAS NORMAL
    ' 1 -Use basic (clear-text) authentication
    ' 2- Use NTLM authentication (Secure Password Authentication in Microsoft® Outlook® Express).
        
     
    Flds("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    'NO RELLENAR NADA SI ES ANONIMO
    Flds("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@ifoc.es"
    Flds("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "correoifoc"
    Flds.update 'Set the message to,from,subject,body properties.
        
    Set iMsg.Configuration = iConf
    iMsg.To = "jmsanchez@calvia.com"
    iMsg.from = from & "@ifoc.es"
    iMsg.Subject = "Turnos"
    iMsg.TextBody = texto
    iMsg.Send
    If Err <> 0 Then
       sendMail = -1
    Else
       sendMail = 0
    End If
  
End Function


Private Sub test_mail_Click()

'CUERPO = ""
'vbf = Chr(10) + Chr(13)
'hay = rs1.RecordCount
'If hay > 0 Then
'    rs1.MoveFirst
'    For topmargen = 0 To 5
'        CUERPO = CUERPO + vbf
'    Next topmargen
'    TIENDA = rs1.fields(0).value
'    CUERPO = CUERPO & " Listado Turnos Tienda " & TIENDA
'
'    For topmargen = 0 To 2
'        CUERPO = CUERPO + vbf
'    Next topmargen
'
'     For X = 1 To hay
'       CUERPO = CUERPO & " " & PADR(Left(rs1.fields(1).value, 15), 15, " ") & Chr$(9) & PADR(Left(rs1.fields(2).value, 30), 30, " ") & Chr$(9) & Left(rs1.fields(3).value, 5) & Chr$(9) & Left(rs1.fields(4).value, 5) & Chr$(9) & Left(rs1.fields(5).value, 5) & Chr$(9) & Left(rs1.fields(6).value, 5) & vbf
'       debugandoCUERPO
'       rs1.MoveNext
'    Next X
'  'MsgBox (cuerpo)
'  res = SENDMAIL(CUERPO, TIENDA)
'  If res = 0 Then
'    MsgBox ("OK Email Enviado  ")
'  Else
'    MsgBox ("Problemas envio email")
'  End If
'
'End If

End Sub

'*****Enviar mail con outlook**************************************
Private Sub CmdEnv_Click()

   'Dim oOutlook As Outlook.Application
   'Dim oFolder As Outlook.MAPIFolder
   'Dim oItem As Outlook.MailItem
   'Dim oAttach As Outlook.Attachment
   'Dim myRecipients As Outlook.Recipients
   'Dim myRecipient As Outlook.Recipient
   'Dim cEmail As String, cAsunto As String, cMensaje
   'Dim cFicOri As String, cFicDes As String

  

   ' Dirección del Correo
   'cEmail = "Pepito@TodoExPertos.com"
  
   ' Asunto del Correoc
   'Asunto = "Asunto del mensaje"
  
   ' Cuerpo del Mensaje
   'cMensaje = "Lo que quieras poner." & vbCrLf & vbCrLf & vbCrLf & "Un saludo,"
  
   'Ficheros Origen y Destino
   'cFicOri = "C:\Documents and Settings\USUARIO\Escritorio\Archivo.mdb"
   'cFicDes = "C:\Documents and Settings\USUARIO\Escritorio\Archivo.XXX"
  
   ' Copiar el Archivo
   'FileCopy cFicOri, cFicDes
  
   'Set oOutlook = Outlook.Application
   'Set oFolder = oOutlook.GetNamespace("MAPI").GetDefaultFolder(4)
   'Set oItem = oFolder.items.Add("IPM.Note")
  
   'oItem.To = cEmail
   'oItem.Subject = cAsunto
   'oItem.Body = cMensaje
   'oItem.DeleteAfterSubmit = False ' Guarda una copia del mensaje después de enviarlo
  
   ' Datos Adjuntos
   'Set oAttach = oItem.Attachments.Add(cFicDes, olByValue, 1, "Fichero")
  
   ' Envía el Mensaje directamente (sin abrir Outlook)
   'oItem.Send
  
   ' Muestra el Mensaje (abre el mensaje, el usuario tiene que enviarlo)
   'oItem.Display
  
End Sub
'*******************************************
