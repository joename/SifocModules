Attribute VB_Name = "SIFOC_Seguridad"
Option Explicit
Option Compare Database

'------------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  04/12/2009 - Actualizacion: 04/12/2009
'   Name:   captcha
'   Desc:   Verifica que se introdujo correctamente el captcha
'   Param:  -
'   Retur:  boolean, true si es correcto
'                    false si es incorrecto
'-------------------------------------------------------------------------------------
Public Function esCaptchaCorrecto() As Boolean
    Dim num As Integer
    Dim str As String
    Dim strRespuesta As String
    
    Randomize
    num = Int(Rnd() * 10)
    str = "Como medida de seguridad y para asegurarnos de que ha leído este mensaje," & _
           "debe escribir el siguiente número para borrar la gestión." & vbNewLine & _
           "Numero: " & num
    
    strRespuesta = InputBox(str, "Medida de seguridad")
    If Int(strRespuesta) = num Then
        esCaptchaCorrecto = True
    Else
        esCaptchaCorrecto = False
    End If
End Function
