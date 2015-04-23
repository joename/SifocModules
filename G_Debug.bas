Attribute VB_Name = "G_Debug"
Option Explicit
Option Compare Database

'Activa debug en todos los modulos
Public Const U_Debug = False

'---------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  21/1/2009
'   Name:   debugando
'   Desc:   Muestra los comentarios de debug cuando U_Debug esta a TRUE
'   Param:  comentario a mostrar por la ventana de inmediato
'   Retur:  -
'---------------------------------------------------------------------------
Public Function debugando(comentario)
    If (U_Debug = True) Then
        Debug.Print comentario
    End If
End Function

