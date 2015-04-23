Attribute VB_Name = "SIFOC_Menu"
Option Explicit
Option Compare Database

'------------------------------------------------------------------------------
'   Autor:  Jose Manuel Sanchez
'   Fecha:  15/10/2009
'   Name:   openObject
'   Desc:   Abre el objeto indicado en el menu1(formulario, consulta, informe)
'   Param:
'           tipo(string), tipo de objeto a abrir
'           agrumento(string), nombre del objeto a abrir
'           dataMode(string), modo de apertura en caso de form
'   Retur:  devuelve dniNie correcto NNNNNNNL
'------------------------------------------------------------------------------
Public Function openObject(tipo As String, _
                           argumento As String, _
                           Optional dataMode As String = "", _
                           Optional args As String = "")
    Select Case tipo
        Case "form"
            If (dataMode = "acFormEdit") Then
                If (args = "") Then
                    DoCmd.openForm argumento, acNormal, , , acFormEdit, acWindowNormal
                Else
                    DoCmd.openForm argumento, acNormal, , , acFormEdit, acWindowNormal, args
                End If
            ElseIf (dataMode = "acFormAdd") Then
                If (args = "") Then
                    DoCmd.openForm argumento, acNormal, , , acFormAdd, acWindowNormal
                Else
                    DoCmd.openForm argumento, acNormal, , , acFormAdd, acWindowNormal, args
                End If
            ElseIf (dataMode = "acFormPropertySettings") Then
                If (args = "") Then
                    DoCmd.openForm argumento, acNormal, , , acFormPropertySettings, acWindowNormal
                Else
                    DoCmd.openForm argumento, acNormal, , , acFormPropertySettings, acWindowNormal, args
                End If
            ElseIf (dataMode = "acFormReadOnly") Then
                If (args = "") Then
                    DoCmd.openForm argumento, acNormal, , , acFormReadOnly, acWindowNormal
                Else
                    DoCmd.openForm argumento, acNormal, , , acFormReadOnly, acWindowNormal, args
                End If
            End If
        Case "query"
            hitQuery argumento, U_idIfocUsuarioActivo
            DoCmd.OpenQuery argumento, acViewNormal, acReadOnly
        Case "report"
            DoCmd.OpenReport argumento, acViewPreview, , , acWindowNormal
    End Select
End Function

