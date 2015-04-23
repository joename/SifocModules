Attribute VB_Name = "LigaFormulario"
Option Compare Database
Option Explicit

Public U_colFormularios As New Collection


'Author: Jose M. Huerta Guill�n
'Date: 03/11/09
'Date update: 03/11/09
'Name: LigaFormulario
'Descr: A�ade una referencia de un formulario a una colecci�n global para evitar
'       que el formulario se cierre al salir del �mbito del c�digo que lo ha creado.
'Param: frm, puntero al formulario que se desea ligar.Public Function LigaFormulario(frm As Object)

Public Function LigaFormulario(frm As Object)
    U_colFormularios.Add frm
End Function
