Attribute VB_Name = "LigaFormulario"
Option Compare Database
Option Explicit

Public U_colFormularios As New Collection


'Author: Jose M. Huerta Guillén
'Date: 03/11/09
'Date update: 03/11/09
'Name: LigaFormulario
'Descr: Añade una referencia de un formulario a una colección global para evitar
'       que el formulario se cierre al salir del ámbito del código que lo ha creado.
'Param: frm, puntero al formulario que se desea ligar.Public Function LigaFormulario(frm As Object)

Public Function LigaFormulario(frm As Object)
    U_colFormularios.Add frm
End Function
