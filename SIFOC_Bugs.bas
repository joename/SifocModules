Attribute VB_Name = "SIFOC_Bugs"
Option Explicit
Option Compare Database

Public Function countBugs(formulario As String) As Long
    countBugs = DCount("[id]", "sysIncidencias", "[form]='" & formulario & "' AND [listo]=0")
End Function
