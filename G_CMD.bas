Attribute VB_Name = "G_CMD"
Option Explicit
Option Compare Database

Public Function command(Ruta As String)
    Dim RetVal
    RetVal = Shell(Ruta, 1)
End Function

Public Function swriter(Ruta As String)
    Dim RetVal
    RetVal = Shell("cmd /c start swriter " & Ruta, vbNormalFocus)
End Function

