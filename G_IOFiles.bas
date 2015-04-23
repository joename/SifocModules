Attribute VB_Name = "G_IOFiles"
Option Explicit
Option Compare Database

Public Function createFile(txt As String)
    Dim fs As Object
    Dim a As Object
    Dim Ruta As String

    Ruta = CurrentProject.path
    ChDir (Ruta)
    Ruta = Ruta & "\ArchivoEnvioOfertas.xml"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(Ruta, True, True)
a.write ("hola")
    a.write (txt)
a.write ("adios")
    a.Close
    
End Function
