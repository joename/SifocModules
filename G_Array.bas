Attribute VB_Name = "G_Array"
Option Explicit
Option Compare Database

Public Function IsInitialized(arr() As String) As Boolean
    On Error GoTo ErrHandler
    Dim nUbound As Long
    nUbound = UBound(arr)
    IsInitialized = True
    Exit Function
ErrHandler:
    Exit Function
End Function

Public Function aaa() As Boolean
    Dim arr() As String
    
    aaa = IsInitialized(arr)
End Function
