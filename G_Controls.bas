Attribute VB_Name = "G_Controls"
Option Explicit
Option Compare Database

Public Function selTodos(ctl As control)
    Dim i As Integer
    For i = 0 To ctl.ListCount - 1
        ctl.Selected(i) = True
    Next i
End Function

Public Function deselTodos(ctl As control)
    Dim i As Integer
    For i = 0 To ctl.ListCount - 1
        ctl.Selected(i) = False
    Next i
End Function
