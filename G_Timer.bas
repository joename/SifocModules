Attribute VB_Name = "G_Timer"
Option Explicit
Option Compare Database

Public Static Function wait(sec As Integer)
    Dim now As Long
    
    now = Timer()
    
    Do
        DoEvents
    Loop While (Timer < now + sec)

End Function
