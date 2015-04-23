Attribute VB_Name = "G_Delay"
Option Explicit
Option Compare Database

Private Declare Sub AppSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
 
Public Sub PauseApp(PauseInSeconds As Long)
    
    Call AppSleep(PauseInSeconds * 1000)
    
End Sub
