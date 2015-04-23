Attribute VB_Name = "G_Message"
Option Explicit
Option Compare Database

Public Function MessageBoxTimer(title As String, msg As String, sec As Integer)
    Dim AckTime As Integer, InfoBox As Object
    Set InfoBox = CreateObject("WScript.Shell")
    'Set the message box to close after 10 seconds
    AckTime = sec
    Select Case InfoBox.PopUp(msg, AckTime, title, 0)
        Case 1, -1
            Exit Function
    End Select
End Function
