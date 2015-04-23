Attribute VB_Name = "z_pruebas"
Option Explicit
Option Compare Database

Function FindFiles(path As String, SearchStr As String, _
       FileCount As Integer, DirCount As Integer)
      Dim FileName As String   ' Walking filename variable.
      Dim DirName As String    ' SubDirectory Name.
      Dim dirNames() As String ' Buffer for directory name entries.
      Dim nDir As Integer      ' Number of directories in this path.
      Dim i As Integer         ' For-loop counter.

      On Error GoTo sysFileERR
      If Right(path, 1) <> "\" Then path = path & "\"
      ' Search for subdirectories.
      nDir = 0
      ReDim dirNames(nDir)
      DirName = Dir(path, vbDirectory Or vbHidden Or vbArchive Or vbReadOnly _
Or vbSystem)  ' Even if hidden, and so on.
      Do While Len(DirName) > 0
         ' Ignore the current and encompassing directories.
         If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetAttr(path & DirName) And vbDirectory Then
               dirNames(nDir) = DirName
               DirCount = DirCount + 1
               nDir = nDir + 1
               ReDim Preserve dirNames(nDir)
               'List2.AddItem path & DirName ' Uncomment to list
            End If                           ' directories.
sysFileERRCont:
         End If
         DirName = Dir()  ' Get next subdirectory.
      Loop

      ' Search through this directory and sum file sizes.
      FileName = Dir(path & SearchStr, vbNormal Or vbHidden Or vbSystem _
      Or vbReadOnly Or vbArchive)
      While Len(FileName) <> 0
         FindFiles = FindFiles + FileLen(path & FileName)
         FileCount = FileCount + 1
         ' Load List box
         'List2.AddItem path & FileName & vbTab & _
            FileDateTime(path & FileName)   ' Include Modified Date
         FileName = Dir()  ' Get next file.
      Wend

      ' If there are sub-directories..
      If nDir > 0 Then
         ' Recursively walk into them
         For i = 0 To nDir - 1
           FindFiles = FindFiles + FindFiles(path & dirNames(i) & "\", _
            SearchStr, FileCount, DirCount)
         Next i
      End If

AbortFunction:
      Exit Function
sysFileERR:
      If Right(DirName, 4) = ".sys" Then
        Resume sysFileERRCont ' Known issue with pagefile.sys
      Else
        MsgBox "Error: " & Err.Number & " - " & Err.description, , _
         "Unexpected Error"
        Resume AbortFunction
      End If
End Function

Public Function holaa() As Boolean
    DoCmd.SetWarnings True
End Function

Public Function a1(str As String) As String
    Dim str1 As String
    
    If (Len(str) = 8) Then
        str1 = Left(str, 4) & "." & Mid(str, 5, 3) & "." & Right(str, 1)
    Else
        str1 = str
    End If
    a1 = str1
End Function

Public Function a2(idCno94 As Integer) As Integer
    Dim idCno11 As Integer
    Dim strCno94 As String
    
    strCno94 = Nz(DLookup("[codigo]", "a_cno94", "[id]=" & idCno94), 0)
    idCno11 = Nz(DLookup("[id]", "a_cno2011", "[cno94]='" & strCno94 & "'"), 0)
    
    a2 = idCno11
End Function


Public Static Function initSql() As String

End Function
