Attribute VB_Name = "G_APIWindows"
Option Explicit
Option Compare Database

'API para obtener el nombre del ordenador actual
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long
    
Public Const MAX_COMPUTERNAME_LENGTH = 255

'API para obtener el usuario actual
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

'API para obtener la version de windows

Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Private Type so
    soname As String
    num As Integer
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFOEX) As Long

'El siguiente Tipo definido por el usuario, es para utilizar la propia API
'de Access (Realmente del ejecutable MSACCESS.EXE para abrir ficheros.
'Es la primera variedad de poder escoger ficheros. Expuesto en las News
'de Microsoft Access por Juan M. Afán de Ribera.
'Ojo el tipo de datos definido por el usuario «Type», no existe en Access 97

Type OFFICEGETFILENAMEINFO
    hwndOwner As Long
    szAppName As String * 255
    szDlgTitle As String * 255
    szOpenTitle As String * 255
    szFile As String * 4096
    szInitialDir As String * 255
    szFilter As String * 255
    nFilterIndex As Long
    lView As Long
    flags As Long
End Type
Declare Function GetFileName _
        Lib "msaccess.exe" _
        Alias "#56" _
        (gfni As OFFICEGETFILENAMEINFO, _
        ByVal fOpen As Integer) As Long

'------------------------------------------------------------------------------
'       Funciones para saber el nombre de equipo
'------------------------------------------------------------------------------
'Esta funcion devuelve el nombre de pc
Public Function computerName() As String
    'Devuelve el nombre del equipo actual
    Dim sComputerName As String
    Dim ComputerNameLength As Long

    sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    ComputerNameLength = MAX_COMPUTERNAME_LENGTH
    Call GetComputerName(sComputerName, ComputerNameLength)
    computerName = Mid(sComputerName, 1, ComputerNameLength)
End Function

'------------------------------------------------------------------------------
'       Funciones para saber el usuario de windows
'------------------------------------------------------------------------------
'Esta función devuelve el nombre del Usuario
Public Function userName() As String
    Dim sBuffer As String
    Dim lSize As Long
    Dim sUser As String

    sBuffer = Space$(260)
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then
        sUser = Left$(sBuffer, lSize)
        'Quitarle el CHR$(0) del final...
        lSize = InStr(sUser, Chr$(0))
        If lSize Then
            sUser = Left$(sUser, lSize - 1)
        End If
    Else
        sUser = ""
    End If
    userName = sUser
End Function

'------------------------------------------------------------------------------
'       Funciones para saber el usuario de windows
'------------------------------------------------------------------------------
Private Function szTrim(ByVal s As String) As String
    ' Quita los caracteres en blanco y los Chr$(0)
    Dim i As Long

    i = InStr(s, vbNullChar)
    If i Then
        s = Left$(s, i - 1)
    End If
    s = Trim$(s)
    
    szTrim = s
End Function

Public Function WindowsVersion() As String
    Dim OSInfo As OSVERSIONINFOEX
    Dim sistema As so
    Dim ret As Long
    Dim s As String
    
    Dim mayor(3 To 6) As so
    Dim minor(0 To 90) As so
    
    'Cargamos sistemas operativos Mayor
    sistema.soname = "Windows NT 3.51"
    sistema.num = 3
    mayor(3) = sistema
    sistema.soname = "Windows 95, 98, Me y NT 4.0"
    sistema.num = 4
    mayor(4) = sistema
    sistema.soname = "Windows 2000, XP y 2003"
    sistema.num = 5
    mayor(5) = sistema
    sistema.soname = "Windows Vista/Longhorn"
    sistema.num = 6
    mayor(6) = sistema

    'Cargamos sistemos operativos Minor
    sistema.soname = "Windows 95, NT 4.0, 2000, Vista/Longhorn"
    sistema.num = 0
    minor(0) = sistema
    sistema.soname = "Windows XP"
    sistema.num = 1
    minor(1) = sistema
    sistema.soname = "Windows 2003"
    sistema.num = 2
    minor(2) = sistema
    sistema.soname = "Windows 98"
    sistema.num = 10
    minor(3) = sistema
    sistema.soname = "Windows NT 3.51"
    sistema.num = 51
    minor(51) = sistema
    sistema.soname = "Windows Me"
    sistema.num = 90
    minor(90) = sistema
    
    OSInfo.szCSDVersion = Space$(128)
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    ret = GetVersionEx(OSInfo)
    
    s = "MajorVersion     " & OSInfo.dwMajorVersion & vbCrLf & _
        "MinorVersion     " & OSInfo.dwMinorVersion & vbCrLf & _
        "BuildNumber      " & OSInfo.dwBuildNumber & vbCrLf & _
        "PlatformId       " & OSInfo.dwPlatformId & vbCrLf & _
        "CSDVersion       " & szTrim(OSInfo.szCSDVersion)

    'debugando minor(OSInfo.dwMinorVersion).num & "|" & minor(OSInfo.dwMinorVersion).soname
    'debugando mayor(OSInfo.dwMajorVersion).num & "|" & mayor(OSInfo.dwMajorVersion).soname
    
    WindowsVersion = minor(OSInfo.dwMinorVersion).soname & " " & szTrim(OSInfo.szCSDVersion)
End Function

'----------------------------------------------------------------------------------
'   Open File Dialog
'----------------------------------------------------------------------------------
'***************************************************************************
' Utilizando API de Access
'***************************************************************************
Function EscogeFichero_Con_ApiAccess() As String
 Dim FileInfo As OFFICEGETFILENAMEINFO

   With FileInfo
      .hwndOwner = access.hWndAccessApp
      .szFilter = "Ficheros de texto (*.txt)" + Chr$(0) + "*.txt" + Chr$(0)
      .szInitialDir = "C:\"
      .szDlgTitle = "hola" '"Escoja fichero de texto a cargar"
      .szOpenTitle = "hola1" '"Escojer ficheros con la API de Access"
   End With
   GetFileName FileInfo, True
   If FileInfo.szFile <> "" Then
    EscogeFichero_Con_ApiAccess = Trim(FileInfo.szFile)
   Else
     EscogeFichero_Con_ApiAccess = ""
   End If
  
End Function
'**************************************************************************************
'Aqui se termina el primer método de escoger fichero, utilizando API del propio Access
'**************************************************************************************

'**************************************************************************************
'método: Utilizando el objeto Indocumentado Wizhook de Access:
'Seguramente, la más ingeniosa, gracias a la labor de investigación de Juan
'sobre el objeto indocumentado WizHook
'**************************************************************************************
'***************************************************************************
' U T I L I Z A N D O    WIZHOOK
'***************************************************************************
 Public Function EscogeFichero_Con_WizHook() As String
  '© Juan M. Afán de Ribera
  Dim wzhwndOwner As Long
  Dim wzAppName As String
  Dim wzDlgTitle As String
  Dim wzOpenTitle As String
  Dim wzFile As String
  Dim wzInitialDir As String
  Dim wzFilter As String
  Dim wzFilterIndex As Long
  Dim wzView As Long
  Dim wzflags As Long
  Dim wzfOpen As Boolean
  Dim ret As Long
    
    WizHook.Key = 51488399
    
    wzhwndOwner = 0&
    wzAppName = ""
    wzDlgTitle = "Cuadro de diálogo con WizHook"
    wzOpenTitle = "Abrir con Wz"
    wzFile = String(255, Chr(0))
    wzInitialDir = "C:\"
    wzFilter = "Archivos gráficos " _
    & "(*.txt)"
    wzFilterIndex = 1
    wzView = 1
    wzflags = 64
    wzfOpen = True

    ret = WizHook.GetFileName(wzhwndOwner, _
        wzAppName, wzDlgTitle, wzOpenTitle, wzFile, _
        wzInitialDir, wzFilter, wzFilterIndex, _
        wzView, wzflags, wzfOpen)
        
    ' Si no se ha pulsado el botón Cancelar (-302)
    If ret <> -302 Then
     EscogeFichero_Con_WizHook = wzFile
    Else
     EscogeFichero_Con_WizHook = ""
    End If
 End Function
'*****************************************************************************************
'Aqui se termina el tercer método de escoger fichero, utilizando WizHook del propio Access
'Si quieres saber más sobre este objeto indocumetado, visita la Web de Juan M. Afan
'http://www.juanmafan.tk/
'******************************************************************************************

