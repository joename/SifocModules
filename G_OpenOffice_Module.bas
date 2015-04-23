Attribute VB_Name = "G_OpenOffice_Module"
Option Explicit
Option Compare Database

'--------------------------------------------------------------------------
'   Name: openDoc
'   Desc: Load an existing writer document, with opening parameters
'   Parm:
'--------------------------------------------------------------------------
Public Sub openDoc()
      Dim oSM, oDesk As Object 'root object from OOo API
      Dim oDoc As Object       'The document to be opened
      Dim OpenPar(2) As Object 'a Visual Basic array, with 3 elements
    
    'Instanciate OOo : the first line is always required from Visual Basic for OOo
      Set oSM = CreateObject("com.sun.star.ServiceManager")
      Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
    
    'We call the MakePropertyValue function, defined just before, to access the structure
      Set OpenPar(0) = MakePropertyValue("ReadOnly", True)
      Set OpenPar(1) = MakePropertyValue("Password", "secret")
      Set OpenPar(2) = MakePropertyValue("Hidden", False)
    
    'Now we can call the OOo loadComponentFromURL method, giving it as
    'fourth argument the result of our precedent MakePropertyValue call
      Set oDoc = oDesk.loadComponentFromURL("file:///c|test.sxw", "_blank", 0, OpenPar)
End Sub

'--------------------------------------------------------------------------
'   Name: ConvertToUrl
'   Desc: Converts a Ms Windows local pathname in URL (RFC 1738)
'   Parm: UNC pathnames, more character conversions
'--------------------------------------------------------------------------
Public Function Convert2Url(strFile) As String
    strFile = Replace(strFile, "\", "/")
    strFile = Replace(strFile, ":", "|")
    strFile = Replace(strFile, " ", "%20")
    strFile = "file:///" + strFile
    Convert2Url = strFile
End Function

'--------------------------------------------------------------------------
'   Name: MakePropertyValue
'   Desc: Creates a sequence of com.sun.star.beans.PropertyValue s
'   Parm:
'--------------------------------------------------------------------------
Public Function MakePropertyValue(cName, uValue) As Object
Dim oStruct, oServiceManager As Object
    Set oServiceManager = CreateObject("com.sun.star.ServiceManager")
    Set oStruct = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
    oStruct.name = cName
    oStruct.value = uValue
    Set MakePropertyValue = oStruct
End Function

'--------------------------------------------------------------------------
'   Name: CreateUnoService
'   Desc: A simple shortcut to create a service
'   Parm:
'--------------------------------------------------------------------------
Public Function CreateUnoService(strServiceName) As Object
Dim oServiceManager As Object
    Set oServiceManager = CreateObject("com.sun.star.ServiceManager")
    Set CreateUnoService = oServiceManager.createInstance(strServiceName)
End Function

'--------------------------------------------------------------------------
'   Name: FileName2URL
'   Desc: Convert file name to an URL name
'   Parm:
'--------------------------------------------------------------------------
Public Function FileName2URL(ByVal pFileName As String, _
                             Optional ByVal pConvertBackslashesToSlashes As Boolean = True) _
                             As String
  Dim s As String
  Dim z As String
  Dim j As Long
  Dim x As Integer
  
  On Error Resume Next
  s = ""
  For j = 1 To Len(pFileName)
    z = Mid(pFileName, j, 1)
    x = Asc(z)
    Select Case x
      Case 9
        z = "%09"
      Case 13
        z = "%0d"
      Case 10
        z = "%0a"
      'Case 32
      '  z = "+"
      Case 32 To 35, 37 To 41, 43, 44, 59 To 63, 91, 93, 94, 96, Is >= 123
        z = "%" & Hex(x)
      Case 92
        If pConvertBackslashesToSlashes Then z = "/"
    End Select
    s = s & z
  Next
  s = "file:///" & s
  FileName2URL = s
End Function

'--------------------------------------------------------------------------
'   Name: SaveAsPDF
'   Desc: Save an existing writer document as PDF
'   Parm:
'--------------------------------------------------------------------------
Public Sub SaveAsPDF()
'
' Save an existing writer document as PDF
'
    Dim oSM, oDesk, oDoc As Object 'OOo objects
    Dim OpenParam(1) As Object 'Parameters to open the doc
    Dim SaveParam(1) As Object 'Parameters to save the doc
    
    Set oSM = CreateObject("com.sun.star.ServiceManager")
    Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
    
    Set OpenParam(0) = MakePropertyValue("Hidden", True)  ' Open the file hidden
    Set oDoc = oDesk.loadComponentFromURL("file:///C:/tmp/testdoc.odt", "_blank", 0, OpenParam())
    
    Set SaveParam(0) = MakePropertyValue("FilterName", "writer_pdf_Export")
    Call oDoc.storeToURL("file:///C:/tmp/testdoc.pdf", SaveParam())
    
    Set oDesk = Nothing
    Set oSM = Nothing

End Sub
