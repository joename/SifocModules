Attribute VB_Name = "G_XML"
Option Explicit
Option Compare Database

Public Function ExportXML()
    Dim objOrderInfo As AdditionalData
    Dim objOrderDetailsInfo As AdditionalData
    
    Set objOrderInfo = Application.CreateAdditionalData
    
    ' Add the Orders and Order Details tables to the data to be exported.
    Set objOrderDetailsInfo = objOrderInfo.Add("a_servicio")
    objOrderDetailsInfo.Add "id"
    
    ' Export the contents of the Customers table. The Orders and Order
    ' Details tables will be included in the XML file.
    Application.ExportXML ObjectType:=acExportTable, DataSource:="a_accion", _
                          DataTarget:="c:/a_sevicio.xml", _
                          AdditionalData:=objOrderInfo
End Function

